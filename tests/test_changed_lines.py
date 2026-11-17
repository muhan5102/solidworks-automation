"""用真实Git提交验证事件范围和强推后缺失基准的获取，不连接外网。"""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


CHECKER = Path(__file__).resolve().parents[1] / 'tools/check_changed_lines.py'


class ChangedLinesTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='solidworks-changed-lines-')
        self.root = Path(self.temporary.name).resolve()
        self.assertEqual(self.root.parent, Path(tempfile.gettempdir()).resolve())
        self.addCleanup(self.temporary.cleanup)
        self.repo = self.root / 'origin'
        self.repo.mkdir()
        self.git('init', '-q')
        self.write('base.txt', b'base\n')
        self.git('add', '.')
        self.git('commit', '-qm', 'fixture baseline')
        self.base = self.git('rev-parse', 'HEAD')

    def git(self, *args, cwd=None):
        prefix = ['git', '-c', 'commit.gpgsign=false', '-c', 'user.name=ci-fixture',
                  '-c', 'user.email=ci@example.com', '-c', 'core.hooksPath=', '-c', 'core.autocrlf=false']
        result = subprocess.run(prefix + list(args), cwd=cwd or self.repo,
                                capture_output=True, text=True, encoding='utf-8', check=False)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        return result.stdout.strip()

    def write(self, name, data):
        path = (self.repo / name).resolve()
        self.assertTrue(path.is_relative_to(self.repo))
        path.write_bytes(data)

    def check(self, event, *, base='', before='', cwd=None):
        env = os.environ.copy()
        env.update(GITHUB_EVENT_NAME=event, PR_BASE_SHA=base, PUSH_BEFORE_SHA=before)
        return subprocess.run([sys.executable, '-X', 'utf8', str(CHECKER)], cwd=cwd or self.repo,
                              env=env, capture_output=True, text=True, encoding='utf-8', check=False)

    def test_manual_and_first_push_accept_clean_commit(self):
        for event, before in [('workflow_dispatch', ''), ('push', '0' * 40)]:
            with self.subTest(event=event):
                result = self.check(event, before=before)
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_pr_and_push_detect_earlier_bad_commit(self):
        self.write('bad.txt', b'bad whitespace \n')
        self.git('add', '.')
        self.git('commit', '-qm', 'fixture whitespace')
        self.write('later.txt', b'clean later change\n')
        self.git('add', '.')
        self.git('commit', '-qm', 'fixture later clean change')
        self.assertEqual(self.check('workflow_dispatch').returncode, 0)
        for event in ('pull_request', 'push'):
            with self.subTest(event=event):
                result = self.check(event, base=self.base, before=self.base)
                self.assertNotEqual(result.returncode, 0)
                self.assertIn('trailing whitespace', result.stdout + result.stderr)

    def test_force_push_fetches_missing_base(self):
        self.git('checkout', '-qb', 'old')
        self.write('old.txt', b'old branch\n')
        self.git('add', '.')
        self.git('commit', '-qm', 'fixture old merge')
        old = self.git('rev-parse', 'HEAD')
        self.git('checkout', '-qb', 'replacement', self.base)
        self.write('new.txt', b'replacement\n')
        self.git('add', '.')
        self.git('commit', '-qm', 'fixture replacement')
        clone = self.root / 'clone'
        self.git('clone', '-q', '--single-branch', '--no-local', '--branch', 'replacement', str(self.repo), str(clone))
        missing = subprocess.run(['git', 'cat-file', '-e', old], cwd=clone, capture_output=True, check=False)
        self.assertNotEqual(missing.returncode, 0)
        result = self.check('push', before=old, cwd=clone)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.git('cat-file', '-e', old, cwd=clone)

    def test_unavailable_base_fails_instead_of_skipping(self):
        result = self.check('push', before='1' * 40)
        self.assertNotEqual(result.returncode, 0)

    def test_pr_requires_base(self):
        self.assertNotEqual(self.check('pull_request').returncode, 0)


if __name__ == '__main__':
    unittest.main()
