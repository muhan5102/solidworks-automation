"""验证仓库检查能拦住实际错误；样本脚本只做静态分析。"""
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


CHECKER = Path(__file__).resolve().parents[1] / 'tools/check_repository.py'
CALLER = '''from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from shared.cad import helper
'''


class RepositoryCheckTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='solidworks-repository-check-')
        self.root = Path(self.temporary.name).resolve()
        self.assertEqual(self.root.parent, Path(tempfile.gettempdir()).resolve())
        self.addCleanup(self.temporary.cleanup)
        for directory in ('tools', 'practice/shared', 'practice/source/case'):
            (self.root / directory).mkdir(parents=True)
        (self.root / 'tools/check_repository.py').write_bytes(CHECKER.read_bytes())
        self.write('practice/shared/cad.py', 'def helper():\n    return True\n')
        self.caller = 'practice/source/case/build.py'
        self.write(self.caller, CALLER)
        self.write('README.md', '[案例](practice/source/case/build.py)\n')

    def write(self, relative, text):
        path = (self.root / relative).resolve()
        self.assertTrue(path.is_relative_to(self.root))
        path.write_text(text, encoding='utf-8', newline='\n')

    def run_checker(self):
        return subprocess.run(
            [sys.executable, '-X', 'utf8', 'tools/check_repository.py'],
            cwd=self.root, capture_output=True, text=True, encoding='utf-8', check=False,
        )

    def assert_passes(self):
        result = self.run_checker()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def assert_rejects_then_recovers(self, relative, invalid, restored, message):
        self.write(relative, invalid)
        result = self.run_checker()
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn(message, result.stdout + result.stderr)
        self.write(relative, restored)
        self.assert_passes()

    def test_clean_checkout_needs_no_cad_or_outputs(self):
        self.assertFalse((self.root / 'outputs').exists())
        self.assert_passes()

    def test_rejects_python_syntax_error(self):
        self.assert_rejects_then_recovers(self.caller, 'def broken(:\n', CALLER, 'invalid syntax')

    def test_rejects_missing_shared_symbol(self):
        self.assert_rejects_then_recovers(
            self.caller, CALLER.replace('import helper', 'import missing_helper'), CALLER,
            '缺少shared.cad.missing_helper',
        )

    def test_rejects_missing_shared_module(self):
        self.assert_rejects_then_recovers(
            self.caller, CALLER.replace('shared.cad', 'shared.missing'), CALLER,
            '缺少shared.missing.helper',
        )

    def test_rejects_missing_package_path(self):
        invalid = CALLER.replace('sys.path.insert(0, str(Path(__file__).resolve().parents[2]))\n', '')
        self.assert_rejects_then_recovers(self.caller, invalid, CALLER, '缺少practice公共包路径')

    def test_rejects_legacy_cross_case_import(self):
        self.assert_rejects_then_recovers(
            self.caller, CALLER.replace('from shared.cad import', 'from build_connector import'),
            CALLER, '公共函数请从shared包导入',
        )

    def test_rejects_broken_documentation_link(self):
        self.assert_rejects_then_recovers(
            'README.md', '[案例](missing.md)\n', '[案例](practice/source/case/build.py)\n',
            '失效相对链接',
        )

    def test_never_executes_case_script(self):
        self.write(self.caller, CALLER + "raise RuntimeError('CAD script must not be executed')\n")
        self.assert_passes()


if __name__ == '__main__':
    unittest.main()
