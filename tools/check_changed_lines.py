"""按GitHub事件检查完整差异；强推后按需取回不在当前历史中的比较基准。"""
import os
import subprocess
import sys


def main():
    event = os.environ.get('GITHUB_EVENT_NAME', 'workflow_dispatch')
    if event == 'pull_request':
        base = os.environ.get('PR_BASE_SHA', '')
        if not base:
            print('PR缺少base SHA，不能确定比较范围', file=sys.stderr)
            return 1
    elif event == 'push':
        base = os.environ.get('PUSH_BEFORE_SHA', '')
        if base and base != '0' * 40:
            available = subprocess.run(
                ['git', 'cat-file', '-e', base + '^{commit}'],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
            )
            if available.returncode:
                print('推送基准不在当前检出中，按SHA获取后检查完整差异', flush=True)
                # 临时credential helper使用CI只读GH_TOKEN；不把凭证写入Git配置。
                fetched = subprocess.run(
                    ['git', '-c', 'credential.helper=', '-c', 'credential.helper=!gh auth git-credential',
                     'fetch', '--no-tags', '--depth=1', 'origin', base], check=False,
                )
                if fetched.returncode:
                    return fetched.returncode
        else:
            base = ''
    elif event == 'workflow_dispatch':
        base = ''
    else:
        print('未支持的事件，不能确定比较范围: ' + event, file=sys.stderr)
        return 1
    args = ['git', 'diff', '--check', base, 'HEAD'] if base else ['git', 'show', '--format=', '--check', 'HEAD']
    return subprocess.run(args, check=False).returncode


if __name__ == '__main__':
    sys.exit(main())
