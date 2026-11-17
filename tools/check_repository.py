"""只读检查维护代码与文档；不执行案例、不连接CAD。"""
import ast
from pathlib import Path
import re
import sys
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parents[1]


def exported_names(tree):
    names = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            names.update(alias.asname or alias.name.split('.')[0] for alias in node.names)
        elif isinstance(node, ast.Assign):
            names.update(target.id for target in node.targets if isinstance(target, ast.Name))
    return names


def path_value(expression, path):
    """仅求值路径表达式，不执行被检查脚本的任何语句。"""
    env = {'Path': Path, 'str': str, '__file__': str(path),
           'ROOT': ROOT, 'HERE': path.parent}
    for node in ast.walk(expression):
        if isinstance(node, ast.Name) and node.id not in env:
            return None
        if isinstance(node, ast.Call):
            direct = isinstance(node.func, ast.Name) and node.func.id in ('str', 'Path')
            resolve = isinstance(node.func, ast.Attribute) and node.func.attr == 'resolve' and not node.args and not node.keywords
            if not (direct or resolve):
                return None
        if isinstance(node, ast.Attribute) and node.attr not in ('resolve', 'parent', 'parents'):
            return None
        if not isinstance(node, (ast.Expression, ast.Call, ast.Name, ast.Attribute, ast.Constant,
                                 ast.BinOp, ast.Div, ast.Subscript, ast.Load)):
            return None
    try:
        return Path(eval(compile(expression, '<path-only>', 'eval'), {'__builtins__': {}}, env)).resolve()
    except (TypeError, ValueError, IndexError):
        return None


def main():
    errors = []
    files = [p for base in ('practice', 'tools') for p in (ROOT / base).rglob('*.py')]
    trees = {}
    for path in files:
        try:
            trees[path] = ast.parse(path.read_text(encoding='utf-8-sig'), filename=str(path))
        except SyntaxError as error:
            errors.append(f'{path.relative_to(ROOT)}:{error.lineno}: {error.msg}')
    shared = {f'shared.{p.stem}': exported_names(tree) for p, tree in trees.items()
              if p.parent == ROOT / 'practice/shared'}
    imports = 0
    for path, tree in trees.items():
        uses_shared = False
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module in ('build_connector', 'single_part'):
                errors.append(f'{path.relative_to(ROOT)}:{node.lineno}: 公共函数请从shared包导入')
            if isinstance(node, ast.ImportFrom) and node.module and node.module.startswith('shared.'):
                uses_shared = True
                imports += 1
                for alias in node.names:
                    if alias.name not in shared.get(node.module, set()):
                        errors.append(f'{path.relative_to(ROOT)}:{node.lineno}: 缺少{node.module}.{alias.name}')
        if uses_shared:
            inserts = [n for n in ast.walk(tree) if isinstance(n, ast.Call)
                       and ast.unparse(n.func) == 'sys.path.insert' and len(n.args) == 2]
            if not any(path_value(ast.Expression(n.args[1]), path) == ROOT / 'practice' for n in inserts):
                errors.append(f'{path.relative_to(ROOT)}: 缺少practice公共包路径')
    documents = [*ROOT.glob('*.md'), *[p for base in ('docs', 'practice', 'tools')
                 for p in (ROOT / base).rglob('*.md') if 'history' not in p.parts]]
    links = 0
    for path in documents:
        text = re.sub(r'```.*?```', '', path.read_text(encoding='utf-8-sig'), flags=re.S)
        for raw in re.findall(r'!?\[[^\]\n]*\]\(([^\s)]+)\)', text):
            target = unquote(raw.split('#', 1)[0])
            if not target or re.match(r'\w+:', target):
                continue
            links += 1
            resolved = (path.parent / target).resolve()
            if not resolved.is_relative_to(ROOT) or not resolved.exists():
                errors.append(f'{path.relative_to(ROOT)}: 失效相对链接 {target}')
    for error in errors:
        print(error)
    print(f'Python语法 {len(trees)}/{len(files)}；公共导入 {imports}；相对链接 {links}；错误 {len(errors)}')
    return bool(errors)


if __name__ == '__main__':
    sys.exit(main())
