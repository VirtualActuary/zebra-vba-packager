import doctest
import unittest
import locate
import sys

repo_dir = locate.this_dir().parent
sys.path.insert(0, str(repo_dir))


def load_tests(loader, tests, ignore):
    """
    See https://docs.python.org/3/library/doctest.html#unittest-api
    """
    modules = find_modules_with_doctests()

    for module in modules:
        tests.addTests(doctest.DocTestSuite(module))
    return tests


def find_modules_with_doctests():
    modules = []
    skip_n_parts = len(repo_dir.parts)
    for path in repo_dir.joinpath("zebra_vba_packager").rglob("*.py"):
        if path.name == "__init__.py":
            continue

        module = ".".join(path.parts[skip_n_parts:])
        module = module[:-3]
        modules.append(module)
    return modules


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
