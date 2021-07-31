import subprocess
from contextlib import suppress
from pathlib import Path
import shutil
import os
from contextlib import contextmanager
import download


@contextmanager
def working_directory(path):
    """
    A context manager which changes the working directory to the given
    path, and then changes it back to its previous value on exit.
    Usage:
    > # Do something in original directory
    > with working_directory('/my/new/path'):
    >     # Do something in new directory
    > # Back to old directory
    """

    prev_cwd = os.getcwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(prev_cwd)


def sh_lines(command, **kwargs):
    shell = isinstance(command, str)
    lst = subprocess.check_output(command,
                                  shell=isinstance(command, str),
                                  **kwargs).decode("utf-8").strip().split("\n")
    return [] if lst == [''] else [i.strip() for i in lst]


def sh_quiet(command):
    return subprocess.call(command,
                           shell=isinstance(command, str),
                           stderr=subprocess.DEVNULL,
                           stdout=subprocess.DEVNULL)


def git_download(git_source, dest, revision=None):
    os.makedirs(dest, exist_ok=True)

    with working_directory(dest):
        if Path(os.getcwd()).resolve() != Path(dest).resolve():
            raise (RuntimeError(f"Could not create and enter {dest}"))

        try:
            git = str(Path(sh_lines('where git')[0]).resolve())
        except (subprocess.CalledProcessError, IndexError) as e:
            raise(RuntimeError("Could not find git through `where git`"))

        # Test if we are currently tracking the ref
        def is_on_ref(revision):
            if revision is None:
                return False
            commit = sh_lines([git, 'rev-parse', 'HEAD'])[0]
            return commit == sh_lines([git, "rev-list", "-n", "1", revision])[0]

        if Path(".git").is_dir():
            if is_on_ref(revision):
                sh_quiet([git, "reset", "--hard"])
                sh_quiet([git, "clean", "-qdfx"])
                return None

        gitremote = None # noqa
        with suppress(subprocess.CalledProcessError):
            gitremote = sh_lines([git, 'config', '--get', 'remote.origin.url'])[0]

        # If already correct git source, don't re-download
        if gitremote != git_source:
            for i in Path(".").glob("*"):
                if i.is_file():
                    os.remove(i)
                else:
                    shutil.rmtree(i)

            subprocess.call([git, "clone", git_source, str(Path(".").resolve())])
            if not Path("./.git").is_dir():
                raise(RuntimeError(f"Could not `git clone {git_source} .`"))

        sh_quiet([git, "reset", "--hard"])
        sh_quiet([git, "clean", "-qdfx"])

        for get_all_upstream in [False, True]:
            if get_all_upstream:
                for branch in sh_lines([git, "branch", "-a"]):
                    if "->" in branch:
                        continue
                    sh_quiet([git, "branch", "--track", branch.split("/")[-1], branch])

            sh_quiet([git, "fetch",  "--all"])
            sh_quiet([git, "fetch", "--tags", "--force"])

            # set revision to default branch
            if revision is None:
                revision = sh_lines([git, "symbolic-ref", "refs/remotes/origin/HEAD"])[0].split("/")[-1]

            sh_quiet([git, "pull", "origin", revision])
            sh_quiet([git, "checkout", "--force", revision])

            if revision is None or is_on_ref(revision):
                return None

        raise RuntimeError(f"Could not check out {revision}")
