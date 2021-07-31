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
            where_git = Path(sh_lines('where git')[0]).resolve()
        except (subprocess.CalledProcessError, IndexError) as e:
            raise(RuntimeError("Could not find git through `where git`"))

        i = 0
        where_sh = where_git.joinpath("bin", "sh.exe")
        while not (where_sh := where_sh.parent.parent.parent.joinpath("bin", "sh.exe")).is_file():
            if (i := i+1) == 1000:
                raise(RuntimeError("Could not find sh relative to `where git`"))

        git = str(where_git)
        sh = str(where_sh)

        # If already on correct commit, don't do extra work
        def is_on_ref():
            if revision is not None:
                with suppress(subprocess.CalledProcessError, IndexError):
                    if sh_lines([git, 'rev-parse', 'HEAD'],
                                stderr=subprocess.DEVNULL)[0].startswith(revision):
                        return True

                with suppress(subprocess.CalledProcessError):
                    if revision in (sh_lines([git, "branch", "--show-current"], stderr=subprocess.DEVNULL) +
                                    sh_lines([git, "tag", "-l", "--contains", "HEAD"], stderr=subprocess.DEVNULL)):
                        return True

            return False

        if is_on_ref():
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
        sh_quiet([sh, "-c", "for i in `git branch -a | grep remote | grep -v HEAD | grep -v master`;"
                            "do git branch --track ${i#remotes/origin/} $i;"
                            "done"])
        sh_quiet([git, "fetch",  "--all"])
        sh_quiet([git, "fetch", "--tags", "--force"])

        # set revision to default branch
        if revision is None:
            revision = sh_lines(
                [sh, "-", "git symbolic-ref refs/remotes/origin/HEAD | sed 's@^refs/remotes/origin/@@'"])[0]

        sh_quiet([git, "pull", "origin", revision])
        sh_quiet([git, "-c", "advice.detachedHead=false", "checkout", "--force", revision])
        sh_quiet([git, "reset", "--hard"])
        sh_quiet([git, "clean", "-qdfx"])

        if revision is None or is_on_ref():
            return None
        else:
            raise RuntimeError(f"Could not check out {revision}")
