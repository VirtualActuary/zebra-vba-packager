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


def git_download(git_source, dest, revision=None):
    os.makedirs(dest, exist_ok=True)

    with working_directory(dest):
        if Path(os.getcwd()).resolve() != Path(dest).resolve():
            raise (RuntimeError(f"Could not create and enter {dest}"))

        try:
            where_git = Path([i.strip()
                              for i in subprocess.check_output(["where", "git"]).decode("utf-8").strip().split("\n")
                              if i.strip() != ""][0]).resolve()
        except (subprocess.CalledProcessError, IndexError) as e:
            raise(RuntimeError("Could not find git through `where git`"))

        i = 0
        where_sh = where_git.joinpath("bin", "sh.exe")
        while not (where_sh := where_sh.parent.parent.parent.joinpath("bin", "sh.exe")).is_file():
            if (i:=i+1) == 1000:
                raise(RuntimeError("Could not find sh relative to `where git`"))

        git = str(where_git)
        sh = str(where_sh)

        # If already on correct commit, don't do extra work
        if revision is not None and len(revision) >= 6: #possible hash
            commit = ''
            with suppress(subprocess.CalledProcessError):
                commit = [i.strip() for i in subprocess.check_output(
                            ['git', 'rev-parse', 'HEAD'], stderr=subprocess.DEVNULL).decode("utf-8").strip().split("\n")
                                if i.strip() != ""][0]
            if commit.startswith(revision):
                subprocess.call([git, "reset", "--hard"])
                subprocess.call([git, "clean", "-qdfx"])
                return

        gitremote = None
        with suppress(subprocess.CalledProcessError):
            gitremote = [i.strip() for i in subprocess.check_output(
                            ['git', 'config', '--get', 'remote.origin.url']).decode("utf-8").strip().split("\n")
                                if i.strip() != ""][0]

        # If already correct git source, don't re-download
        if gitremote != git_source:
            for i in Path(".").glob("*"):
                if i.is_file():
                    os.remove(i)
                else:
                    shutil.rmtree(i)

            subprocess.call(["git", "clone", git_source, "."])
            if not Path("./.git").is_dir():
                raise(RuntimeError(f"Could not `git clone {git_source} .`"))

        subprocess.call([git, "reset", "--hard"])
        subprocess.call([git, "clean", "-qdfx"])
        subprocess.call([sh, "-c", "for i in `git branch -a | grep remote | grep -v HEAD | grep -v master`;"
                                   "do git branch --track ${i#remotes/origin/} $i;"
                                   "done"], stderr=subprocess.DEVNULL)
        subprocess.call([git, "fetch",  "--all"], stderr=subprocess.DEVNULL)
        subprocess.call([git, "fetch", "--tags", "--force"],  stderr=subprocess.DEVNULL)

        # set revision to default branch
        if revision is None:
            revision = subprocess.check_output(
                [sh, "-", "git symbolic-ref refs/remotes/origin/HEAD | sed 's@^refs/remotes/origin/@@'"]
            ).decode("utf-8").strip()

        subprocess.call([git, "pull", "origin", revision], stderr=subprocess.DEVNULL)
        subprocess.call([git, "-c", "advice.detachedHead=false", "checkout", "--force", revision])
        subprocess.call([git, "reset", "--hard"], stderr=subprocess.DEVNULL)
