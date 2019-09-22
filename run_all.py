import json
import os
import shutil
import stat
import subprocess
import tempfile
from glob import glob
from pathlib import Path

import convert
import email_sender
import merge

CODEDIR = Path.cwd()
CONFIG = CODEDIR / "config.json"


def update_tests_for_branch(i, CDTESTDIR):
    subprocess.run(f"cd {CDTESTDIR} && git checkout {i} && git pull --rebase", shell=True, check=True)


def server_call(filename, config, email):
    tempdir = tempfile.TemporaryDirectory()
    TEMPDIR = Path(tempdir.name)
    CDTESTDIR = TEMPDIR / "cd-test"
    OUTPUTDIR = TEMPDIR / "output"

    if not Path(OUTPUTDIR).is_dir():
        os.mkdir(OUTPUTDIR)

    subprocess.run([
        "git", "clone", "https://git.labs.nuance.com/nlps-qa/cd-test",
        str(CDTESTDIR)
    ], check=True)

    dest_filename = OUTPUTDIR / f"{filename}.xlsx"
    for i in config:
        branch = i['branch']
        folder = i['folder']
        update_tests_for_branch(branch, CDTESTDIR)
        convert.generate_excel(CDTESTDIR / folder,
                                OUTPUTDIR / f"{branch}+{folder}.xlsx")
                                
    merge.merge_excel(OUTPUTDIR, dest_filename)
    email_sender.send_email(dest_filename, filename, email)
