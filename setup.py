import os
import shutil
import subprocess


def cleanup():
    dirs = ["build", "dist"]
    for i in dirs:
        if os.path.isdir(i):
            shutil.rmtree(i)


if __name__ == '__main__':
    # cleanup_prev
    cleanup()

    # build_nsclient_deploy
    cmd = ["python", "-O", "-m", "PyInstaller", "exchange_postitems.spec"]

    subprocess.run(cmd)

    # move to root
    os.replace(os.curdir + "\\dist\\exchange_postitems.exe", os.curdir + "\\exchange_postitems.exe")

    # cleanup
    cleanup()
