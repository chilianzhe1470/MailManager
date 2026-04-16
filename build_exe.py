import os
import shutil

import PyInstaller.__main__


def clean_build():
    """Remove old build artifacts before packaging."""
    for dir_name in ["build", "dist", "build_fixed", "dist_fixed"]:
        if os.path.exists(dir_name):
            print(f"Cleaning {dir_name}/")
            try:
                shutil.rmtree(dir_name)
            except PermissionError as exc:
                raise PermissionError(
                    f"无法清理 {dir_name}/。请先关闭正在运行的 MailManager.exe 后再重新打包。"
                ) from exc

    for file_name in ["MailManager_fixed.spec"]:
        if os.path.exists(file_name):
            print(f"Removing {file_name}")
            os.remove(file_name)


def build_exe():
    """Build the desktop executable with PyInstaller."""
    print("=" * 60)
    print("Building MailManager...")
    print("=" * 60)

    args = [
        "ui_app.py",
        "--name=MailManager",
        "--onefile",
        "--windowed",
        "--add-data=config;config",
        "--add-data=attachments;attachments",
        "--hidden-import=email",
        "--hidden-import=imaplib",
        "--hidden-import=smtplib",
        "--hidden-import=PySide6",
        "--hidden-import=PySide6.QtCore",
        "--hidden-import=PySide6.QtGui",
        "--hidden-import=PySide6.QtWidgets",
        "--hidden-import=json",
        "--hidden-import=logging",
        "--clean",
    ]

    if os.path.exists("icon.ico"):
        args.insert(4, "--icon=icon.ico")

    print("\nPyInstaller arguments:")
    for arg in args:
        print(f"  {arg}")
    print()

    PyInstaller.__main__.run(args)

    print("\n" + "=" * 60)
    print("Build complete")
    print("exe path: dist/MailManager.exe")
    print("=" * 60)


if __name__ == "__main__":
    clean_build()
    build_exe()
