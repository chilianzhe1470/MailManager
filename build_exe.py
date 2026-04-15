import PyInstaller.__main__
import os
import shutil


def clean_build():
    """清理旧的构建文件"""
    for dir_name in ['build', 'dist']:
        if os.path.exists(dir_name):
            print(f"清理 {dir_name}/")
            shutil.rmtree(dir_name)


def build_exe():
    """构建 exe 文件"""
    print("=" * 60)
    print("开始打包 MailManager...")
    print("=" * 60)

    args = [
        'ui_app.py',
        '--name=MailManager',
        '--onefile',
        '--windowed',
        '--add-data=config;config',
        '--add-data=attachments;attachments',
        '--hidden-import=email',
        '--hidden-import=imaplib',
        '--hidden-import=smtplib',
        '--hidden-import=tkinter',
        '--hidden-import=json',
        '--hidden-import=logging',
        '--clean',
    ]

    # 如果有图标文件，添加图标
    if os.path.exists('icon.ico'):
        args.insert(4, '--icon=icon.ico')

    print("\nPyInstaller 参数:")
    for arg in args:
        print(f"  {arg}")
    print()

    PyInstaller.__main__.run(args)

    print("\n" + "=" * 60)
    print("✅ 打包完成！")
    print(f"📦 exe 文件位置: dist/MailManager.exe")
    print("=" * 60)


if __name__ == '__main__':
    clean_build()
    build_exe()

