from setuptools import setup

APP = ['KM_Assembly_Automation.py']
DATA_FILES = ['IconPictures\check_mark.png',
              'IconPictures\compare_zs_63.png',
              'IconPictures\error_icon.png',
              'IconPictures\\feedback_folder.png',
              'IconPictures\Graphical_User_Int_Theme.png',
              'green_check_mark_icon.png',
              'in_progress.png',
              'IconPictures\quit_icon.png','IconPictures\Rename.png',
              'IconPictures\quit_icon.png','IconPictures\\reset.png',
              'IconPictures\\robot_assemble.png',
              'IconPictures\source_folder.png']
OPTIONS = {
 'iconfile':'IconPictures\KraussmaffeiLogo.ico',
 'argv_emulation': True,
 'packages': ['certifi'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)