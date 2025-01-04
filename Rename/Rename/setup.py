from setuptools import setup

setup(
    name='file_renamer_app',
    version='1.0',
    description='An application to rename files based on Excel data',
    author='Your Name',
    py_modules=['main'],
    install_requires=['pandas', 'openpyxl', 'tkinter'],
    entry_points={
        'console_scripts': ['file-renamer=main:main'],
    }
)