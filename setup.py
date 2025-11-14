from setuptools import setup, find_packages

setup(
    name="MT5Reporter",
    version="1.0",
    packages=find_packages(),
    install_requires=[
        'MetaTrader5',
        'pandas',
        'openpyxl',
        'tkcalendar',
        'sv-ttk',
        'pywin32',
    ],
    entry_points={
        'console_scripts': [
            'mt5reporter=src.app:main',
        ],
    },
)