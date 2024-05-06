from setuptools import setup

APP = ['process_spreadsheet.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['pandas', 'numpy', 'pytz'],
    'includes': ['bs4', 'dateutil'],  # Ensure bs4 and dateutil are explicitly included
}

setup(
    app=APP,
    name='Process Spreadsheet',
    version='0.1',
    description='A tool to process and analyze spreadsheet data.',
    author='Your Name',  # Replace 'Your Name' with your name
    author_email='your.email@example.com',  # Replace with your email
    url='https://www.yourwebsite.com',  # Replace with your website or remove if not applicable
    license='MIT',  # Ensure this matches your intended license
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
    install_requires=[
        'pandas>=1.0',
        'beautifulsoup4',
        'numpy',
        'python-dateutil',
        'pytz'
    ]
)
