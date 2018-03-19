try:
	from setuptools import setup
except ImportError:
	from distutils.core import setup

config = {
	'description': 'The program takes two integers and a filename string as command line arguments.  It then adds a number of rows based on the two integers provided.',
	'author': 'Sunny Lam',
	'url': 'https://github.com/sunnylam13/blankRowInserter_031918_1',
	'download_url': 'https://github.com/sunnylam13/blankRowInserter_031918_1',
	'author_email': 'sunny.lam@gmail.com',
	'version': '0.1',
	'install_requires': ['nose'],
	'packages': ['openpyxl,sys'],
	'scripts': [],
	'name': 'Blank Row Inserter'
}

setup(**config)