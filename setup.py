from distutils.core import setup
import py2exe
setup(
	options = {'py2exe': {
		'includes': ['lxml.etree', 'lxml._elementpath', 'gzip'],
        'bundle_files': 1
    }},
	console = [{'script': 'htmlParser.py'}],
	zipfile = None
)