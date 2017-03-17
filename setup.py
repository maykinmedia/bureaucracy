import os
from setuptools import find_packages, setup

with open(os.path.join(os.path.dirname(__file__), 'README.rst')) as readme:
    README = readme.read()

# allow setup.py to be run from any path
os.chdir(os.path.normpath(os.path.join(os.path.abspath(__file__), os.pardir)))

setup(
    name='burocracy',
    version='0.2',
    license='MIT',

    install_requires=[
        'pypandoc',
        'python-docx',
        'python-pptx>=0.6.2',
    ],
    include_package_data=True,
    packages=find_packages(exclude=["tests"]),

    setup_requires=['pytest-runner'],
    tests_require=[
        'PyPDF2',
        'pytest-cov',
    ],

    description='Templating and pdf generation for docx/pptx files',
    long_description=README,
    author='Maykin Media, Robin Ramael, Sergei Maertens',
    author_email='robin.ramael@maykinmedia.nl, sergei@maykinmedia.nl',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Topic :: Internet :: WWW/HTTP',
        'Topic :: Internet :: WWW/HTTP :: Dynamic Content',
    ],
)
