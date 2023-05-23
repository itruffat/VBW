from setuptools import setup, find_packages

p  = find_packages("src")

setup(
    name='VBW',
    version='0.0.0a1',
    author='itrufat',
    description='A wrapper to run VBS from Python.',
    long_description='A Wrapper to run VBS from Python.',
    long_description_content_type='text/markdown',
    url='https://github.com/itruffat/JestingLang',
    packages=['VBW', 'VBW.core', 'VBW.wrappers'] ,
    package_dir={'': 'src'},
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
)
