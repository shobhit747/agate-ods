from setuptools import find_packages, setup

setup(
    name='agate-ods',
    version='0.1.0',
    description='agate-ods adds read support for Ods files to agate.',
    author='Shobhit S. Thakur',
    author_email='shobhitthakur70@gmail.com',
    license='MIT',
    install_requires=[
        'agate>=1.5.0',
        'lxml==6.0.0'
    ],
)