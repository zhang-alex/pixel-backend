from setuptools import setup

setup(
    name='pixel-backend',
    version='0.1',
    description='a primitive python module that takes an array of values and exports it to a Microsoft Word document.',
    author='Alex Zhang',
    author_email='alexzhang02030203@gmail.com',
    packages=['env'],
    install_requires=[
        'numpy',
        'setuptools'
    ],
)
