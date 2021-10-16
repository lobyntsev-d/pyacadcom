from setuptools import setup, find_packages

setup(
    name='pyacadcom',
    version='0.0.2',
    description='Python AutoCAD COM automation via pywin32',
    packages=find_packages(),
    author='Dmitriy Lobyntsev',
    author_email='dmitriy@lobyntsev.com',
    url="http://github.com/lobyntsev-d/pyacadcom",
    install_requires=[
        'pywin32>=214',
    ],
    keywords=["autocad", "automation", "activex", "pywin32"],
    license="BSD License",
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python",
        "License :: OSI Approved :: BSD License",
        "Development Status :: 4 - Beta",
        "Environment :: Win32 (MS Windows)",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3.7",
        "Operating System :: Microsoft :: Windows",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Scientific/Engineering",
    ],
    long_description=open('README.rst').read()
)