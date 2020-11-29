from setuptools import setup, find_packages

with open("README.md") as f:
    long_description = f.read()

setup(
    name = 'JNAirport_Excel_Tool',
    version = "0.0.1",
    author="lbcoder",
    author_email="lbcoder@hotmail.com",
    description="A excel helper tool for jinan airport",
    long_description = long_description,
    long_description_content_type = "text/markdown",
    url = "https://github.com/weilancys/jnairport_excel_tool",
    packages = find_packages(),
    install_requires = [
        "openpyxl"
    ],
    classifiers = [
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    entry_points = {
        'console_scripts': [
            'jnairport_excel_tool = JNAirport_Excel_Tool:main',
        ],
    }
)