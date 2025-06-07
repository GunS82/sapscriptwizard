from pathlib import Path
from setuptools import setup, find_packages

this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding="utf-8")

setup(
    name="sapscriptwizard",
    version="0.1.0",
    description="SAP GUI Scripting automation library for Python (Windows)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Your Name",
    author_email="your.email@example.com",
    url="https://github.com/GunS82/sapscriptwizard",
    project_urls={
        "Homepage": "https://github.com/GunS82/sapscriptwizard",
    },
    packages=find_packages(include=["sapscriptwizard", "sapscriptwizard.*"]),
    py_modules=[
        "sapscriptwizard",
        "window",
        "shell_table",
        "gui_tree",
        "element_finder",
        "locator_helpers"
    ],
    package_dir={"": "."},
    install_requires=[
        "pywin32",
        "pandas",
        "polars",
        "Pillow"
    ],
    python_requires=">=3.8",
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: Microsoft :: Windows",
        "License :: OSI Approved :: MIT License",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries :: Python Modules"
    ],
    zip_safe=False,
)
