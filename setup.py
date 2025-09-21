"""
setup.py for pyhub-office-automation package
"""

import os
import sys

from setuptools import find_packages, setup

# Add the package directory to the path to import version
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "pyhub_office_automation"))
from version import get_version, get_version_info


# Read README
def read_readme():
    try:
        with open("README.md", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return "Python-based Excel and HWP automation package for AI agents"


# Read requirements
def read_requirements():
    with open("requirements.txt", "r", encoding="utf-8") as f:
        lines = f.readlines()

    requirements = []
    for line in lines:
        line = line.strip()
        if line and not line.startswith("#"):
            # Handle conditional dependencies
            if "; sys_platform" in line:
                requirements.append(line)
            else:
                requirements.append(line.split(";")[0].strip())
    return requirements


# Platform-specific dependencies
install_requires = read_requirements()

# Development dependencies
extras_require = {
    "dev": [
        "pytest>=7.0.0",
        "pytest-cov>=4.0.0",
        "pytest-mock>=3.10.0",
        "black>=24.0.0",
        "isort>=5.13.0",
        "flake8>=7.0.0",
        "pre-commit>=3.0.0",
    ],
    "build": [
        'PyInstaller>=5.0.0; sys_platform == "win32"',
    ],
}

setup(
    name="pyhub-office-automation",
    version=get_version(),
    author="Chinseok Lee",
    author_email="me@pyhub.kr",
    description="Python-based Excel and HWP automation package for AI agents",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/pyhub-apps/pyhub-office-automation",
    project_urls={
        "Bug Reports": "https://github.com/pyhub-apps/pyhub-office-automation/issues",
        "Source": "https://github.com/pyhub-apps/pyhub-office-automation",
    },
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.13",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    python_requires=">=3.13",
    install_requires=install_requires,
    extras_require=extras_require,
    entry_points={
        "console_scripts": [
            "oa=pyhub_office_automation.cli.main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "pyhub_office_automation": ["*.txt", "*.md"],
    },
    keywords="office automation excel hwp ai agent cli",
    platforms=["Windows"],
    zip_safe=False,
)
