from setuptools import setup
import os

here = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(here, "cost_calculator", "version.py")) as fp:
    exec(fp.read())

setup(
    name="cos",
    version=__version__,
    author="Ryomei Osaki",
    author_email="o.ryomei1020@gmail.com",
    packages=["cost_calculator"],
    install_requires=["openpyxl", "pdfminer.six"],
    url="https://github.com/KART-Software/cost-calculator",
    entry_points={
        "console_scripts": ["costcalc = cost_calculator.cli:main"],
    },
    python_requires=">=3.7",
    classifiers=[],
    diescription=("Python 3 library for Cost Calculation."),
)
