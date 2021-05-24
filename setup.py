from setuptools import setup

setup(name="cost_calculator",
      version=__version__,
      author="Ryomei Osaki",
      author_email="osaki.rr@gmail.com",
      packages=["cost_calculator"],
      url="https://github.com/KART-Software/cost-calculator",
      entry_points={
          "console_scripts": ["cost_calculator = cost_calculator.cli:main"],
      },
      classifiers=[],
      diescription=("Python 3 library for Cost Calculation."))
