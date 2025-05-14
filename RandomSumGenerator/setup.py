from setuptools import setup, find_packages

setup(
    name="random_sum_generator",
    version="0.2.0",
    description="Generate random numbers that sum to a total with min/max bounds per part.",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    author="Md. Fakhruddin Gazzali Fahim",
    author_email="fahim9778@gmail.com",
    url="https://github.com/fahim9778/random_sum_generator",
    project_urls={
        "Bug Tracker": "https://github.com/fahim9778/random_sum_generator/issues",
        "Source": "https://github.com/fahim9778/random_sum_generator"
    },
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries",
    ],
    license="MIT",
    python_requires='>=3.6',
    install_requires=[],
)