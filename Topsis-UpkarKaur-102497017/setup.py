from setuptools import setup, find_packages

setup(
    name="Topsis-UpkarKaur-102497017",
    version="1.0.1",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "numpy"
    ],
    entry_points={
        'console_scripts': [
            'topsis=topsis_upkar.topsis:main'
        ]
    },
    author="Upkar Kaur",
    author_email="your_email@gmail.com",
    description="A Python package to implement TOPSIS method.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    python_requires='>=3.6',
)
