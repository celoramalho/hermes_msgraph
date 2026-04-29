from setuptools import setup, find_packages

setup(
    name="hermes_msgraph",
    version="0.1",
    description="A library for interacting with Microsoft Graph API.",
    author="Marcelo Ramalho",
    author_email="marceloramalho.dev@gmail.com",
    url="https://github.com/celoramalho/hermes_msgraph",
    packages=find_packages(),
    py_modules=["hermes_msgraph"],
    install_requires=[
        "requests>=2.0",
        "pandas>=1.0",
        "pyyaml>=5.0",
        "tqdm",
    ],
    python_requires=">=3.7",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
