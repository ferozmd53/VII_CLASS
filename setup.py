from setuptools import setup

setup(
    name="VII_CLASS",
    version="1.0.3",
    description="Shoonya OAuth-enabled Finvasia Excel terminal with indicators",
    py_modules=[
        "VII_CLASS",
        "Shoonya_OAuth",
        "Shoonya_Indicators",
        "Weekly",
        "Intraday_Indicators",
        "get_auth",
    ],
    include_package_data=True,
    install_requires=[
        "requests",
        "pyotp",
        "xlwings",
        "pyttsx3",
        "pandas",
        "numpy",
        "scipy",
        "websocket-client",
    ],
)
