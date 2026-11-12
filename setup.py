from setuptools import setup

setup(
    name="VII_CLASS",
    version="1.0.0",
    description="Shoonya OAuth-enabled Finvasia Excel terminal with indicators",
    py_modules=[
        "Finvasia_Core_V3_002",
        "Shoonya_OAuth",
        "Shoonya_Indicators",
        "get_auth",
    ],
    include_package_data=True,
    install_requires=[
        "requests", "pyotp", "xlwings", "pyttsx3",
        "pandas", "numpy", "websocket-client",
    ],
)
