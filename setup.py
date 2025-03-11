from setuptools import setup, find_packages

setup(
    name='yeeko_xlsx_export',
    version='0.0.2',
    description='generic xlsx export tool',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    author='vash Lucian',
    author_email='lucian@yeeko.org',
    url='https://github.com/yeeko-org/xlsx_export',
    packages=find_packages(),
    install_requires=[
        "django>=4.1.15",
        "djangorestframework>=3.13.1",
        "XlsxWriter>=3.0.2",
        "pytz>=2021.3",
    ],
    python_requires='>=3.6',
)
