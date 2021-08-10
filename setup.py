import setuptools
import os
bin_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "zebra_vba_packager", "bin")

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="zebra-vba-packager",
    author="Rudolf Byker, Simon Streicher",
    author_email="rudolfbyker@gmail.com, sfstreicher@gmail.com",
    description="A system for aggregating and mapping VBA projects from different sources into a single library ",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/AutoActuary/zebra-vba-packager",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: Other/Proprietary License",
        "Operating System :: OS Independent",
    ],
    package_data={'zebra_vba_packager': [os.path.join('bin', i) for i in os.listdir(bin_dir)]},
    python_requires='>=3.4',
    use_scm_version={
        'write_to': 'zebra_vba_packager/version.py',
    },
    setup_requires=[
        'setuptools_scm',
    ],
    install_requires=[
        'locate',
        'download',
        'pathvalidate'
    ]
)
