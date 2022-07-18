# `MilPy`: My Personal Python Code

This package includes code I've written for personal projects. Currently it 
includes a "flight simulator" which makes rotating glove animations for 
commercial airline flights and some code for converting and tagging MP4 video
files with metadata.

## Installation
The flight simulator uses Cartopy, which is a real bitch to install unless you
use Anaconda. I created an `environment.yml` file to make installation of
dependencies easy. The easiest way to install this is probably to clone this
repository to your computer, then build a virtual environment (automatically
named `milpy`) by
1. `cd /path/to/cloned/MilPy`
2. `conda env create -f environment.yml`
3. `conda activate milpy`
4. `python -m pip install git+https://github.com/zachariahmilby/MilPy.git`

If the last step doesn't work, you might need to do a direct `pip3 install .`