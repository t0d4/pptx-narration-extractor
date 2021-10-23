# Extract narrations and save into single file
## Requirements

### Software
- ffmpeg or libav

Linux
`sudo apt install ffmpeg libavcodec-extra`

Mac
`brew install ffmpeg`

Windows
1. Download and extract libav from (here)[http://builds.libav.org/windows/]
2. Add the libav /bin folder to your PATH

### Python package
- pydub
- tqdm

Version of the packages are arbitrary.

Please install them by executing:
`pip install -r requirements.txt`

## Usage

python extractor.py [OPTIONS] PPTX_FILE

Options:
 --speed : the desired speed of the output audio (ex: 1.2) **Note: any value lower than 1.0 is currently not supported.**