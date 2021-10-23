# Extract narrations and save into single file
## Requirements

- pydub
- tqdm

Please install them by executing:
`pip install -r requirements.txt`

## Usage

python extractor.py [OPTIONS] PPTX_FILE

Options:
 --speed : the desired speed of the output audio (ex: 1.2) **Note: any value lower than 1.0 is currently not supported.**