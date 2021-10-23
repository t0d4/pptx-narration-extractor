import os
import pathlib
import shutil
import glob
import sys
import xml.etree.ElementTree as ET
import argparse
import zipfile
from pydub import AudioSegment

TMP_DIR = pathlib.Path("tmp/")
MEDIA_DIR = os.path.join(TMP_DIR, pathlib.Path("ppt/media"))
TRANSITION_BEEP_SOUND = pathlib.Path("beep.m4a")

def get_extension(filename: str) -> str:
    return filename.split(".")[-1]

def match_audio_volume(modulated_sound: AudioSegment, base_sound: AudioSegment) -> AudioSegment:
    change_of_dbFS = base_sound.dBFS - modulated_sound.dBFS
    return modulated_sound.apply_gain(change_of_dbFS)

def main() -> None:

    parser = argparse.ArgumentParser(description='Extract narrations in slides and combine them into single file.')
    parser.add_argument('filename', metavar='{path to pptx}', type=str, help='the target pptx file.')
    parser.add_argument('--speed', metavar='{float value}', type=float, help='the speed of generated audio file (ex: 1.2)')
    args = parser.parse_args()
    pptx_filename = args.filename
    desired_speed = args.speed

    if not os.path.exists(pptx_filename):
        print("Couldn't find such pptx file.")
        sys.exit(0)

    if os.path.exists(TMP_DIR):
        print("Found a folder named \"tmp\". Remove it and try again")
        sys.exit(0)

    os.mkdir(TMP_DIR)
    with zipfile.ZipFile(pptx_filename) as pptx_file:
        pptx_file.extractall(TMP_DIR)

    XML_DIR = os.path.join(TMP_DIR, "ppt/slides/_rels/")
    slide_xmls = glob.glob(os.path.join(XML_DIR, "*.xml.rels"))

    audio_filenames_in_the_file = []
    for xml in slide_xmls:
        audio_filenames_in_the_slide = []
        with open(xml) as xml_file:
            parsed_xml = ET.parse(xml_file)
        for node in parsed_xml.iter():
            media_filepath = node.attrib.get("Target")
            if media_filepath is not None:
                media_filename = os.path.basename(media_filepath)
                extension = get_extension(media_filename)
                if extension in ("m4a", "wav", "wma") and media_filename not in audio_filenames_in_the_slide:
                    audio_filenames_in_the_slide.append(media_filename)
        audio_filenames_in_the_file.append(audio_filenames_in_the_slide)

    print(audio_filenames_in_the_file)



    #shutil.rmtree(TMP_DIR)

if __name__ == '__main__':
    main()