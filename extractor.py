import argparse
import datetime
import glob
import os
import pathlib
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile

from pydub import AudioSegment
from pydub import effects
from tqdm import tqdm

TMP_DIR = pathlib.Path("tmp/")
MEDIA_DIR = os.path.join(TMP_DIR, pathlib.Path("ppt/media"))
TRANSITION_SOUND_DIR = pathlib.Path("beeps/")

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

    if desired_speed is not None:
        if desired_speed < 1.0:
            print("values lower than 1.0 is not supported for parameter \"speed\".")
            sys.exit(0)

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
                    audio_extension = extension
                    audio_filenames_in_the_slide.append(media_filename)
        audio_filenames_in_the_file.append(audio_filenames_in_the_slide)

    TRANSITION_BEEP_SOUND = os.path.join(TRANSITION_SOUND_DIR, f'beep.{audio_extension}')

    basetime = datetime.datetime(year=2021, month=1, day=1, hour=0, minute=0, second=0)  # year, month and day are dummy value
    elapsed_second = 0
    merged_audio = None
    with open("chapters.txt", "w+") as chapter_file:
        for slide_idx, audio_filenames_in_the_slide in enumerate(tqdm(audio_filenames_in_the_file, "Processing audio files")):
            if audio_filenames_in_the_slide == []:
                continue

            elapsed_time_formatted = (basetime+datetime.timedelta(seconds=elapsed_second)).strftime("%H:%M:%S")
            chapter_file.write(f'{elapsed_time_formatted} slide{slide_idx+1}\n')

            for audio_filename in audio_filenames_in_the_slide:
                audio = AudioSegment.from_file(os.path.join(MEDIA_DIR, audio_filename))
                if desired_speed is not None:
                    audio = audio.speedup(playback_speed=desired_speed, chunk_size=150, crossfade=25)
                
                # When process the first audio file 
                if merged_audio is None:
                    transition_sound_filename = os.path.join(TRANSITION_SOUND_DIR, f'beep.{get_extension(audio_filename)}')
                    transition_sound = AudioSegment.from_file(transition_sound_filename)
                    transition_sound = match_audio_volume(modulated_sound=transition_sound, base_sound=audio)
                    merged_audio = audio
                    elapsed_second += audio.duration_seconds
                    continue

                appended_sound = transition_sound + audio
                merged_audio += appended_sound
                elapsed_second += appended_sound.duration_seconds

            merged_audio.export("narration.mp3", format="mp3", parameters=["-ac","2","-ar","8000"])

    shutil.rmtree(TMP_DIR)

if __name__ == '__main__':
    main()