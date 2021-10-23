import argparse
import datetime
import glob
import os
import pathlib
import platform
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

from pydub import AudioSegment
from pydub import effects
from tqdm import tqdm

# Choose from 'low', 'medium', 'high'. The better the quality is, the bigger the output becomes.
SOUND_QUALITY = 'medium'

TMP_DIR = pathlib.Path("tmp/")
XML_DIR = os.path.join(TMP_DIR, "ppt/slides/_rels/")
MEDIA_DIR = os.path.join(TMP_DIR, pathlib.Path("ppt/media"))
TRANSITION_SOUND_DIR = pathlib.Path("beeps/")

PATTERN_TO_EXTRACT_SLIDE_NUM = re.compile(r'.*slide(\d+).xml.rels')
SAMPLING_FREQUENCIES = {'low': 8000, 'medium': 22050, 'high': 44100}

def get_extension(filename: str) -> str:
    return filename.split(".")[-1]

def match_audio_volume(modulated_sound: AudioSegment, base_sound: AudioSegment) -> AudioSegment:
    change_of_dbFS = base_sound.dBFS - modulated_sound.dBFS - 20  # Adjustment
    return modulated_sound.apply_gain(change_of_dbFS)

def main() -> None:

    parser = argparse.ArgumentParser(description='Extract narrations from slides and combine them into single mp3 file.')
    parser.add_argument('filepath', metavar='{path to pptx}', type=str, help='path to the target pptx file.')
    parser.add_argument('--speed', metavar='{float value}', type=float, help='the relative speed of output audio file (ex: 1.2)')
    args = parser.parse_args()
    pptx_filepath = args.filepath
    pptx_dirpath = os.path.dirname(pptx_filepath)
    output_dir = os.path.join(pptx_dirpath, pathlib.Path("audio/"))
    pptx_basename = os.path.basename(pptx_filepath).replace(" ", "_")
    desired_speed = args.speed

    if desired_speed is not None:
        if desired_speed < 1.0:
            print("values lower than 1.0 is not supported for parameter \"speed\".")
            sys.exit(0)

    if not os.path.exists(pptx_filepath):
        print("Couldn't find such pptx file.")
        sys.exit(0)

    if os.path.exists(TMP_DIR):
        print("Found a folder named \"tmp\". Remove it and try again")
        sys.exit(0)

    os.mkdir(TMP_DIR)
    os.makedirs(output_dir, exist_ok=True)
    with zipfile.ZipFile(pptx_filepath) as pptx_file:
        try:
            pptx_file.extractall(TMP_DIR)
        except zipfile.BadZipfile:
            print("It seems python's pptx extraction process has failed. Do you want to try system's?")
            answer = input("[Y/n]:")
            if answer == 'n':
                print("Aborted.")
                sys.exit(0)
            
            current_os = platform.system()
            # # Windows very often fails to extract pptx so currently deprecated
            # if current_os == "Windows":
            #     subprocess.run(f'call powershell -command "Expand-Archive -Force {pptx_filepath} {TMP_DIR}"', encoding="shift-jis", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if current_os in ("Linux", "macOS"):
                subprocess.run(f"unzip -o {pptx_filepath} -d {TMP_DIR}", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                print("You are using such a rare OS. There's nothing I can do for you anymore.")
                sys.exit(0)

    slide_xmls = glob.glob(os.path.join(XML_DIR, "*.xml.rels"))
    slide_xmls = sorted(slide_xmls, key=lambda xmlpath : int(PATTERN_TO_EXTRACT_SLIDE_NUM.search(xmlpath).group(1)))

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

    basetime = datetime.datetime(year=2021, month=1, day=1, hour=0, minute=0, second=0)  # year, month and day are dummy values
    elapsed_second = 0
    merged_audio = None
    with open(os.path.join(output_dir, f"chapters-{pptx_basename}.txt"), "w+") as chapter_file:
        for slide_idx, audio_filenames_in_the_slide in enumerate(tqdm(audio_filenames_in_the_file, "Processing audio files")):
            if audio_filenames_in_the_slide == []:
                continue

            elapsed_time_formatted = (basetime+datetime.timedelta(seconds=elapsed_second)).strftime("%H:%M:%S")
            chapter_file.write(f'{elapsed_time_formatted} slide{slide_idx+1}\n')

            for audio_filename in audio_filenames_in_the_slide:
                audio = AudioSegment.from_file(os.path.join(MEDIA_DIR, audio_filename))
                if desired_speed is not None:
                    if desired_speed <= 1.4:
                        chunk_size = 150
                    elif 1.4 < desired_speed <= 1.6:
                        chunk_size = 100
                    else:
                        chunk_size = 50
                    audio = audio.speedup(playback_speed=desired_speed, chunk_size=chunk_size, crossfade=25)
                
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

            merged_audio.export(os.path.join(output_dir, f"narration-{pptx_basename}.mp3"), format="mp3", parameters=["-ac", "2", "-ar", str(SAMPLING_FREQUENCIES.get(SOUND_QUALITY))])

    print("Done.")

    shutil.rmtree(TMP_DIR)

if __name__ == '__main__':
    main()