#!/bin/bash
DISPLAY=''
sudo apt update && sudo apt install -y ffmpeg libavcodec-extra
yes | pip install -r ../requirements.txt --quiet
if [ $? -eq 0 ]; then
echo "Done."
else
echo "Something went wrong."
fi
sleep 5