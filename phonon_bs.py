# PowerPoint Audio Speeder
# Author: Ahnaf Tahmid 2021

import os
import shutil
import zipfile
import sys
from pydub import AudioSegment
import time

start = time.time()
presentation_file = "C:/Users/ahtah/Documents/Uni docs/PHYS 244/A_LECTURE_1.pptx" # Filename of the presentation you want to speed up
SPEED = 1.2
OUTFILE = os.path.basename(presentation_file).split('.')[0]+'_SPEEDUP_%d_x.pptx' % (SPEED)
cwd = os.path.dirname(presentation_file) # Working directory
tmp_dir = cwd+"/tempdir/" # Don't change this

if os.path.exists(tmp_dir):
    shutil.rmtree(tmp_dir)

# Open up the PPT as an archive
ppt_orig = zipfile.ZipFile(presentation_file, 'r')
print("Successfully loaded %s" % presentation_file)

# Create a temp directory
os.mkdir(tmp_dir)

# Extract the PPT contents
ppt_orig.extractall(tmp_dir)
print("Extracting PPT contents...")

# Get the media portion of the PPT
media_files = os.listdir(tmp_dir+"/ppt/media/")

print("Speeding up audio...")
# Go through each audio file
for audio_f in media_files:
    if '.m4a' in audio_f:
        audio_handle = AudioSegment.from_file(tmp_dir+"/ppt/media/"+audio_f) # Load audio

        # Modify the sample framerate
        altered_audio = audio_handle._spawn(audio_handle.raw_data, overrides={
         "frame_rate": int(audio_handle.frame_rate * SPEED)})
        
        # Modify the playback framerate
        output_audio = altered_audio.set_frame_rate(audio_handle.frame_rate)

        # Replace the audio
        output_audio.export(tmp_dir+"/ppt/media/"+audio_f, 'mp4')

# Create the new PPT
shutil.make_archive(base_name=cwd+'/'+OUTFILE, format='zip', root_dir=tmp_dir)
os.rename(cwd+'/'+OUTFILE+'.zip', cwd+'/'+OUTFILE)

# Clean up
ppt_orig.close()
if os.path.exists(tmp_dir):
    shutil.rmtree(tmp_dir)

stop = time.time()

print("Done!")
print("Conversion took %.2f seconds." % (stop-start))