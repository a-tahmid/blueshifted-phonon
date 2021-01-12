Blueshifted Phonon is a small script which allows you to speed up audio recordings within PowerPoint presentations.

It is a very specific tool written for a very specific purpose (to speed up pre-recorded PPT lectures!)

To run this script, ensure that you have Python > 3.5.x and the 'pydub' module installed.
If you do not have `pydub`, install it using pip:

pip install pydub

OR

py -m pip install pydub

OR

python -m pip install pydub

Secondly, you will need an audio codec. This has been tested using Libav but FFMPEG is also expected to work. To get the audio codec,
go to http://builds.libav.org/windows/release-gpl/ and download the latest 64-bit version. It will look something like `libav-10.6-win64.7z`.

You'll need to extract the contents of the 'bin' folder in the same directory as the `phonon_bs.py` script

***Usage***

Open up `phonon_bs.py` in your favourite text editor. You'll need to change two variables:

1) presentation_file # Set this to the full path of the presentation you want to speed up

2) SPEED # Set this to the factor by which you want to speed up the presentation by

And that's it! Just run the script and you'll get your converted presentation in the same directory as your original presentation.
