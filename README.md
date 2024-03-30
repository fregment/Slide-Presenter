# Slide-Presenter
 A tool to automatically generate video presentation with a voice over from Power Point presentations.


## Installation

`pip install -r requirements.txt`

## Preparing Input Files

1. In Keynote or PowerPoint, first delete all hidden or skipped slides in your presentation file. 

2. Next, export the slide presentation as `.pptx` file to project directory set `PPTX_PATH` to file location. The script will read the speaker notes from each slide and generate a audio clip from it. 

3. Finally, export all slides as images in TIFF format to folder `slides`.

## Usage

```$ python slides_to_video.py```

The generated video will be in the `output` folder

## Options

`GAP_BETWEEN_SLIDES_MS` the silence between each slide in milliseconds

`MOV_EXPORT_FPS` the frames per second of the generated video

`TTS_VOICE_NAME` the OpenAI TTS voice for the voice over 