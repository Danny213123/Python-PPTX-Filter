# Python PPTX Linter

This is a Python script that serves as a PPTX linter. It extracts text from the notes section of each slide and filters slides based on specified keywords.

## Usage

1. Make sure you have the `PPTX` library installed. You can install it using `pip`:

```shell
pip install python-pptx
```

2. Set the path to your presentation and copy folder, power points will be added to the presentation path and result power points (after filter) will be in the copy path.

```python
PRESENTATION_PATH = 'presentations/'
COPY_PRESENTATION_PATH = 'copy/'
```

3. Setup the `PPTX` file with the name of the power point you want to run through, and change the name of the copy file with the name you want for the resulting file.

```python
PPTX_FILE = 'your_presentation.pptx'
COPY_FILE = 'copy.pptx'
```

3. Specify the keywords you want to filter slides by in the `FILTER_KEYWORDS` list.

```python
FILTER_KEYWORDS = ['keyword1', 'keyword2', 'keyword3']
```

4. Run the script. It will extract the text from the notes section of each slide, validate the extracted text, and filter the slides based on the specified keywords.

5. The filtered slides will be saved in a new presentation file with the specified name (`COPY_FILE`).

## Correct PPTX format

All Keywords need to be nested in brackets like so:

```
Market=[keyword1, keyword2, keyword3]
```

### Correct
```
Market=[HPC]
```

### Correct
```
Market=[AI, OEM] Words Words Words Market=[HPC]
Words
```

### Incorrect
```
[AI, OEM] Words Words Words [HPC]
Words
