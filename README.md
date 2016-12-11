# ppmd
Converts PowerPoint OOXML (PPTX) files to Markdown-based presentations

Operation: 

* call `python3 main.py [path/to/file.pptx]`

Options:

* **`--slidesep`** - determines how individual slides will be separated. Five dashes (`-----`) by default.
* **`--headingsfirst`/`--noheadings`** - sets whether to assume the first text on each slide is a heading. Y by default. 
* **`--includenotes`/`--excludenotes`** - sets whether to include slide notes in the output. Includes by default
* **`--notesprefix`** - sets the string used to prefix lines in the notes output - three > (`>>>`) by default