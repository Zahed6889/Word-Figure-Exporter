![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

# Word Figure Exporter 

This repository contains VBA macro that extract all InlineShape figures
from a Word document, fit the page to each image, and export each as
a numbered PDF (image1.pdf, image2.pdf, ...).

This code is especially useful when preparing academic papers in Word that must ultimately be compiled in LaTeX, where figures need to be provided as external PDF files. Extracting EMF images directly from Word and converting them with external tools such as Inkscape often leads to scaling issues, missing elements, distorted fonts, or unexpected cropping. This macro avoids those problems entirely by exporting each figure straight from Word into a perfectly sized PDF, preserving vector quality and ensuring LaTeX-friendly output.

## Installation
1. Open Microsoft Word
2. Press ALT+F11 to open the VBA Editor
3. Go to File → Import File...
4. Select `ExportFigures.bas`
5. Run `ExportAllFiguresAsPDFs_SelectFolder` from the Macros menu

## Features
- Automatically detects figures
- Exports each figure into a clean temporary doc
- Adjusts page size to the figure
- Saves numbered PDFs to a user-selected folder

## License

This project is licensed under the MIT License — you are free to use, modify,
and distribute this code, provided that the original copyright notice is
included. See the `LICENSE` file for full details.

## Author
Developed by **Zahed Dastan**, 2025.  
Feel free to contribute or open issues on GitHub.

## Contributing

Contributions, issues, and feature requests are welcome!
Feel free to open a pull request or issue to enhance this project.
