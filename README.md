# OCR Tess AHK

A lightweight AutoHotkey utility that performs OCR (Optical Character Recognition) on clipboard images using Tesseract OCR and ImageMagick preprocessing.

## Features

- **Instant OCR**: Press `Ctrl+Shift+V` to extract text from any image on your clipboard
- **Smart Image Preprocessing**: Automatically enhances images using ImageMagick for better OCR accuracy
- **Persistent Configuration**: Remembers Tesseract and ImageMagick paths via environment variables
- **Works with Any Screenshot Tool**: Compatible with any screenshotting software that copies to clipboard

## Prerequisites

You need to have the following software installed on your system:

1. **Tesseract OCR** - Download from [GitHub](https://github.com/tesseract-ocr/tesseract)
2. **ImageMagick** - Download from [ImageMagick.org](https://imagemagick.org/script/download.php)
3. **AutoHotkey** - Download from [AutoHotkey.com](https://www.autohotkey.com/)

## Installation

1. Install Tesseract OCR and ImageMagick on your computer
2. Download or clone this repository
3. Run `OCR4.2.ahk` with AutoHotkey
4. On first run, you'll be prompted to locate:
   - `tesseract.exe` (Tesseract OCR executable)
   - `magick.exe` (ImageMagick executable)
5. The script will save these paths as environment variables for future use

## Usage

1. Take a screenshot or copy an image to your clipboard using any screenshot tool
2. Press `Ctrl+Shift+V` to perform OCR
3. The extracted text will be automatically copied to your clipboard
4. A notification will confirm the operation was successful

## How It Works

1. Captures the image from your clipboard
2. Preprocesses the image with ImageMagick to improve OCR accuracy:
   - Increases resolution (density 300, resize 150%)
   - Converts to grayscale
   - Corrects skew (deskew)
   - Enhances contrast and sharpness
   - Removes noise (despeckle)
3. Runs Tesseract OCR with optimized parameters
4. Copies the extracted text to your clipboard
5. Cleans up temporary files

## Environment Variables

The script uses the following environment variables:

- `TESSERACT_PATH` - Path to tesseract.exe
- `MAGICKBUBBSY` - Path to magick.exe

These are set automatically on first run and persist across sessions.

## License

Open source - feel free to use and modify as needed.
