# OCR and Text Extraction from Scanned PDFs

## Overview
This project extracts text from scanned PDF documents using EasyOCR and PyMuPDF. It enhances scanned PDFs by adding a selectable text layer and exports the extracted content into an Excel file for further analysis.

## Features
- Converts scanned PDFs into images.
- Extracts text from images using EasyOCR.
- Adds text layer to the original PDF, making it searchable.
- Exports extracted text to an Excel file for structured analysis.

## Technologies Used
- Python
- EasyOCR
- PyMuPDF (fitz)
- pdf2image
- OpenCV
- NumPy
- Pandas

## Installation
1. Clone the repository:
   ```sh
   git clone https://github.com/Kurama-90/OCR-and-Text-Extraction-from-Scanned-PDFs.git
   cd OCR-and-Text-Extraction-from-Scanned-PDFs
   ```
2. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```
3. Ensure that you have `poppler` installed (+Add PATH) for `pdf2image` to work:
   - Linux: `sudo apt install poppler-utils`
   - MacOS: `brew install poppler`
   - Windows: Download from [here](https://github.com/oschwartz10612/poppler-windows/releases)

## Usage
1. Place your scanned PDF in the `uploads` folder and rename it to `scanned_document.pdf`.
2. Run the script:
   ```sh
   python main.py
   ```
3. The processed PDF with a selectable text layer will be saved in the `outputs` folder as `output.pdf`.
4. The extracted text will be saved in `output.xlsx`.

## File Structure
```
📂 project-directory
 ├── 📂 uploads  # Place scanned PDFs here
 ├── 📂 outputs  # Processed files will be saved here
 ├── main.py  # Main script for processing PDFs
 ├── requirements.txt  # Dependencies list
 ├── README.md  # Project documentation
 ├── LICENSE  # MIT License file

```

## Contributing
Feel free to contribute by submitting issues or pull requests.

## License
This project is licensed under the MIT License.

## Author
[Kurama-90](https://github.com/yourusername)

