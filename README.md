KDP Word Search Puzzle Generator
Generate high-quality word search puzzle books (PDF and PPTX) from your own custom word lists, ready to publish on Amazon KDP. Easily customize grid size, difficulty levels, and standard KDP page dimensions through a simple Streamlit interface.

Features
User-friendly Streamlit web app for uploading Excel files and configuring puzzles

Export puzzles as print-ready PDFs or fully editable PowerPoint presentations (PPTX)

Supports official KDP page sizes for seamless publishing

Adjustable difficulty: easy, medium, or hard (controls word placement directions)

Automatically generates solution pages arranged neatly on the PDF/PPTX

Supports multiple themes where each puzzle can have its own word list title

Uses custom embedded fonts for professional print quality

Installation
Clone the repository:

bash
git clone https://github.com/yourusername/kdp-wordsearch-generator.git
cd kdp-wordsearch-generator
Install dependencies:

bash
pip install -r requirements.txt
Key dependencies are:

streamlit

pandas

python-pptx

reportlab

Ensure the font files (RobotoMono-Regular.ttf and RobotoMono-Bold.ttf) are located in the project root directory.

Quickstart
Run the app:

bash
streamlit run app.py
Upload your Excel file:

Column A: Theme (must repeat for each word)

Column B: Word

No header row required.

Select grid size, difficulty level, and output format (PDF or PPTX).

Click Generate and download your ready-to-publish puzzle book.

Example Excel format
Theme	Word
ANIMALS	Cat
ANIMALS	Dog
FRUITS	Apple
FRUITS	Orange
Advanced Configuration
Adjust grid size between 10x10 and 30x30 for custom challenge levels.

Large word lists are split into multiple puzzles automatically.

The app informs if any words could not fit the grid at the selected difficulty.

Contributing
Contributions are welcome! Please fork the repo, create a feature branch, and submit a pull request. Follow standard GitHub contribution guidelines.

License
MIT License. See LICENSE file for details.

FAQ
Is this tool language-agnostic?
Yes, it supports any language using standard letters. Accented characters might need different fonts.

Is the PPTX fully editable?
Yes, you can edit the puzzles, fonts, and layout using any PowerPoint-compatible editor.

Are all KDP page sizes supported?
Common sizes are included. You can add more sizes to the TAMAÑOS_KDP dictionary as needed.
