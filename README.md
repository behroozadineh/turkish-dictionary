# Turkish Dictionary

A desktop application for translating Turkish words and phrases to Persian or English. Saves data to Excel file.

Two versions available:
- turkish_persian.py - Turkish to Persian
- turkish_english.py - Turkish to English

## Features

- Turkish special letters: ç, ğ, ı, ö, ş, ü, Ç, Ğ, İ, Ö, Ş, Ü
- Add custom sources (Book, YouTube, or your own)
- Duplicate word detection
- Edit existing entries
- Export to Excel

## How to Use

1. Enter Turkish word or phrase
2. Use the letter buttons for Turkish special characters
3. Add pronunciation, example, synonym, antonym (optional)
4. Enter Persian or English meaning
5. Select a source (Book or YouTube) or add a new source
6. Add source description if needed
7. Click "Send" (English version) or "ارسال کردن" (Persian version)
8. Review preview window
9. Click "Confirm and Save" (English) or "تأیید و ذخیره" (Persian)

## Edit Existing Words

If you enter a word that already exists in the dictionary:
- Click "Edit existing one" (English) or "ویرایش کلمه موجود" (Persian)
- Make your changes
- Click Send again - it will update instead of creating a duplicate

## File Structure

turkish-dictionary/
├── turkish_persian.py    # Turkish to Persian version
├── turkish_english.py    # Turkish to English version
├── requirements.txt      # Python dependencies
├── README.md            # This file
├── Dict-TU-FA.xlsx      # Excel file for Persian version (auto-created)
└── Dict-TU-EN.xlsx      # Excel file for English version (auto-created)

## Excel File Columns

| Column | Content |
|--------|---------|
| 1 | Row number |
| 2 | Date |
| 3 | Turkish word or phrase |
| 4 | Pronunciation |
| 5 | Example |
| 6 | Synonym |
| 7 | Antonym |
| 8 | Source |
| 9 | Meaning (Persian or English) |

## Installation

pip install openpyxl

python turkish_persian.py   # for Persian version
python turkish_english.py   # for English version

## Requirements

- Python 3.7+
- openpyxl

## Author

Behrooz Adineh

## License

MIT License

## Contributing

Issues and pull requests are welcome.
