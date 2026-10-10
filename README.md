# Quiz Excel Formatter v9

This version uses the **three supplied standard XLSX files as the actual Excel templates**.

## Standard templates

- `templates/LIVEQZ-RO-TT-T2-S1-V1.xlsx`
- `templates/TTTQZ-RO-TT-T2-S1-V1.xlsx`
- `templates/GAMENAME-RO-TT-T2-S1-V1.xlsx`

The generated workbook keeps the template's:
- fonts
- fills
- borders
- alignments
- column widths
- sheet names
- Help sheet
- workbook structure

Only the question data is inserted.

## Tic Tac Toe

Accepts:
- 9 questions
- 18 questions
- 27 questions

Difficulty is filled **only when there are 27 questions**:
- Q1–Q9 → 1
- Q10–Q18 → 2
- Q19–Q27 → 3

For 9 or 18 questions, the Difficulty column remains empty.

## Images

Images are optional.

Select PNG files only:
- Q1.png
- Q2.png
- Q5.png

The matching filename is placed in the Image column. If no image is selected for a question, the Image cell stays empty.

## File names

Live:
`LIVEQZ-RO-TT-T2-S14-V1.xlsx`

Tic Tac Toe:
`TTTQZ-RO-TT-T2-S14-V1.xlsx`

LMS:
`GAMENAME-RO-TT-T2-S14-V1.xlsx`

## GitHub Pages

Upload:
- `index.html`
- `style.css`
- `script.js`
- `templates/` folder with all three XLSX files

The application uses ExcelJS from a CDN to load the supplied template and generate the final XLSX in the browser.
