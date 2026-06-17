# PE Uniform Spreadsheet Maker

Static GitHub Pages app for turning student timetable PDFs into a spreadsheet of PE-uniform days.

## What it does

- Accepts one or more timetable PDFs.
- Reads each PDF page as one student timetable.
- Looks for:
  - Outdoor Environmental Studies
  - Physical Education
  - Athletic Development
  - Agriculture
- Creates columns for `MonA` through `FriA` and `MonB` through `FriB`.
- Exports the result as a formatted `.xlsx` workbook or a plain `.csv`.

All processing happens in the browser. The PDFs are not uploaded to a server.

## Publish on GitHub Pages

1. Create a GitHub repository.
2. Upload `index.html`, `styles.css`, and `app.js` to the repository root.
3. In GitHub, open `Settings` > `Pages`.
4. Set the source to your main branch and root folder.
5. Open the GitHub Pages URL after GitHub finishes publishing.

## Notes

The parser is designed for the uploaded sample format: one student per page, a `Timetable For Student Name` heading, and day columns labelled `MonA` to `FriB`.
