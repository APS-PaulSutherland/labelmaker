Script to aid in creation of archival labels from a .xlsx spreadsheet. It takes text from columns of a spreadsheet and places them on a label in the desired order, with a few formatting options.

 `config-layout`

# Basic usage

1. Make a spreadsheet (.xlsx) with the content of your labels, in the same directory as `labelmaker.py`. Each row will become one label. There should be no column headers.
2. Edit the file `config-content.ini` in a text editor. Modify at minimum: `label-maker-to-use` and `spreadsheet`, and usually the content of each line. Here you can also modify the configuration for your label content.
3. Either double-click `labelmaker.exe` or run `python labelmaker.py`. If successful, you will get a PDF file in the same directory.
4. Print labels. Use a test page first, and ensure no margins: set scaling to 100% or 'actual size', and check that the printer is not adding its own margins.

# Detailed usage

## Getting data to use
ArchivesSpace's Bulk Update Spreadsheet can extract metadata from a collection.

## Configuring the text and formatting in a label
`config-content.ini` describes the content and formatting of text on labels to be produced.
- Under `main settings`:
    - `label-type-to-use`: the name of the label type, as written in square brackets further down, e.g. `folder1`
    - `spreadsheet`: the name of your spreadsheet, with a .xlsx extension, e.g. `example.xlsx`
- Under a label type, e.g. `[folder1]`:
    - `layout`: the layout to use, by default `box` or `folder`, referencing `config-layout.ini`
    - `font-face`: name of font, either `Arial` or `Garamond`