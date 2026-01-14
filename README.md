Script to aid in creation of archival labels from a .xlsx spreadsheet. It takes text from columns of a spreadsheet and places them on a label in the desired order, with a few formatting options.

This code does not adhere well to guidelines and is in a single file. It does its job.

# Installation

If at the APS, copy P:\labelmaker to your computer. This contains a copy of the executable file, for Windows, generated last on 1/14/2026. Please do not modify P:\labelmaker if you are just using it.

For Python/development purposes: Clone this repo and make a virtual environment for Python with requirements.txt

# Basic usage

1. Make a spreadsheet (.xlsx) with the content of your labels, in the same directory as `labelmaker.py`/`labelmaker.exe`. Each row will become one label. There should be no column headers.
2. Edit the file `config-content.ini` in a text editor. Modify at minimum: `label-maker-to-use` and `spreadsheet`, and usually the content of each line. Here you can also modify the configuration for your label content.
3. Either double-click `labelmaker.exe` or run `python labelmaker.py`. If successful, you will get a PDF file in the same directory.
4. Print labels. Use a test page first, and ensure no margins: set scaling to 100% or 'actual size', and check that the printer is not adding its own margins.

# Detailed usage

## Getting data to use
ArchivesSpace's Bulk Update Spreadsheet can extract metadata from a collection.

## Configuring the text and formatting in a label
`config-content.ini` describes the content and formatting of text on labels to be produced.

To make a new label template, it's easiest to copy an existing one and paste it at the bottom of the document.

- Under `main settings`:
    - `label-type-to-use`: the name of the label type, as written in square brackets further down, e.g. `folder1`
    - `spreadsheet`: the name of your spreadsheet, with a .xlsx extension, e.g. `example.xlsx`
- Under a label type, e.g. `[folder1]`:
    - `layout`: the layout to use, by default `box` or `folder`, referencing `config-layout.ini`
    - `font-face`: name of font, either `Arial` or `Garamond`
    - `font-padding`: padding between each line, recommended as `1.5` for `Arial` and `1` for `Garamond`
    - `font-normal`: size of font if small is not invoked; recommended font sizes are no smaller than 8 and no bigger than 12
    - `font-small`: size of font if small is invoked
    - `trailing-ellipsis`: whether an overlong line will cut off with an ellipsis, either `yes` (recommended) or `no`
    - `overflow-into-spaces`: whether overlong text will try to overflow into available gaps, either `yes` (recommended) or `no`
    - `line1` onwards: the content of the label, broken down by line. You can leave lines blank for spacing, which looks nice on boxes. If you try to add more lines that can fit according to your font size and label space, you will get an error. This has two parts:
        - column number from the spreadsheet to put in this line, starting with `1`
        - optional formatting: `b` (bold), `i` (italic), `s` (small)

## Configuring the label layout
You probably will never touch this unless you have new labels to use. Each label type is defined in square brackets and is what is referenced in `config-content.ini`'s `label-type-to-use` parameter.

The main reason to change this file is to adjust `gap-print-error-adjust`. This applies an offset to both X and Y axes, accounting for the margin errors that your printer may have. If text is printing too close to the label edge still, you can increase this number. If this is forcing your text to go too far from the label margins, you can decrease this number.

The second reason to change this file is to adjust `label-internal-padding`. This creates the internal margin for the label. With small labels, the default here may be too large. The program can under-estimate the number of lines that are actually possible to write, so decreasing this number may improve this.

You can define a new layout by copying all the parameters from one of the existing ones, and changing the word in square-brackets. If `use-logo` = yes, you also need `logo-width` and `logo-height` parameters. There’s a lot of trial and error in defining a layout. Start with what the official label template suggests and make adjustments. All measurements are in inches.

# Maintenance

## Authorship
Written by Paul Sutherland, American Philosophical Society, 2024. Adapted for Github January 2026.

## Executable file for distribution
(Re)building an executable file for Windows is done using the PyInstaller library (not required otherwise to run this software). CD to directory and run command `python -m PyInstaller -F labelmaker.py`, and the .exe should appear in ./dist. Distribute a folder containing:
- labelmaker.exe
- *.ttf font files (x8)
- *.ini files (x2)
- logo.png
After this is done, you can delete the files and folders PyInstaller made (./dist, ./build, labelmaker.spec).

## Bugs
- labelmaker.exe was too big for Github so this export lives on our servers.
- text truncated in the middle of two quotations, e.g. "quoted folder name that contains "additional quote"", will only add one quote.

## Improvements
- better code
- turn quotes checking into a single function that receives the last character, especially if adding more possible quotes
- format: underline [u] (this would require using ReportLab's paragraph module)
- format: right-align [r] (this would probably not be that difficult)
- format: multiple columns on a line [|] (this would almost definitely require paragraph module)