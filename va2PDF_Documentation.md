% **Documentation:** `va2PDF.py`
% Jason Thomas, Samuel Clark, & Martin Bratschi \
  in collaboration with the Bloomberg Data for Health Initiative
% May 10, 2018


---
header-includes:
 - \usepackage{fvextra}
 - \DefineVerbatimEnvironment{Highlighting}{Verbatim}{breaklines,commandchars=\\\{\}}

geometry:
 - margin=1in
---

# Intro
The purpose of the program `va2PDF.py` is to convert a VA questionnaire, in Microsoft Excel format, into a human-readable pdf.  The VA
questionnaire must have the same format as the
[WHO 2016 VA questionnaire](http://www.who.int/healthinfo/statistics/verbalautopsystandards/en/).  Users can add an additional column to
the spreadsheet with the name 'explanation' and the va2PDF.py program will include the text, intended to be an explanation of the
particular question.  The `va2PDF.py` program is implemented in **Python 3** and requires the **pandas** and **pylatex** modules.

# Set up
- Download and install [Python 3](https://www.python.org/downloads/)
- Install the pandas and pylatex modules with the following command (in a terminal)

    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~{.bash}
    $ python3 -m pip install pandas pylatex --user
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    See the [Python documentation page](https://docs.python.org/3/installing/) for more help on installing modules.

# Usage
- **Example 1**: The WHO 2016 VA questionnaire is used for this example and the MS Excel file was renamed to `q1.xlsx`.  The questionnare
can be converted into a PDF with the following command:

    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~{.bash}
    $ python3 va2PDF.py q1.xlsx
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    The program will create a new PDF file: `va2PDF_Output.pdf`.

- **Example 2**: For this example, the WHO 2016 VA questionnaire is modified to include an additional column entitled "explanation."
The values in this column include text that provide greater detail about the particular question.  This additional information can be added
to the output by passing the `--verbose` option (or simply `-v`) as follows:

    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~{.bash}
    $ python3 va2PDF.py q2.xlsx --verbose
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Again the program will create a new PDF file: `va2PDF_Output.pdf`.
