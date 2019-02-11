#------------------------------------------------------------------------------------------------------------------------------------------#
#  va2PDF -- Create human-readable form of WHO verbal autopsy questionnare                                                                 #
#  Copyright (C) 2018  Jason Thomas, Samuel Clark, Martin Bratschi in collaboration with the Bloomberg Data for Health Initiative          #
#                                                                                                                                          #
#  This program is free software: you can redistribute it and/or modify                                                                    #
#  it under the terms of the GNU General Public License as published by                                                                    #
#  the Free Software Foundation, either version 3 of the License, or                                                                       #
#  (at your option) any later version.                                                                                                     #
#                                                                                                                                          #
#  This program is distributed in the hope that it will be useful,                                                                         #
#  but WITHOUT ANY WARRANTY; without even the implied warranty of                                                                          #
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                                                                           #
#  GNU General Public License for more details.                                                                                            #
#                                                                                                                                          #
#  You should have received a copy of the GNU General Public License                                                                       #
#  along with this program.  If not, see <http://www.gnu.org/licenses/>.                                                                   #
#                                                                                                                                          #
#------------------------------------------------------------------------------------------------------------------------------------------#
from pylatex import Document, Section, Subsection, Itemize
from pylatex.utils import bold, italic
import os
import pandas as pd
import argparse

# retrieve and set arguments
parser = argparse.ArgumentParser()
parser.add_argument("fileName", help="file name of the VA questionnare (expected .xlsx format)")
parser.add_argument("-v", "--verbose", help="include explanations for each question be included (if present)",
                    action="store_true")
args = parser.parse_args()

questionnaire = args.fileName

# read in questionnare
df = pd.read_excel(questionnaire, sheet_name="survey")
colNames = list(df)
if args.verbose:
    if "explanation" not in colNames:
        print("Error: --verbose option requires that the input file include a column named 'explanation'.")
        sys.exit(1)
    else:
        explanationIndex = colNames.index("explanation")

# set up LaTeX document
geometryOptions = {"margin": "1in"}
doc = Document(geometry_options=geometryOptions, indent=False)

# walk through .xlsx file and add questions to LaTeX doc
n = 1
for row in df.iterrows():
  if row[1][0] not in ("begin group", "end group", "calculate", "note", "integer", "time", "", "\n") and not pd.isna(row[1][0]):
    cleanText = "".join(row[1][2].splitlines())
    doc.append(bold("Question "))
    doc.append(bold(n))
    doc.append(":  " + cleanText + "\n")
    if args.verbose:
        if not pd.isna(row[1][explanationIndex]):
            with doc.create(Itemize()) as itemize:
                itemize.add_item(italic("Explanation: " + row[1][explanationIndex] + "\n"))
    n += 1

doc.generate_pdf("va2PDF_Output", clean_tex=False)
