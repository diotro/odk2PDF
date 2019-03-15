#!/usr/bin/env python
# coding: utf-8
#------------------------------------------------------------------------------------------------------------------------------------------#
#  va2PDF -- Create human-readable form of WHO verbal autopsy questionnare                                                                 #
#  Copyright (C) 2018 Jason Thomas, Samuel Clark, Martin Bratschi in collaboration with the Bloomberg Data for Health Initiative          #
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
#----------------------------------------------------------------------------

from pylatex import Document, Section, Subsection, Itemize
from pylatex.utils import bold, italic
import os
import pandas as pd
import numpy as np
import argparse

# retrieve and set arguments
parser = argparse.ArgumentParser()
parser.add_argument("fileName", help="file name of the VA questionnare (expected .xlsx format)")
parser.add_argument("-v", "--verbose", help="include explanations for each question be included (if present)",
                    action="store_true")
args = parser.parse_args()

questionnaire = args.fileName

# ### Read Files into dataframe
# df1 = survey sheet
# df2 = choices sheet

df1 = pd.read_excel(questionnaire, 'survey')
df2 = pd.read_excel(questionnaire, 'choices')


# ### Parse Question, ID, and Revelance
df1.loc[(df1['type'] != "begin group") & (df1['type'] !=  "end group"), 'Question'] = df1['label::English']
df1.loc[(df1['type'] != "begin group") & (df1['type'] !=  "end group"), 'ID'] = df1['name']
df1.loc[(df1['type'] != "begin group") & (df1['type'] !=  "end group"), 'Relevance'] = df1['relevant']


# ### Make lines 1 and 2 into separate columns in dataframe
# line1 = (ID#):(Relavance)
# line2 = Question: (question)


df1 = df1.fillna('')

df1['line1'] = df1['ID'].astype(str) + ": " + df1['Relevance'].astype(str)
df1['line2'] = "Question: "+ df1['Question'].astype(str)


# ### Make line 3 into separate column in dataframe
# line3 = (answer selections)

df2 = df2[['list name', 'name']]

selection = df2['list name'].unique()[1:]
selection = selection.tolist()

df2 = df2.groupby(['list name'])['name'].apply(lambda x: "\n ".join(x.astype(str))).reset_index()

dict = df2.groupby('list name')['name'].apply(list).to_dict()

df1.loc[df1['type'] == 'text', "line3"] = "text"
df1.loc[df1['type'] == 'time', "line3"] = "time"
df1.loc[df1['type'] == 'date', "line3"] = "date"
df1.loc[df1['type'] == 'calculate', "line3"] = "calculate"
df1.loc[df1['type'] == 'integer', "line3"] = "integer"
df1.loc[df1['type'] == 'note', "line3"] = "note"

for i in range(0, 37):
    df1.loc[df1['type'].str.contains(selection[i]), "line3"] = dict[selection[i]]

df1['line3'] = df1['line3'].fillna('')


# ### Clean Columns

df1 = df1[(df1.type != "begin group") & (df1.type != "end group") & (df1.type != "")]
df3 = df1[['line1','line2', 'line3']]


# ### Write to $\LaTeX$

geometryOptions = {"margin": "1in"}
doc = Document(geometry_options=geometryOptions, indent=False)

n = 1
for row in df3.iterrows():
    doc.append(bold(row[1][0] + "\n"))
    doc.append(row[1][1] + "\n")
    with doc.create(Itemize()) as itemize:
        itemize.add_item(italic(row[1][2] + "\n"))
    n += 1

doc.generate_pdf("ODK2PDF_Output", clean_tex=False)

