from tabula import read_pdf, convert_into, convert_into_by_batch
import pandas as pd
import numpy as np
import os

from progressbar import ProgressBar\

pbar = ProgressBar()

while True:
    directory = input("Enter your input path(e.g <full_path>\\<folder name>): ")
    print("You entered input path " + directory)
    if directory:
        break
while True:
    output_path = input("Enter your output path(e.g <full_path>\\<folder name>\\): ")
    print("You entered output path " + output_path)
    if output_path:
        break
while True:
    output_file = input("Enter your filename (e.g. <filename>.xlsx): ")
    print("You entered filename " + output_file)
    if output_file:
        break

# directory = "C:\\Users\\z1MMB2712\\PycharmProjects\\pdftoexcel\\test"
# output_path = "C:\\Users\\z1MMB2712\\PycharmProjects\\pdftoexcel\\output\\"
# output_file = "finalresult.xlsx"
finalDf = pd.DataFrame()
fileCount=0
for filename in pbar(os.listdir(directory)):
    if filename.endswith(".pdf") or filename.endswith(".PDF"):

        listpdf = read_pdf(os.path.join(directory,filename), pages='all', multiple_tables=True)
        itercars = iter(listpdf)
        # ignore the first table that is consumed
        next(itercars)
        i=0
        df = pd.DataFrame()
        for car in itercars:
            X = np.array(car[1:])
            # prepare data
            col1data = np.empty(2)
            for row in X[1:, 0]:
                col1data = np.vstack((col1data, row.split()))
            new_data = np.hstack((col1data[1:, :], X[1:, 1:]))
            # prepare header
            Xheader = X[0, 0:]
            columnH1 = Xheader[0].split()
            column1 = columnH1[0] + " " + columnH1[1]
            column2 = columnH1[2]
            columnHeader = [column1, column2, Xheader[1]]
            columnHeader.extend(Xheader[2::])

            if i==0:
                df = pd.DataFrame(new_data,columns = columnHeader)
            else:
                new_df = pd.DataFrame(new_data, columns= columnHeader)
                df=df.append(new_df)
            i+=1
        df = df.reset_index()
        df['filename'] = filename
        if fileCount==0:
            finalDf = df
        else:
            finalDf = finalDf.append(df)
    else:
        continue
    # print("%d files" % fileCount)
    fileCount+=1
finalDf.to_excel(os.path.join(output_path, output_file))
print("Job Completed")