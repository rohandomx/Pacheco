# José Emilio Pacheco's Literary Correspondence @ Princeton University

Pacheco's literary correspondence is deposited in **52 folders distributed in 10 boxes**. The first nine boxes are identified with the prefix 129-  (1291-1299); the last box is identified by the number 1300.

In 'Distribution.csv', each folder's name includes the box number and the file number separated by an underscore (i.e., 1291_1).

## 'count.py' 

is a Python script that reads the Excel workbooks that mirror the internal organization of each box with its respective folders. The script creates a formatted Excel workbook that shows the total of letters and names in each box and creates a chart with the most recurrent authors throughout the collection. This report helps clean the data and find recurrences that may have been registered with different spellings. It also provides the general measures of the archive for each box and, in total: 4,146 letters exchanged with more than 1,600 recipients.

## 'load.py' 

is another Python script that incorporates the metadata from the Excel workbooks into a relational database hosted in SQLite. After executing a query, the database shows the distribution of letters grouped by name and the number of letters per folder ('count') as well as a distinction between "sender" and "receiver" ('direction') in the .csv file **“Distribution”**.

*More data analysis and text mining are in pending to create a visualization of the most recurrent names in the collection.*

