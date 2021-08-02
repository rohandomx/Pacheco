# José Emilio Pacheco's Literary Correspondence @ Princeton University Library

Pacheco's literary correspondence is deposited in **52 folders distributed in 10 boxes**. The first nine boxes are identified with the prefix 129-  (1291-1299); the last box is identified by the number 1300.

In **'Distribution.csv'**, each folder's name includes the box number and the file number, separated by an underscore (i.e., 1291_1).

## 'count.py' 

is a Python script that reads the Excel workbooks that mirror the internal organization of each box and the folders within it. The script generates a formatted Excel workbook that displays the total count f letters and names of correspondents in each box. It also creates a chart of the most recurrent authors throughout the collection. This report helps to clean the data and identify recurrences that may have been registered with different spellings. It also provides a general estimate of the archive’s size, both in each individual box and in total: 4,146 letters exchanged with more than 1,600 recipients.

## 'load.py' 

is another Python script that incorporates the metadata from the Excel workbooks into a relational database hosted in SQLite. After executing a query, the database shows the distribution of letters grouped by name and the number of letters per folder ('count') as well as the distinction between "sender" and "receiver" ('direction') in the .csv file **“Distribution”**.


*More data analysis and text mining are pending and will yield a visualization of the most frequently recurring names in the collection.*

