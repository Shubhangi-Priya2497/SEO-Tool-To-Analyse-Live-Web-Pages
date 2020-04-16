#http://www.python.org/
#import file needed for every factor in python
from urllib.request import urlopen
from bs4 import BeautifulSoup
import re
import sqlite3
import json
import xlsxwriter
import operator
#using webscrapping to recover data from any website
url=input("Which page you would like to check?Enter the full url:")#to input the url through user

#handling HTTPError exception
try:
    html=urlopen(url)
except HTTPError as e:#if the url is wrong the exception will show
        print(e)
data=BeautifulSoup(html,"html.parser")#create a new bs4 object from the html data loaded
for script in data(["script","style"]):#remove all javascript and stylesheet code
    script.extract()#get text
text=data.get_text()#break into line and remove leading and trailing space into each

#To remove space between Paragraph
space=[" ","/n"]
for w in space:
    if w in text:
        text=text.replace(w," ")

lines=(line.strip() for line in text.splitlines())

#handling FileNotFoundError
try:
    f=open('web.txt','w')
except(FileNotFoundError):
    print("File Not Found")
    
#Entering the text using webscrapping
for sline in lines:
    f.write(sline+"\n")
    
f.close()#close the 'f' dile to open it in reading mode

#handling FileNotFoundError
try:
    file=open('stop_words.txt','r')
except(FileNotFoundError):
    print("File Not Found")
#read the content of the file 'file' and converting it into list
str=file.read()
l1=str.split()

#read the content of the file 'file' and converting it into list
try:
    f=open('web.txt','r')
except(FileNotFoundError):
    print("File Not Found")
str1=f.read()
sentence=str1.split()#[w for w in lines if not w in str]

#removing stop words
for w in l1:
    if w in sentence:
        continue
sentence.sort()
sentence=" ".join(sentence)
sentence=sentence.split(" ")
#print(sentence)

#Creating a dictionary and database to save the count and the words 
conn=sqlite3.connect('Mydict.db')
conn.execute('''drop table if exists Counting''')
conn.execute('''CREATE TABLE Counting(Website TEXT NOT NULL,Word TEXT NOT NULL,Count INT NOT NULL);''')
#L=sentence.split(' ')
D={}
for w in sentence:
    c=sentence.count(w)
    #print(w,c)
    D[w]=c
    #print(D)
for w,c in D.items():
    conn.execute('''INSERT INTO Counting(Website,Word,Count)VALUES(?,?,?)''',(url,w,c))
    
conn.commit()
cursor=conn.execute('''SELECT Website,Word,Count FROM Counting ORDER BY Count DESC;''')
for row in cursor:
    print('Website:',row[0])
    print('Words:',row[1])
    print('Count:',row[2])
print("-"*70)    

#Creating a dictionary to save top ten word of the website
D2={}
D2= dict(sorted(D.items(), key=operator.itemgetter(1), reverse=True)[:10])
print("Top ten words of the website is:")
print(D2)

# Create a workbook add a worksheet.
workbook = xlsxwriter.Workbook('Count01.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#Create A new Chart Object
chart = workbook.add_chart({'type': 'line'})

# Write some data headers.
worksheet.write('A1', 'Words', bold)
worksheet.write('B1', 'Count', bold)

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for word, count in (D2.items()):
        worksheet.write(row, col,     word)
        worksheet.write(row, col + 1, count)
        row += 1

# Add a series to the chart.
chart.add_series({'values': '=Sheet1!$A$2:$A$10'})
chart.add_series({'values': '=Sheet1!$B$1:$B$10'})

# Insert the chart into the worksheet.
worksheet.insert_chart('C1', chart)
workbook.close()


