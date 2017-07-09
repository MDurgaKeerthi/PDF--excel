#this file extracts first table in the given pdf into excel with 3columns: date, transactional details and amount 

import os
import xlsxwriter
workbook = xlsxwriter.Workbook('results.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0
worksheet.write('A1', 'Date')
worksheet.write('B1', 'Transactional Details')
worksheet.write('C1', 'Amount')

def isnum(num):
   if num == "1" :
      return 1
   if num == "2" :
      return 2
   if num == "3" :
      return 3
   if num == "4" :
      return 4
   if num == "5" :
      return 5
   if num == "6" :
      return 6
   if num == "7" :
      return 7
   if num == "8" :
      return 8
   if num == "9" :
      return 9   
   if num == "0" :
      return 0
   else :
      return 11      
      
def extractnum(line, index):
   while isnum(line[index]) == 11:
      index += 1
   num = 0.0
   num1 = 0.0
   flag = -1
   while line[index] != " ":
      if line[index] == ".":
         flag = 0
         num = num*1.0 + isnum(line[index+1])*0.1 + isnum(line[index+2])*0.01
         break
      if isnum(line[index]) != 11 and flag != 0:
         num = num*10.0 + isnum(line[index])
      index += 1  
      
   if flag == 0:       
      return num
   else: 
      return -1   
   
def extractString(line, start):
   i=start+1
   
   l = len(line)
   str = ""
   num = -1
   while num == -1:
      while isnum(line[i]) == 11 and i< (l-1):
         i = i+1
      
      i1 = i
      i = i-1
      while line[i] == " " and i>=0:
         i = i-1
      
      num = extractnum(line,i1)
      #print num
      
      if num == -1:
         i = i1
         while line[i] != " " and (line[i] == "," or isnum(line[i]) != 11) :
            i = i+1
      
   j=start+1
   while j<=i:
      str = str + line[j]
      j = j+1 
       
   worksheet.write(row,1, str)   
   worksheet.write(row,2, num)
   print num

os.system("pdftotext -layout -l 1 may_cc.pdf")

readfrom = open("may_cc.txt", "r+")


index = 0
for text in readfrom :
   index = index + 1
   if index==49 or index >50:   
      if isnum(text[0]) == 11:        #loops till the end of the table
         break
      j=0
      date = ""
      while j < 9 :
         date = date + text[j]
         j = j+1
      
      #print date   
      row += 1   
      worksheet.write(row,0, date)  
      extractString(text,9)

workbook.close()
readfrom.close()  

print 'check the results.xlsx file' 
