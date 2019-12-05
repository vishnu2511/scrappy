from urllib.request import urlopen
from bs4 import BeautifulSoup
from heapq import nlargest
from pyexcel_xls import get_data
import sqlite3
import xlsxwriter
conn = sqlite3.connect('countD.db')
conn.execute('''CREATE TABLE COUNTN
              (WORD TEXT NOT NULL,
             URL TEXT NOT NULL,
              COUNTS INT NOT NULL);''')

data = get_data("foger.xlsx")
D = dict(data)
workbook = xlsxwriter.Workbook("shiraz.xlsx")
#print(D)
k = D.values()
#print(k)
l = list(k)
count=1
for i in l:
  for j in i:
    for k in j:
      url= k
      file_handle = urlopen(url)
      name="Sheet"+str(count) 
      worksheet = workbook.add_worksheet(str(name))
      count += 1
      worksheet.write("A1",k)
      worksheet.write("A3","WORDS")
      worksheet.write("B3","COUNT")
      chart = workbook.add_chart({'type': 'column'})
      
      
#print(file_handle.read())
      soup = BeautifulSoup(file_handle,"html.parser")
      for script in soup(["script","style"]):
            script.extract()
      text = soup.get_text()
      t = text.lower()
      #print(t)
      a = t.split()
      s=set(a)
      h = {'to', 'in', 'at', 'and', 'the','for', 'not', 'by', 'on','but','of','or','a','is','was','where','were','about','with','&','this','that','then'}
      o = s - h
#print(a)
      d = {}
      for i in o:
          d[i]=a.count(i)
          
      #print(d)
#top = 1    
#for k,v in d.items():
        #if v > top:
         #  top = v
           #print(k, '=', top) 
      five_largest = nlargest(5, d, key=d.get)
      n = 3
      for j in five_largest:
           #print(i,'=',d[i])
            conn.execute('''INSERT INTO COUNTN(WORD,URL,COUNTS)VALUES(?,?,?)''',(j,k,d[j]));
            conn.commit()
            worksheet.write(n,0,j)
            worksheet.write(n,1,d[j])
            n += 1
      area="="+str(name)+"!$B$4:$B$8"
      cate="="+str(name)+"!$A$4:$A$8"
      chart.add_series({'categories':cate,'values':area})
      worksheet.insert_chart('E5', chart)
                       
c = conn.execute("select WORD,URL,COUNTS FROM COUNTN")
for row in c:
  print("WORD:",row[0])
  print("URL:",row[1])
  print("COUNTS:",row[2])
workbook.close()

    

      
   
   




  
