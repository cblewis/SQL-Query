import pyodbc
import csv

#Masterfully crafted by Cameron Lewis

connection = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=ODSUNITEDDBLS;"
                      "Database=XLHEALTH;"
                      "Trusted_Connection=yes;")

cursor = connection.cursor()
cursor.execute('''
SELECT m.companyid, JN = i.JobNumber, Start_ImpID = MIN(i.ImportFileID), End_ImpID = MAX(i.ImportFileID), MemberCount=count (*)
FROM Members m
INNER JOIN Member_Raws r ON r.MemberRecID = m.MemberRecID
INNER JOIN ImportFile i ON i.ImportFileID = m.ImportFileID AND i.ImportDate >= convert(varchar(10),CAST(CONVERT(varchar,GetDate(),23) as DateTime),112)
INNER JOIN ImportFile oi ON oi.ImportFileID = m.oImportFileID AND oi.ImportDate >= convert(varchar(10),CAST(CONVERT(varchar,GetDate(),23) as DateTime),112)
WHERE oi.ImportFileID <> 0
GROUP BY m.companyid, i.JobNumber, i.ImportFileID, i.ImportDate
ORDER BY i.jobnumber, m.companyid
''')
result = cursor.fetchall()

column_names = [i[0] for i in cursor.description]
fp = open("UHCDailyJobs.csv", "w")
myFile = csv.writer(fp, lineterminator = '\n')
myFile.writerow(column_names+["Status"])
myFile.writerows(result)
fp.close()
#print('result = %r' % (result))
