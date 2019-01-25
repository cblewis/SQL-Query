import pyodbc
import csv

#Masterfully crafted by Cameron Lewis

connection = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=ODSCSDBALIAS;"
                      "Database=CARESOURCE;"
                      "Trusted_Connection=yes;")

cursor = connection.cursor()
cursor.execute('''
SELECT max(m.companyid) as CompanyID, case when max(m.companyid) = 681 then 'CTP' when max(m.companyid) in (683,684) then 'MA' when max(m.companyid) in (674,675,676,677,678,679) then 'MEDCD' when max(m.companyid) in (685,686,687,688) then 'MKPL' when max(m.companyid) in (671) then 'MA_EOB' else 'OTHER' end as 'ProductType',
JN = i.JobNumber, Start_ImpID = MIN(i.ImportFileID), End_ImpID = MAX(i.ImportFileID), i.ImportDate, case when i.importstatus = 1 then 'YES' else 'NO' end as OnQV,
i.QVConfirmDate as Confirmed_By_Client, sum(case when buildid not like '%error%' and buildid not like '%dependent%' then 1 else 0 end) as Exp_Barcodes_Count,
case when sum(case when buildid not like '%error%' and buildid not like '%dependent%' then 1 else 0 end) > 0 and sum(case when odsbarcode is not null then 1 else 0 end) =0 and i.QVConfirmDate is not null and ((convert(varchar, QVConfirmDate, 110) = convert(varchar, getdate(), 110) and datepart(hh,QVConfirmDate)<8)) then '!! Need to Process !!'
when ((convert(varchar, QVConfirmDate, 110) = convert(varchar, getdate(), 110) and datepart(hh,QVConfirmDate)>=8)) then 'GOOD - Can Process Tomorrow' when i.QVConfirmDate is not null and sum(case when odsbarcode is not null then 1 else 0 end) >0 then 'GOOD - Ready'
when  i.QVConfirmDate is null then 'Not Confirmed by client yet' else 'N/A' end as PRINTFILES,
MemberCount=count (*) FROM Members m with (nolock)
INNER JOIN ImportFile i with (nolock) ON i.ImportFileID = m.ImportFileID AND i.ImportDate >= convert(varchar(10),CAST(CONVERT(varchar,GetDate(),23) as DateTime),112)
WHERE i.ImportFileID <> 0
GROUP BY i.JobNumber, i.ImportFileID, i.ImportDate, i.importstatus, i.QVConfirmDate
ORDER BY i.jobnumber, companyid
''')
result = cursor.fetchall()

column_names = [i[0] for i in cursor.description]
fp = open("CareSourceDailyJobs.csv", "w")
myFile = csv.writer(fp, lineterminator = '\n')
myFile.writerow(column_names)
myFile.writerows(result)
fp.close()
print('result = %r' % (result))
