'需要安装mysql-connector-odbc-8.4.0-winx64.msi : https://downloads.mysql.com/archives/get/p/10/file/mysql-connector-odbc-8.4.0-winx64.msi
strConn="Driver={MySQL ODBC 8.4 Unicode driver};SERVER=127.0.0.1;port=3306;UID=root;PWD=root;DATABASE=company;"
Set conn = CreateObject("ADODB.Connection")

conn.Open strConn
' add data
' strsql = "insert into admin values('10','Kenny','12345');"
' update data
' strsql = "update admin set userName='Ken' where id=8;"
strsql = "update admin set id=9 where id=10;"
' delete data
' strsql = "delete from admin where id=9;"
conn.Execute(strsql)
conn.Close

Set conn = Nothing



