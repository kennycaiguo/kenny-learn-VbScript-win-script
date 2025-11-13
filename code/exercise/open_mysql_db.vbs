'需要安装mysql-connector-odbc-8.4.0-winx64.msi : https://downloads.mysql.com/archives/get/p/10/file/mysql-connector-odbc-8.4.0-winx64.msi
strConn="Driver={MySQL ODBC 8.4 Unicode driver};SERVER=127.0.0.1;port=3306;UID=root;PWD=root;DATABASE=company;"
Set conn = CreateObject("ADODB.Connection")

conn.Open strConn

strsql = "select * from admin"
Set rs = CreateObject("adodb.recordset")
rs.Open strsql,conn,1,3
 WScript.Echo "id  " & "  " & "userName" & " " & "pwd" 
Do Until rs.EOF
  WScript.Echo rs.Fields("id") & "     " & rs.Fields("userName") & "   " & rs.Fields("pwd")
  rs.MoveNext
Loop

rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing

'增、删、改
'conn.execute strSql,intRowsAffect

