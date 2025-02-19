<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>удаление таблицы </title>
</head>
<BODY bgcolor=LightCyan>
<!--#INCLUDE FILE="create_stat_ip_new.inc"-->  
<%
Tim_start=Time
'Response.Write "Довлетбаев З.С.<BR>"
'Response.Write "2022.07.18  <BR>"
'Response.Write "удаление записей из таблицы согласно заданному условию <BR>"
call Delete_ZAD("stat_ip","Dat <='2021.01.01'")

db.Close
Set db = Nothing
Response.Write "<HR>"
Response.Write "время работы " & mid(Time-Tim_start,1,4) & " сек.<BR>"
%>
</BODY>
</HTML>
