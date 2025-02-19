<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>создание фиктивн таблицы для анализа кто просматривае базы</title>
</head>
<BODY bgcolor=LightCyan>
<!--#INCLUDE FILE="create_stat_ip_new.inc"-->  
<%
Tim_start=Time
'Response.Write "Довлетбаев З.С.<BR>"
'Response.Write "2022.07.05  <BR>"
'Response.Write "создание фиктивной таблицы  <BR>"
call Create_ZAD("stat_ip_new","stat_ip","Dat <='2022.01.01'")
'drop table stat_ip_new;
'create Table stat_ip_new select * from zoma.stat_ip where Dat >= '2022.04.07';

db.Close
Set db = Nothing
Response.Write "<HR>"
Response.Write "время работы " & mid(Time-Tim_start,1,4) & " сек.<BR>"
%>
</BODY>
</HTML>
