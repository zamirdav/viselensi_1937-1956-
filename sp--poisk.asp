
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>поиск ALL</title>
</HEAD>

<!--#INCLUDE FILE="sp--poisk.inc"--> 
<%
' 2022.08.26 переставил реквизиты местами
'2022.08.12 Or F1="" поставил для статистики строка 40 для внутреннего вложения 
'
'
Tim_start=Time
'Response.Write "Довлетбаев З.С.<BR>"
'Response.Write "2022.01.24<BR>"
'27.11.2022 добавил поиск по ОСК О3



T = Request.QueryString("T") 

if T="1" then
   if F1="" then 
      Response.Write "без фамилии искать не буду !<BR>"
      Break
   end if
   Call POISK_ZAD_V("SP", 1, "F1", "F2", "F3", "DR", "поселенцы")     

   Response.Write "<HR><p><font size='3' > "  

end if

if T="2" then
   if F3="" And F2="" then
      Response.Write "без № уд искать не буду !<BR>"
      Break
   end if
   Call POISK_ZAD("SP", 1, "5", "NLD", "AN", "7",  "Архив")
end if

db.Close
Set db = Nothing
'Response.Write "<HR><p><font size='1' > "  
Response.Write "время поиска "  
'response.write(DateDiff("h",time,tim_start) & " час ")
'response.write(DateDiff("n",time,tim_start) & " минут ")
response.write(DateDiff("s",time,tim_start) & " секунд <br />")
Response.Write " </font> "  


%>
<A href="sp--form.asp?T=1"><font color=Red>Выход</FONT></A><BR>
<!--<FONT size=1 Color=Tan>просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.06.2020 </FONT> -->
</HTML>
                                              