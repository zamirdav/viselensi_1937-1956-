
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>поиск ALL</title>
</HEAD>
<style>
  body { background: url(../ИМЯ1.png);
        background-repeat: repeat;

       }
</style>

<!--#INCLUDE FILE="sp--poisk.inc"--> 
<%
' 2020.08.26 переставил реквизиты местами
'2021.08.12 Or F1="" поставил для статистики строка 40 для внутреннего вложения 
'
'
Tim_start=Time
'Response.Write "Довлетбаев З.С.<BR>"
'Response.Write "2019.01.24<BR>"
'27.11.2020 добавил поиск по ОСК О3


Response.Write "<DIV align=right><A href='pass.asp'>&nbsp;" & KOT_FIO & "&nbsp;</A></DIV>"  

T = Request.QueryString("T") 

if T="1" then
   if F1="" then 
      Response.Write "без фамилии искать не буду !<BR>"
      Break
   end if
   Call POISK_ZAD("SP", 1, "F1", "F2", "F3", "DR", "поселенцы")     
   if KOL_ZAP=0 then
     Response.Write "<A href='add_IF.asp?F1=" & F1 & "&F2=" & F2 & "&F3=" & F3  & "&DR=" & DR& "'>введём новую запись ?</A>"
   end if

end if

if T="2" then
   if F3="" And F1="" then
      Response.Write "без № уд искать не буду !<BR>"
      Break
   end if
   Call POISK_ZAD("SP", 1, "5", "NLD", "AN", "7",  "Архив")
end if

db.Close
Set db = Nothing
Response.Write "<HR>"
Response.Write "время работы  " & mid(Time-Tim_start,1,4) & " сек.<BR>"


%>
<A href="sp--form.asp?T=1"><font color=Red><IMG src=назад.gif border=0 alt=Выход>Выход</FONT></A><BR>
<!--<FONT size=1 Color=Tan>просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.06.2020 </FONT> -->
</HTML>
                                              