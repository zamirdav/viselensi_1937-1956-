
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>del multi</title>
</HEAD>
<BODY>
<%
Name_Zad="ZZ"
ADRES = Request.QueryString("N")
NN    = Request.QueryString("NN")
nFil  = Request.QueryString("F")

ZAPROS = "Delete from ATOC_"&Name_Zad&"_"&nFil&" where AT_0=" & NN & ";"
'Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC"                                  
Set RZ_Z = DB.Execute(ZAPROS) 
Set RZ_Z = Nothing

db.Close
Set db = Nothing

Response.Redirect "edit.asp?N=" & ADRES
%>
