<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>smotr_f</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="initf.inc"-->
<%
'Response.Write "smotr_f<BR>"
'Response.Write "zamir<BR>"
'Response.Write "2022.10.17<BR>"

ID = Request.QueryString("N")
'Response.Write "ID="&ID&"<BR>"
ZAPROS="select * from ATOC_SP_F where AT_0='" & ID & "';"

Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  

   On Error Resume Next  
   Set RS = DB.Execute(ZAPROS) 
   if db.Errors.Count > 0 then
      Response.Write "ошибка исполнения запроса " & ZAPROS & db.Errors.Count&"<BR>" 
      Set RS = Nothing
   else
      if not RS.eof then
         ADR=RS.Fields("AT_0").value
         F_NAME=RS.Fields("F_NAME").value
         PATH ="/foto_gala/"
         FIL_NAME =PATH & F_NAME

'        Response.Write "такая запись есть <BR>"
   Response.Write "<EMBED src='" & FIL_NAME&"'  height=540 width =1400></EMBED>" 
     end if
   end if
                  


'end if


%>
<A href="all-form.asp?T=1"> Выход </A><BR>
