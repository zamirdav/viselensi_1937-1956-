
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  

Response.Write "<H2> делаем ATOC_AI_1 </H2>"

   ZAPR_DEL = "Drop Table ATOC_AI_1;"
   ZAPR_CRE = "Create Table ATOC_AI_1(AT_0 AutoIncrement CONSTRAINT AT_0 Primary Key,"
   ZAPR_CRE=ZAPR_CRE & " rO   VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rG   VarChar(15),"  
   ZAPR_CRE=ZAPR_CRE & " rN   VarChar(15),"  
   ZAPR_CRE=ZAPR_CRE & " r04  VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR7 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR2 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR3 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR4 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR5 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR6 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR8 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR9 VarChar(15),"
   ZAPR_CRE=ZAPR_CRE & " rAR0 VarChar(15));"
                       
Response.Write "команда=" & ZAPR_DEL & "<BR>"
On Error Resume Next
Set rs = db.Execute(ZAPR_DEL)  
if Err.Number<>0 then
   Response.Write "ошибка удаления ATOC_AI_1<BR>"
   Response.Write  Err.Number & ": " & Err.Description & "<HR>"
   Err.Clear
end if
Set rs = Nothing

Response.Write "команда=" & ZAPR_CRE & "<BR>"
Set rs = db.Execute(ZAPR_CRE)  
if Err.Number<>0 then
   Response.Write "ошибка создания ATOC_AI_1<BR>"
   Response.Write  Err.Number & ": " & Err.Description & "<HR>"
   Err.Clear
end if

Set rs = Nothing

Response.Write "<H2> делаем ATOC_AI_2 </H2>"


   ZAPR_DEL = "Drop Table ATOC_AI_2;"
   ZAPR_CRE = "Create Table ATOC_AI_2(AT_0 AutoIncrement CONSTRAINT AT_0 Primary Key, AT_1 integer, rAR1 VarChar(15));"
   ZAPR_CRE=ZAPR_CRE & ""



Response.Write "команда=" & ZAPR_DEL & "<BR>"
On Error Resume Next
Set rs = db.Execute(ZAPR_DEL)  
if Err.Number<>0 then
   Response.Write "ошибка удаления ATOC_AI_2<BR>"
   Response.Write  Err.Number & ": " & Err.Description & "<HR>"
   Err.Clear
end if
Set rs = Nothing

Response.Write "команда=" & ZAPR_CRE & "<BR>"
Set rs = db.Execute(ZAPR_CRE)  
if Err.Number<>0 then
   Response.Write "ошибка создания ATOC_AI_2<BR>"
   Response.Write  Err.Number & ": " & Err.Description & "<HR>"
   Err.Clear
end if

Set rs = Nothing




%>
