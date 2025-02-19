<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>Lico</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="init.inc"-->
<%
'Session("Name_Zad")="ZZ"
'Response.Write Session("Name_Zad") & "<BR>"
'Response.Write "Lico<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.21<BR>"





sub Write_R(ID,nRek,znaRekPar)  
    Z_VVOD = "update LICO set " & nRek & "='" & znaRekPar & "' where L_Id=" & ID & ";"
    Response.Write " запрос " & Z_VVOD 
    DB.Execute(Z_VVOD) 
end sub                               



sub LICO_SPISOK_1
    On Error Resume Next
    ZAPROS = "select * from LICO;"
    Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
'   RS.MoveTop
    ID0=RS.Fields("L_ID").value 
    F10=RS.Fields("F1").value 
    RS.MoveNext          
    Do while NOT RS.EOF 
       ID=RS.Fields("L_ID").value 
       F1=RS.Fields("F1").value 
       call Write_R(ID,"F1",F10)  
       F10=F1
       RS.MoveNext          
    Loop                    
    call Write_R(ID0,"F1",F10)  
    Set RS = Nothing
end sub






if OPEN_DSN=0 then
   T = Request.QueryString("T")
   if T = "1" then call LICO_SPISOK_1
end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
