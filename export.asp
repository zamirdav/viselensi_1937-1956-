<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>Lico Export</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="init.inc"-->
<%
'Response.Write "Lico Export<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.26<BR>"





sub FOTO_EXPORT
    On Error Resume Next
    ZAPROS = "select * from FOTO;"
    Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Set objStream = Server.CreateObject("ADODB.Stream")

    N=0
    Do while NOT RS.EOF 
       N = N + 1     
       objStream.Type = 1
       objStream.Open
       objStream.Write RS.Fields("FOT").Value
       objStream.SaveToFile "C:/TEMP/x-" & N & ".jpg", 2
       objStream.Close
       RS.MoveNext          
    Loop                    
    Set objStream = Nothing
    Set RS = Nothing
end sub


if OPEN_DSN=0 then
'   call LICO_EXPORT
   call FOTO_EXPORT
else

end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
