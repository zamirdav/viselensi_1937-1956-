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



sub LICO_SPISOK(ZAPR_DOP)
    On Error Resume Next
    ZAPROS = "select * from LICO " & ZAPR_DOP & " order by F1;"
    Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 bordercolor=blue>"
    N=0
    Do while NOT RS.EOF 
       N = N + 1     
       ID=RS.Fields("L_ID").value 
       Response.Write "<TR><TD>" & N
       Response.Write "<TD><A href='lico.asp?ID=" & ID & "'>" & RS.Fields("F1").value & "</A>"
       Response.Write "<TD>" & RS.Fields("F2").value 
       Response.Write "<TD>" & RS.Fields("F3").value 
       Response.Write "<TD>" & RS.Fields("DR").value 
       RS.MoveNext          
    Loop                    
    Set RS = Nothing
    Response.Write "</TABLE>"
    if ZAPR_DOP&"~" = "~" then Response.Write "<A href='lico.asp?T=ADD'>добавить лицо</A><BR>"
    Response.Write "</CENTER>"
end sub






sub FOTO_VID_ALL(ID)
    On Error Resume Next
    ZAPR_F = "select * from FOTO where LIC=" & ID & ";"
    Response.Write ZAPR_F & "<BR>"
    Set RS_F=DB.Execute(ZAPR_F)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPR_F & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2>"
    RS_F.MoveFirst
    do while NOT RS_F.EOF 
       F_ID=RS_F.Fields("F_ID").value
       Response.Write "<TR><TD>"   & RS_F.Fields("FIL_Name").value 
       Response.Write "<BR>size="  & RS_F.Fields("FIL_SIZE").value 
       Response.Write "<TR><TD><A href='lico.asp?ID="&F_ID&"&T=1'>" & "<IMG src='Download.asp?ID=" & F_ID & "' height=150></A>" 
       RS_F.MoveNext
    Loop
    Set RS_F = Nothing
    Set ZAPR_F = Nothing
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
end sub






sub FOTO_VID(ID)
    On Error Resume Next
    ZAPR_F = "select * from FOTO where LIC=" & ID & " and TIP=1;"
'   Response.Write ZAPR_F & "<BR>"
    Set RS_F=DB.Execute(ZAPR_F)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPR_F & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2><TR>"
    if NOT RS_F.EOF then
       F_ID=RS_F.Fields("F_ID").value
       Response.Write "<TD><A href='lico.asp?ID="&F_ID&"&T=1'>" & "<IMG src='Download.asp?ID=" & F_ID & "' height=150></A>" 
    end if
    RS_F.Close
    ZAPR_F = "select * from FOTO where LIC=" & ID & " and TIP=2;"
'   Response.Write ZAPR_F & "<BR>"
    Set RS_F=DB.Execute(ZAPR_F)
    if NOT RS_F.EOF then
       F_ID=RS_F.Fields("F_ID").value
       Response.Write "<TD><A href='lico.asp?ID="&F_ID&"&T=1'>" & "<IMG src='Download.asp?ID=" & F_ID & "' height=150></A>" 
    end if
    RS_F.Close
    Set RS_F = Nothing
    Set ZAPR_F = Nothing
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
    Response.Write "<A href='lico.asp?ID="&ID&"&T=ALL'>смотреть все</A><BR>"
end sub



sub LICO_VID(ID)
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=0 cellpadding=5><TR>"
    Response.Write "<TD valign=top>"
    LICO_VID_ODIN(ID)
    Response.Write "<TD valign=top>"
    call FOTO_VID(ID)
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
    Response.Write "<A href='lico_edit.asp?ID=" & ID & "'>edit222</A><BR>"
end sub





sub FOTO_VID_BIG(ID)
    On Error Resume Next
    ZAPROS = "select * from FOTO where F_ID=" & ID & ";"
'   Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
'   Response.Write "запрос выполнен<BR>"
    if NOT RS.EOF then
       F_ID=RS.Fields("F_ID").value
       Response.Write "Name "  & RS.Fields("FIL_Name").value & "<BR>"
       Response.Write "Size "  & RS.Fields("FIL_SIZE").value & "<BR>"
       Response.Write "<IMG src='Download.asp?ID=" & F_ID & "'>" 
    end if
    Set RS = Nothing
end sub







sub LICO_FORM_ZAPROS
    Response.Write "<BR>"
    Response.Write "<FORM action='lico.asp'>"
    Response.Write "<INPUT type='hidden' name='T' value='ISK'>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2>"
       Response.Write "<TR><TD>пол       <TD><SELECT name='PL' style='width: 60px'>"  & VID_SLV("s_PL",1) 
       Response.Write "<TR><TD>категория <TD><SELECT name='KU' style='width: 120px'>" & VID_SLV("s_KU",100) 
    Set RS = Nothing
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
    Response.Write "<input type='SUBMIT' value='Искать'>"
    Response.Write "</FORM>"
end sub







if OPEN_DSN=0 then
ID = Request.QueryString("ID")
T = Request.QueryString("T")
if ID&"~" = "~" then 
   select case T 
     case "ADD" :
           call LICO_ADD("Фамилия")
     case "ZAPROS" :
           call LICO_FORM_ZAPROS
     case "ISK" :
           PL = Request.QueryString("PL")
           KU = Request.QueryString("KU")
           A="where PL=" & PL & " and KU=" & KU
           call LICO_SPISOK(A)
     case else  
           call LICO_SPISOK("")
   end select
else
   if T&"~" = "~" then 
      call LICO_VID(ID)
   else
      if T = "ALL" then 
         call FOTO_VID_ALL(ID)
      else
         call FOTO_VID_BIG(ID)
      end if
   end if
end if
else

end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
