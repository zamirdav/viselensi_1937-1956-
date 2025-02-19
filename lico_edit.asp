<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>Lico</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="initF.inc"-->
<%
'Session("Name_Zad")="ZZ"
'Response.Write Session("Name_Zad") & "<BR>"
'Response.Write "Lico<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.21<BR>"



sub FOTO_VID(ID)
    On Error Resume Next
    ZAPR_F = "select * from ATOC_sp_F where AT_1=" & ID & ";"
    Response.Write ZAPR_F & "<BR>"
    Set RS_F=DB.Execute(ZAPR_F)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2>"
    do while NOT RS_F.EOF 
       F_ID=RS_F.Fields("F_ID").value
       Response.Write "<TR><TD>Tip      <TD>"  & RS_F.Fields("TIP").value 
       Response.Write "<TR><TD>Prim     <TD>"  & RS_F.Fields("FIL_NAME").value 
       Response.Write "<TR><TD>Fil_Size <TD>"  & RS_F.Fields("Fil_Size").value 
       Response.Write "<TR><TD>F_Prim   <TD>"  & RS_F.Fields("F_Prim").value 
       Response.Write "<TR><TD>TIP_CONT <TD>"  & RS_F.Fields("TIP_COBT").value 
       Response.Write "<TR><TD>&nbsp    <TD><TABLE width=100% bgcolor=aquamarine><TR>"
       Response.Write "<TD align=left><A href='foto_edit.asp?ID=" & F_ID & "&T=FOTO_EDIT'>исправить</A>" 
       Response.Write "<TD align=right><A href='foto_edit.asp?ID=" & F_ID & "&T=FOTO_DEL'>удалить</A>" 
       Response.Write "</TABLE>" 
       RS_F.MoveNext
    Loop
    Set RS_F = Nothing
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
    Response.Write "<HR><A href='foto_add.asp?ID=" & ID & "'>добавить</A><BR>"
end sub



sub LICO_EDIT(ID)
    On Error Resume Next
    ZAPROS = "select * from atoc_sp_1 where AT_0=" & ID & ";"
    Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Response.Write "<CENTER>"
    Response.Write "<TABLE border=0 cellpadding=5><TR><TD>"            ' 
    Response.Write "<FORM action='lico_save.asp'>"
    Response.Write "<INPUT type='hidden' name='ID' value='" & ID & "'>"
    Response.Write "<TABLE border=1 cellspacing=0 cellpadding=2>"
    if NOT RS.EOF then
       Response.Write "<TR><TD>Фамилия <TD><INPUT type='TEXT' name='F1' size=21 value='" & RS.Fields("F1").value & "'>"
       Response.Write "<TR><TD>Имя     <TD><INPUT type='TEXT' name='F2' size=21 value='" & RS.Fields("F2").value & "'>"
       Response.Write "<TR><TD>Отч     <TD><INPUT type='TEXT' name='F3' size=21 value='" & RS.Fields("F3").value & "'>"
       Response.Write "<TR><TD>DR      <TD>"  & RS.Fields("DR").value 
       Response.Write "<TR><TD>PL      <TD><SELECT name='PL' style='width: 60px'>" & VID_SLV("s_PL",RS.Fields("PL").value) 
       Response.Write "<TR><TD>м/р     <TD><INPUT type='TEXT' name='MR' size=21 value='" & RS.Fields("MR").value & "'>"
       Response.Write "<TR><TD>KU      <TD><SELECT name='KU' style='width: 150px'>" & VID_SLV("s_KU",RS.Fields("KU").value) 
       Response.Write "<TR><TD>Adr     <TD>"  & RS.Fields("Adr").value 
    end if                    
    Set RS = Nothing
    Response.Write "</TABLE>"
    Response.Write "<input type='SUBMIT' value='Сохранить'>"
    Response.Write "</FORM>"
    Response.Write "<TD>"
    call FOTO_VID(ID)
    Response.Write "</TABLE>"
    Response.Write "</CENTER>"
'   Response.Write "<A href='lico_edit.asp?ID=" & ID & "'>edit</A><BR>"
end sub





if OPEN_DSN=0 then
ID = Request.QueryString("ID")
if ID&"~" = "~" then 
   T = Request.QueryString("T")
   if T&"~" = "~" then 
      Response.Write "нету параметра.<BR>"
   else
'     call LICO_SAVE
   end if
else
   call LICO_EDIT(ID)
end if
else

end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
