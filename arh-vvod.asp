
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>vvod AI</title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->

<%
' � arh-poisk.inc ������� KOL_ZAP=0  ������ 84
O   = Request.QueryString("O") 
G   = Request.QueryString("G") 
N   = Request.QueryString("N") 
ADRES = Request.QueryString("ADRES") 

Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  


sub RECORD_ADD  ' ���������� ������ � ���� ATOC_AI_1
   ZAPROS="select * from ATOC_AI_1 where rO='" & O & "' and rG='" & G & "' and rN='" & N & "';"

   On Error Resume Next  
   Set RS = DB.Execute(ZAPROS) 
   if db.Errors.Count > 0 then
      Response.Write "������ ����������� ���������� 0550102614 " & ZAPROS & "<BR>" 
      Set RS = Nothing
   else
      if not RS.eof then
         ADR=RS.Fields("AT_0").value
         Response.Write "����� ������ ��� ���� <BR>"  
         Response.Write "<A href='edit-ai.asp?N=" & ADR & "'>�������</A><BR>"   
         Set RS = Nothing
      else
         Z_VVOD1 = "insert into ATOC_AI_1(rO,rG,rN)"  
         Z_VVOD2 = " values('" & O & "','" & G & "','" & N & "');" 
         Response.Write Z_VVOD1&Z_VVOD2 & "<BR>"
         Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
         if db.Errors.Count > 0 then
            Response.Write "������ ����������<BR>"   
            Set RZ_Z = Nothing
         else
            Set RZ_Z = Nothing
            Set RS = DB.Execute(ZAPROS) 
            if db.Errors.Count > 0 then
               Response.Write "������ ����������� ���������� 0550102614 " & ZAPROS & "<BR>" 
               Set RS = Nothing
            else
               if RS.eof then
                  Response.Write "������-�� ����� !<BR>"   
               else
                  ADR=RS.Fields("AT_0").value
                  Set RS = Nothing
                  Response.Redirect "edit-ai.asp?N=" & ADR 
'                 Response.Write "���������<BR>"   
               end if
            end if
         end if
      end if
   end if
end sub



if O&"~" = "~" then
   Response.Write "�� �������� ����� �����<BR>"   
else
   if G&"~" = "~" then
      Response.Write "�� �������� ��� ��<BR>"   
   else
      if N&"~" = "~" then
         Response.Write "�� �������� N ��<BR>"   
      else
         call RECORD_ADD
      end if
   end if
end if


%>
<!--
<CENTER>
<table BORDER=40>
   <tr>
   <td WIDTH="340" BGCOLOR="lightblue" valign=top>
   <CENTER><br>
   <FONT size=5> ������� �� </FONT><BR>
   <BR>
   <FORM action="add_ia.asp">
      <TABLE>
<%
   znaRek=Slv_Get("N1_ORG",O)
   Response.Write "<TR><TD>�����  <TD><SELECT name='O'>" & znaRek & "</SELECT>"
   Response.Write "<TR><TD>���    <TD><input type='TEXT' name='G' value='" & G & "'>"
   Response.Write "<TR><TD>� ��   <TD><input type='TEXT' name='N' value='" & N & "'>"
%>
	   <TR><TD> </TD><TD><BR>
           <input type="SUBMIT" value="���������"></TD></TR>
      </TABLE>
   </FORM>
   </CENTER>
</table>
</CENTER>
<BR>
<%




%>
-->