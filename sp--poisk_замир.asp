
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>����� ALL</title>
</HEAD>
<style>
  body { background: url(../���1.png);
        background-repeat: repeat;

       }
</style>

<!--#INCLUDE FILE="sp--poisk.inc"--> 
<%
' 2020.08.26 ���������� ��������� �������
'2021.08.12 Or F1="" �������� ��� ���������� ������ 40 ��� ����������� �������� 
'
'
Tim_start=Time
'Response.Write "���������� �.�.<BR>"
'Response.Write "2019.01.24<BR>"
'27.11.2020 ������� ����� �� ��� �3


Response.Write "<DIV align=right><A href='pass.asp'>&nbsp;" & KOT_FIO & "&nbsp;</A></DIV>"  

T = Request.QueryString("T") 

if T="1" then
   if F1="" then 
      Response.Write "��� ������� ������ �� ���� !<BR>"
      Break
   end if
   Call POISK_ZAD("SP", 1, "F1", "F2", "F3", "DR", "���������")     
   if KOL_ZAP=0 then
     Response.Write "<A href='add_IF.asp?F1=" & F1 & "&F2=" & F2 & "&F3=" & F3  & "&DR=" & DR& "'>����� ����� ������ ?</A>"
   end if

end if

if T="2" then
   if F3="" And F1="" then
      Response.Write "��� � �� ������ �� ���� !<BR>"
      Break
   end if
   Call POISK_ZAD("SP", 1, "5", "NLD", "AN", "7",  "�����")
end if

db.Close
Set db = Nothing
Response.Write "<HR>"
Response.Write "����� ������  " & mid(Time-Tim_start,1,4) & " ���.<BR>"


%>
<A href="sp--form.asp?T=1"><font color=Red><IMG src=�����.gif border=0 alt=�����>�����</FONT></A><BR>
<!--<FONT size=1 Color=Tan>�������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.06.2020 </FONT> -->
</HTML>
                                              