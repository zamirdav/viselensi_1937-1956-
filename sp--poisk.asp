
<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>����� ALL</title>
</HEAD>

<!--#INCLUDE FILE="sp--poisk.inc"--> 
<%
' 2022.08.26 ���������� ��������� �������
'2022.08.12 Or F1="" �������� ��� ���������� ������ 40 ��� ����������� �������� 
'
'
Tim_start=Time
'Response.Write "���������� �.�.<BR>"
'Response.Write "2022.01.24<BR>"
'27.11.2022 ������� ����� �� ��� �3



T = Request.QueryString("T") 

if T="1" then
   if F1="" then 
      Response.Write "��� ������� ������ �� ���� !<BR>"
      Break
   end if
   Call POISK_ZAD_V("SP", 1, "F1", "F2", "F3", "DR", "���������")     

   Response.Write "<HR><p><font size='3' > "  

end if

if T="2" then
   if F3="" And F2="" then
      Response.Write "��� � �� ������ �� ���� !<BR>"
      Break
   end if
   Call POISK_ZAD("SP", 1, "5", "NLD", "AN", "7",  "�����")
end if

db.Close
Set db = Nothing
'Response.Write "<HR><p><font size='1' > "  
Response.Write "����� ������ "  
'response.write(DateDiff("h",time,tim_start) & " ��� ")
'response.write(DateDiff("n",time,tim_start) & " ����� ")
response.write(DateDiff("s",time,tim_start) & " ������ <br />")
Response.Write " </font> "  


%>
<A href="sp--form.asp?T=1"><font color=Red>�����</FONT></A><BR>
<!--<FONT size=1 Color=Tan>�������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.06.2020 </FONT> -->
</HTML>
                                              