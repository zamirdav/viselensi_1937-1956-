<%
' 2022.11.09
'
DIM T
sub FORMA_FIO
%>
<!--
<h6>��������� ������ <BR>
���������� ����� ������������<BR>
� 2022 ���� <BR>
-->
<CENTER>
<CENTER>
<table BORDER=5>
   <td WIDTH="270" height="180" BGCOLOR="lightblue" valign=top>
   <CENTER>
   <FONT size=5> ���� ����� ������ ? </FONT>
   <FORM action="sp--poisk.asp">        
      <input type="HIDDEN" name="T" value="1">
      <TABLE >
         <TR>
         <TD>
	 �������    <TD><input type="TEXT" name="P1" value="�">
         <TR><TD>���        <TD><input type="TEXT" name="P2" value="">
         <TR><TD>��������   <TD><input type="TEXT" name="P3" value="">
         <TR><TD>���� ����. <TD><input type="TEXT" name="P4" value="">
         <TR><TD> </TD><TD><BR>
         <input type="SUBMIT" value="������"></TD></TR>
         <TD>
         <A href="../index.htm" ><align=right>�����</A></TD></TR>
      </TABLE>
   </FORM>
   </CENTER>
</table>
</CENTER>
<BR>
<BR>
<%
end sub



sub FORMA_UD
%>
����� �� � �� ��� ����<BR>
<BR>
<BR>
<CENTER>
<CENTER>
<table BORDER=5>
   <tr>
   <td WIDTH="340" height="140" BGCOLOR="lightblue" valign=top>
   <CENTER><br>
   <FONT size=5> ����� �������� ����� ? </FONT><BR>
   <BR>
   <FORM action="sp--poisk.asp">
      <input type="HIDDEN" name="T" value="2">
      <TABLE>
	   <TR><TD>�������� �   <TD><input type="TEXT" name="P3" value="50">
	   <TR><TD>������ �     <TD><input type="TEXT" name="P2" value="">
	   <TR><TD> </TD><TD><BR>
           <input type="SUBMIT" value="������"></TD></TR>
           <TR><TD> </TD><TD><BR>
           <A href="../index.htm">�����</A></TD></TR>

      </TABLE>
   </FORM>
   </CENTER>
</table>
</CENTER>
<BR>
<BR>
<%
end sub

%>
