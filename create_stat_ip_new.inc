<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
   Response.Write "<H2> �������� ��������� ������� </H2>"  

sub Create_ZAD(Tab,Ima,Ika)
    Response.Write "Ima-�������� �������     <BR>"   
    Response.Write "Tab-��� ��������� �������<BR>"   
    Response.Write "Ika-������� �������      <BR>"   
    ZAPR_IND = "Create Table "&Tab & " Select * from "&Ima &" where "&Ika&";"   ' 
'    Response.Write"�������=" & ZAPR_Ind & "<BR>"
    Set rs = db.Execute(ZAPR_Ind)                        
    if db.Errors.Count > 0 then
      Response.Write "������ �������� ������� <BR>"   
      Response.Write ZAPR_IND & "<BR>"
      Response.Write "</CENTER>"
      exit sub
    end if
    Response.Write "�������=" & ZAPR_IND & "<BR>"
Set rs = Nothing

end sub

sub Delete_ZAD(Ima,Ika)
    Response.Write "Ima-�������� �������     <BR>"   
    Response.Write "Tab-��� ��������� �������<BR>"   
    Response.Write "Ika-������� �������      <BR>"   
    ZAPR_DEL = "Delete from "&Ima &" where "&Ika&";"   ' 
'    Response.Write"�������=" & ZAPR_Ind & "<BR>"
    Set rs = db.Execute(ZAPR_DEL)                        
    if db.Errors.Count > 0 then
      Response.Write "������ �������� ������� <BR>"   
      Response.Write ZAPR_DEL & "<BR>"
      Response.Write "</CENTER>"
      exit sub
    end if
    Response.Write "�������=" & ZAPR_DEL & "<BR>"
Set rs = Nothing

end sub

%>
