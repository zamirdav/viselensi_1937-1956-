
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  

dim NAM

sub FILE_CRE
   Response.Write "<H2> ������ SLV_"&NAM&"</H2>"             
   ZAPR = "Create Table SLV_" & NAM & "(ID AutoIncrement CONSTRAINT ID Primary Key,"
   ZAPR=ZAPR & " KD   Int,"
   ZAPR=ZAPR & " TX   VarChar(31));"  
   Response.Write "�������=" & ZAPR & "<BR>"
   On Error Resume Next
   Set RS = db.Execute(ZAPR)  
   if Err.Number<>0 then
      Response.Write "������ �������� SLV_"&NAM&"<BR>"
      Response.Write  Err.Number & ": " & Err.Description & "<HR>"
      Err.Clear
   end if
   Set RS = Nothing
end sub
                 

sub FILE_INS(KD,TX)
   ZAPR = "Insert into SLV_" & NAM & "(KD,TX) values("&KD&",'"&TX&"');"
   Response.Write "�������=" & ZAPR & "<BR>"
   On Error Resume Next
   Set RS = db.Execute(ZAPR)  
   if Err.Number<>0 then
      Response.Write "������ ���������� � SLV_"&NAM&"<BR>"
      Response.Write  Err.Number & ": " & Err.Description & "<HR>"
      Err.Clear
   end if
   Set RS = Nothing
end sub
                 


NAM="N1_ORG"

call FILE_CRE
call FILE_INS(1,  "���������")
call FILE_INS(2,  "��������")
call FILE_INS(3,  "������������")
call FILE_INS(5,  "����")
call FILE_INS(8,  "����������")
call FILE_INS(9,  "�����������")
call FILE_INS(11, "������")
call FILE_INS(12, "�������")
call FILE_INS(55, "�����������")
call FILE_INS(50, "������")
      


   Response.Write "<H2> ��������� </H2>"             


%>
