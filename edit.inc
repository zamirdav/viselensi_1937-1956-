<%
' 2020.08.26 ��������� ������� Edit_RekS
'
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  

dim RS, Name_Zad, N, RS2, nRea


sub Edit_Rek(pA, pRek)                            ' �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   Response.Write chr(13) & "<TR><TD>" & pA & "<TD><input type='TEXT' name='" & pRek & "' value='" & znaRek & "'>"
end sub

sub Vid_Rek(pA, pRek)                     ' �������� ��������� �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
   end if
end sub                               

'sub Vid_RekS(pA, pRek, pSlv)              ' �������� ��������� �� ������� ����� + slv
'   znaRek=rs.Fields("r"&pRek).value
'   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
'      if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
'      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
'
'   end if
'end sub                               

sub Vid_RekC(pA, pRek)                     ' �������� ��������� �� ������� ����� � �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;<FONT color=RED><B>" & znaRek & "</B></FONT>"
   end if
end sub                               


      

sub Edid_RekD(pA, pRek)                     ' �������� ���� �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      znaRek=mid(znaRek,1,4) & "." & mid(znaRek,5,2) & "." & mid(znaRek,7,2) 
      Response.Write chr(13) & "<TR><TD>&nbsp;" & pA & "<TD> <input type='TEXT' name='" & pRek & "' value='" & znaRek & "'>" '&nbsp;" & znaRek 
   end if
end sub                               
                         
sub Edit_RekC(pA, pRek)                            ' �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   Response.Write chr(13) & "<TR><TD><FONT color=RED>" & pA & "<TD><input type='TEXT' name='" & pRek & "' value='" & znaRek & "'>"
end sub                               

sub Edit_Mem(pA, pRek)                            ' Memo-���� (������)
   znaRek=rs.Fields("r"&pRek).value
   Response.Write chr(13) & "<TR><TD>" & pA & "<TD><TextArea name='"&pRek&"' Cols=31 Rows=8>"&znaRek&"</TextArea>"
end sub                               


sub Edit_Mp(pA, pRek, pNomer)                       '  ����������
   Set rs2 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & N & ";")
   A=""
   Kol=0
   Do while NOT Rs2.EOF 
      Kol=Kol+1
      znaRek=rs2.Fields("r"&pRek).value
      A=A & "<input type='TEXT' name='" & pRek & "_" & Kol & "' value='" & znaRek & "'><BR>"
      Rs2.MoveNext  
   Loop 
   Kol=Kol+1                                       
   A=A & "<input type='TEXT' name='" & pRek & "_" & Kol & "' value=''>" 
   Response.Write chr(13) & "<TR><TD><FONT color=DeepSkyBlue>" & pA & "<TD>" & A
end sub                               


sub Edit_MPs(pA, pRek, pNomer, pSlv)                       '  ���������� + �������
   Set rs2 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & N & ";")
   A=""
   Kol=0
   Do while NOT Rs2.EOF 
      Kol=Kol+1
      znaRek=rs2.Fields("r"&pRek).value
      A=A & "<SELECT name='" & pRek & "_" & Kol & "'>" & Slv_Get(pSlv,znaRek) & "</SELECT><BR>"
      Rs2.MoveNext  
   Loop 
   Kol=Kol+1                                       
   A=A & "<SELECT name='" & pRek & "_" & Kol & "'>" & Slv_Get(pSlv,"") & "</SELECT>"
'  A=A & "<input type='TEXT' name='" & pRek & "_" & Kol & "' value=''>" 
   Response.Write chr(13) & "<TR><TD><FONT color=DeepSkyBlue>" & pA & "<TD>" & A
end sub                               


sub Edit_RekM(pA, pRek)                                '   ������������
   znaRek=rs2.Fields("r"&pRek).value
   Response.Write chr(13) & "<TR><TD>" & pA & "<TD><input type='TEXT' name='" & nRea & "_" & pRek & "' value='"&znaRek&"'>"     ' 2020.06.29 
end sub                               


sub Edit_RekMs(pA, pRek, pSlv)                                '   ������������ + �������
   znaRek=rs2.Fields("r"&pRek).value
'   A="<input type='TEXT' name='" & nRea & "_" & pRek & "' value='"&znaRek&"'>" 
   A="<SELECT name='" & nRea & "_" & pRek & "'>" & Slv_Get(pSlv,znaRek) & "</SELECT>"
   Response.Write chr(13) & "<TR><TD>" & pA & "<TD>" & A
end sub                               


sub Edit_MgMp(pA, pRek, pNomer)                        '  ���������� � ������������
   Set rs3 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & NN & ";")
   A=""
   Kol=0
   Do while NOT Rs3.EOF 
      Kol=Kol+1
      znaRek=rs3.Fields("r"&pRek).value
      A=A & "<input type='TEXT' name='" & nRea & "_" & pRek & "_" & Kol & "' value='" & znaRek & "'><BR>"     ' ------ 2020.06.29 
      Rs3.MoveNext          
   Loop 
   Kol=Kol+1                                       
   A=A & "<input type='TEXT' name='" & nRea & "_" & pRek & "_" & Kol & "' value=''>"                          ' ------ 2020.06.29 
   Response.Write chr(13) & "<TR><TD><FONT color=Orange>" & pA & "</FONT><TD>" & A
end sub                               



function Slv_Get(pSlv,znaRek)             ' ���������� �� ������� ��� � ����������� 
   On Error Resume Next
   Set RsSLV = db.Execute("Select * From slv_" & pSlv & ";")
   if db.Errors.Count > 0 then
      A="<option value='" & znaRek & "'> - �� �������� !"
   else
      A="<option value='0'>" 
      Do while NOT RsSLV.EOF 
         S=""
         KD=RsSLV.Fields("KD").value
         if KD&"~"=znaRek&"~" then S=" selected"
'                   Response.Write "-" & znaRek & "-" & KD & "-" & S & "<BR>"   
         A=A & "<option value='" & KD & "'" & S & ">" & KD & "-" & RsSLV.Fields("TX").value
         RsSLV.MoveNext          
      Loop 
   end if
   Set RsSLV = Nothing
   Slv_Get=A
end function




sub Edit_RekS(pA, pRek, pSlv)                            ' �� ������� ����� �� ��������
   znaRek=rs.Fields("r"&pRek).value
'   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
'   end if
   Response.Write chr(13) & "<TR><TD>" & pA & "<TD><SELECT name='" & pRek & "'>" & znaRek & "</SELECT>"
end sub                               

sub Edit_RekSC(pA, pRek, pSlv)                            ' �� ������� ����� �� ��������
   znaRek=rs.Fields("r"&pRek).value
'   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
'   end if
   Response.Write chr(13) & "<TR><TD><FONT color=Red>" & pA & "<TD><SELECT name='" & pRek & "'>" & znaRek & "</SELECT>"
end sub                               




sub OPEN_BAZA(pNam)
   Name_Zad=pNam                'Name_Zad=Session("Name_Zad") 
   Response.Write Name_Zad & " " & TipBazaT & " ������������� "
   
   N = Request.QueryString("N")     
   
   ZAPROS = " Select * From ATOC_"&Name_Zad&"_1 where AT_0=" & N & ";"
'   Response.Write " <FONT color=white>" & ZAPROS & "</FONT><BR>"
   Set rs = db.Execute(ZAPROS)
end sub
   




sub CLOSE_BAZA
   Set rs = Nothing
   db.Close
   Set db = Nothing
   Response.Write chr(13) & "<A href='../index.htm'> ����� </A><BR>"
end sub



sub Vid_Rek(pA, pRek)                     ' �������� ��������� �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
   end if
end sub                               



%>
