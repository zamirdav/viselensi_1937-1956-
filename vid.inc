<HTML>
<HEAD>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"    
' 01012020 �������� ������������� ������ 273
'
'<style>
'.modalbackground{
'    background: rgba(160, 143, 143, 0.8);
'    padding: 15px;
'    margin: 2px 50%;
'    position: absolute;
'    width: 50%;
'    height: 50%;
'    color: #000;
'}
'
'.modalbackground a {
' color: #fff;
'}
'</style>
'<div id="mokno" class="modalbackground">
'    <div class="modalwindow">
'        <h3>������!</h3>
'        <p>
'        </p>
'        <a href="http://127.0.0.1/ATOC_2020-09-02/ATOC/arh-poisk.asp?T=1&P1=11&P2=14&P3=50&P4=">�������</a>
'    </div>
'</div>
'<%
dim RS, RS_F, Name_Zad, N, RS2, nRea

   KOT_VES = Session("VES")                          
'  Response.Write "KOT_VES=" & KOT_VES & "<BR>"
   if KOT_VES&"~" = "~" then
      IP_ADR= Request.ServerVariables("REMOTE_ADDR")     
'     Response.Write "ADDR=" & IP_ADR & "<BR>"
      ZAPR_P="select * from IP_SPISOK where IP='" & IP_ADR & "';"
      Set RSP = DB.Execute(ZAPR_P)  
      if Err.Number<>0 then
         Response.Write "������ " & ZAPR_P & " �� ����������.<BR>"
         Response.Write "DSN=" & DSN & "<A href='PAROL/'>.</A><BR>"
         Err.Clear
      else
         KOT_FIO="����"
         KOT_VES=0
         Do while NOT RSP.EOF 
            Session("VES")=rsP.Fields("KOT").value
            Session("NAM")=rsP.Fields("FIO").value
            KOT_VES = Session("VES")
'           Response.Write "KOT_VES=" & KOT_VES & "<BR>"
            RSP.MoveNext          
         Loop 
         Set RSP = Nothing
      end if                                              
   end if
   KOT_FIO = Session("NAM")                          



sub Vid_Rek(pA, pRek)                     ' �������� ��������� �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
   end if
end sub                               

sub Vid_RekB(pA, pRek)                     ' �������� ��������� �� ������� ����� ��� r ����� �����
   znaRek=rs.Fields(pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
   end if
end sub                               



sub Vid_RekD(pA, pRek)                     ' �������� ���� �� ������� �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      znaRek=mid(znaRek,1,4) & "." & mid(znaRek,5,2) & "." & mid(znaRek,7,2) 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
   end if
end sub                               


sub Vid_RekC(pA, pRek)                     ' �������� ��������� �� ������� ����� � �����
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;<FONT color=RED><B>" & znaRek & "</B></FONT>"
   end if
end sub                               


function Slv_Get(pSlv,znaRek)             ' ���������� �� ������� ��� � ����������� 
   A=znaRek 
   On Error Resume Next
   Set rsSLV = db.Execute("Select * From slv_" & pSlv & " where KD=" & znaRek & ";")
   if db.Errors.Count > 0 then
      A=A & " - �� ��������"
   else
      if NOT rsSLV.EOF then A=A & " <FONT size=-1 color=DarkGreen>" & rsSLV.Fields("TX").value & "</FONT>"
      Set rsSLV = Nothing
   end if
   Slv_Get=A
end function


sub Vid_RekS(pA, pRek, pSlv)              ' �������� ��������� �� ������� ����� + slv
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 

   end if
end sub                               



sub Vid_RekSC(pA, pRek, pSlv)              ' �������� ��������� �� ������� ����� + slv
   znaRek=rs.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
'      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;" & znaRek 
      Response.Write "<TR><TD>&nbsp;" & pA & "<TD>&nbsp;<FONT color=Red><B>" & znaRek & "</B></FONT>"
   end if
end sub                               

sub Vid_MpS(pA, pRek, pNomer, pSlv)       '  ���������� ���������
   Set rs2 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & N & ";")
   A=""
   Do while NOT Rs2.EOF 
      znaRek=rs2.Fields("r"&pRek).value
      if (znaRek<>"0")and(znaRek&"~"<>"~") then 
         if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
         A=A & "&nbsp;" & znaRek & "<BR>"
      end if
      Rs2.MoveNext          
   Loop 
   Set rs2 = Nothing
   if A&"~"<>"~" then Response.Write "<TR><TD>&nbsp;<FONT color=DeepSkyBlue>" & pA & "<TD>" & A
end sub                               


sub Vid_Mp(pA, pRek, pNomer)                       '  ����������
    call Vid_MpS(pA, pRek, pNomer, "")
end sub                               



sub Vid_RekMD(pA, pRek)                                 '   ������������ + ����
    znaRek=rs2.Fields("r"&pRek).value
    if (znaRek<>"0")and(znaRek&"~"<>"~") then 
       znaRek=mid(znaRek,1,4) & "." & mid(znaRek,5,2) & "." & mid(znaRek,7,2) 
       Response.Write "<TR><TD>&nbsp;" & pA & "<TD>" & znaRek 
    end if
end sub                               


sub Vid_RekMs(pA, pRek, pSlv)                                '   ������������ + SLV
    znaRek=rs2.Fields("r"&pRek).value
    if (znaRek<>"0")and(znaRek&"~"<>"~") then 
       if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
       Response.Write "<TR><TD>&nbsp;" & pA & "<TD>" & znaRek 
    end if
end sub                               

sub Vid_RekMsC(pA, pRek, pSlv)                                '   ������������ + SLV + ����
    znaRek=rs2.Fields("r"&pRek).value
    if (znaRek<>"0")and(znaRek&"~"<>"~") then 
       if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
       Response.Write "<TR><TD>&nbsp;<FONT color=RED>" & pA & "<TD>" & znaRek
    end if
end sub                               


sub Vid_RekM(pA, pRek)                                '   ������������
    call Vid_RekMs(pA, pRek, "")
end sub                               

sub Vid_RekMC(pA, pRek)                                '   ������������
    call Vid_RekMsC(pA, pRek, "")
end sub                               


sub Vid_RekMexe(pA, pRek, pProg)                       '   ������������ + ����� ��������� 
   znaRek=rs2.Fields("r"&pRek).value
   if (znaRek<>"0")and(znaRek&"~"<>"~") then 
      Response.Write chr(10) & "<TR><TD>&nbsp;<FONT color=RED>" & pA & "<TD>" & pProg &   znaRek & "</A>"
   end if
end sub                               


sub Vid_MgMpS(pA, pRek, pNomer, pSlv)                        '  ���������� � ������������ + SLV
   Set rs3 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & NN & ";")
   A=""
   Do while NOT Rs3.EOF 
      znaRek=rs3.Fields("r"&pRek).value
      if (znaRek<>"0")and(znaRek&"~"<>"~") then 
         if pSlv<>"" then znaRek=Slv_Get(pSlv,znaRek)
         A=A & znaRek & "<BR>"
      end if
      Rs3.MoveNext          
   Loop 
   Set rs3 = Nothing
   if A&"~"<>"~" then Response.Write "<TR><TD>&nbsp;<FONT color=Orange>" & pA & "</FONT><TD>" & A
end sub                               


sub Vid_MgMp(pA, pRek, pNomer)                        '  ���������� � ������������ 
   Set rs3 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & NN & ";")
   A=""
   Do while NOT Rs3.EOF 
      znaRek=rs3.Fields("r"&pRek).value
      if (znaRek<>"0")and(znaRek&"~"<>"~") then A=A & znaRek & "<BR>"
      Rs3.MoveNext          
   Loop 
   Set rs3 = Nothing                                    
   if A&"~"<>"~" then Response.Write "<TR><TD>&nbsp;<FONT color=Orange>" & pA & "</FONT><TD>" & A   ' --- 2020.06.29
end sub                               



sub Stat_IP_Add(Zad,Vid,Str)                            
   IP  = Request("REMOTE_ADDR")
   Dat = Date()
   Dat = mid(Dat,7,4) & "." & mid(Dat,4,3) & mid(Dat,1,2) 
   Tim = Time()
   if Len(Tim) < 8 then Tim="0" & Tim
   Par = Request.QueryString
   Str1= Replace(Str,"'","''")
   ZAPR="insert into Stat_IP (Dat,Tim,Zad,Vid,IP,Par,Str) values ('" &Dat&"','"&Tim&"','"&Zad&"','"&Vid&"','"&IP&"','"&Par&"','"&Str1&"');"
   On Error Resume Next
   Set rStat = db.Execute(ZAPR)  
   if db.Errors.Count > 0 then
      Response.Write "ZAPR=" & ZAPR & "<BR>"
      Response.Write "������ ������� ������ � ������� <BR>"   
      exit sub
   end if
   Set rStat = Nothing
end sub



sub OPEN_BAZA(pNam,pCL)
   Name_Zad=pNam                '   Name_Zad=Session("Name_Zad") 
   Response.Write "<TITLE>��� " & Name_Zad & "</TITLE>"
   Response.Write "</HEAD>"
   Response.Write "<BODY bgcolor=" & pCL & ">"
'   Response.Write "�������� " & Name_Zad
   N = Request.QueryString("N")    ' ����� ������
   F = Request.QueryString("F")    ' ����� �����

'Response.Write "    N=" & N & "<BR>"
'Response.Write "    F=" & F & "<BR>"
   if F&"~" = "~" then F="1"

   if F <> "1" then
      ZAPROS = "Select * From ATOC_"&Name_Zad&"_" & F & " where AT_0=" & N & ";"  
      Set RS = db.Execute(ZAPROS)                                
      if RS.EOF then
         Response.Write "<HR> ����� !<BR>"
      else
         N=RS.Fields("AT_1").value  
      end if
      Set RS = Nothing
   end if

   ZAPROS = "Select * From ATOC_"&Name_Zad&"_1 where AT_0=" & N & ";"
   call Stat_IP_Add(Name_Zad, "Vid-"&F, ZAPROS)      
'   Response.Write "<FONT color=white>" & ZAPROS & "</FONT><BR>"
'  Response.Write    "��������� ������<BR>"
   Set RS = db.Execute(ZAPROS)
'  Response.Write "������ ��������<BR>"
   Response.Write "<CENTER>"
end sub


sub CLOSE_BAZA
   Response.Write "</CENTER>"
   db.Close
   Set db = Nothing
'   Response.Write chr(13) & "<A href='edit-" & Name_Zad & ".asp?N=" & N & "'><DIV align=left> �������������</DIV></A><BR>" 
   Response.Write chr(13) & "<A href='sp--form.asp?T=1'><DIV align=left>����� � ���� ������</DIV></A><BR>"
   Response.Write chr(13) &"<BR>"
   Response.Write chr(13) &"<BR>"
end sub

sub CLOSE_BAZA_2 '����� ������������� Response.Write chr(13) & "<A href='edit-" & Name_Zad & ".asp?N=" & N & "'><DIV align=left> �������������</DIV></A><BR>"  
   Response.Write "</CENTER>"
   db.Close
   Set db = Nothing
   Response.Write chr(13) & "<A href='sp--form.asp?T='&T><DIV align=left>����� � ���� ������</DIV></A><BR>"

end sub


%>                                        