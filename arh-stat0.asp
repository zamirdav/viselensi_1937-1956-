<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>ARH-stat.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="ATOC_AI_1"   
pSlv="sp_nac"   '������� ������� �������� ������ ���������� 
BOK="r6"       '������� ���� �������� ������ ����������
SAP="r5"

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


dim MAC(999,999)

sub STAT_1
   ZAPROS = "Select * From"&Name_Zad&" order by "&rO&","&rG&";"
'      Response.Write ZAPROS & "<BR>"
   On Error Resume Next          ' �������� �������� ������
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "������ ����������� ���������� 0550102614<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      KOL=0
      Do while NOT Rs.EOF 
         KOL = KOL + 1 
         S=rs.Fields(BOK).value 
         K=rs.Fields(SAP).value 
'         if S<1 then S=999
'         if K<1 then K=999
'         if S>999 then S=999
'         if K>999 then K=999
         MAC(S,K)=MAC(S,K)+1
         Rs.MoveNext          
      Loop                    
end sub


sub PRINT(Z)
    if Z=0 then Z=""
    Response.Write "<TD>" & Z
end sub

'�������� MAC(S,K)
'sub MAC_CLEAR
   S=1
   Do while S<999
      K=1
      Do while K<999
         MAC(S,K)=0
         K=K+1
      Loop                    
      S=S+1
   Loop                    
'end sub



'  call MAC_CLEAR
   call STAT_1

   Response.Write "<TABLE border=1><TR><TD>-"
   K=1
   Do while K<122
      Response.Write "<TD>" & K 
      K=K+1
   Loop                    
   S=1
   Do while S<152
      DA=0                                    
      K=1
      Do while K<122
         if MAC(S,K)>0 then DA=1
         K=K+1
      Loop                    
      if DA>0 then
         Response.Write "<TR><TD>" & Slv_Get("N1_ORG",S)
         K=1
         Do while K<122
            call PRINT(MAC(S,K))
            K=K+1
         Loop                    
     end if
     S=S+1
   Loop                    
   Response.Write "</TABLE>"

%>
<BR>                  
<BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="�����.gif" border=0 alt="�����"></A><BR>
<FONT size=1 Color=Tan>  �������� ATOC.mdb ����� ASP &nbsp; &nbsp;  &nbsp; &nbsp; ������ 29.08.2020 </FONT><BR>
</BODY>
</HTML>

