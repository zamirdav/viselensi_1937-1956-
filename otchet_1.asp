<HTML>
<HEAD>
   <META http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>Natan отчет1</TITLE>
</HEAD>
<BODY BGCOLOR="lightblue">
<!--#INCLUDE FILE="init.inc"-->
<%
' отчет1
' Бобровский П.А.
' 2021.11.10               aquamarine


dim MAC_KOL, MAC_PODR(99), MAC_PODR_KOL(99), MAC_ODIR(99), MAC_HDD(99), MAC_INV(99)


function NOMER_PODR(ZNA)
   I=1
   do while I<=MAC_KOL
      if ZNA = MAC_PODR(I) then
         NOMER_PODR=I
         MAC_PODR_KOL(I)=MAC_PODR_KOL(I)+1
         exit function
      end if
      I=I+1
   Loop
   if MAC_KOL<99 then MAC_KOL=MAC_KOL+1
   MAC_PODR(MAC_KOL)=ZNA
   MAC_PODR_KOL(MAC_KOL)=1
   MAC_ODIR(MAC_KOL)=0
   MAC_HDD (MAC_KOL)=0
   MAC_INV (MAC_KOL)=0
   NOMER_PODR=MAC_KOL
end function



sub PUT_R(R)
    RR=R
    if R=0 then RR="&nbsp;"
    Response.Write "<TD align=right>" & RR
end sub


sub PUT_RESULT
   Response.Write "<TABLE align=center border=1 cellspacing=0 bordercolor=blue>"
   Response.Write "<TR><TD>N<TD>подразделение<TD>всего<TD>HDD<BR>ОДиР<TD>HDD<TD>инв.N"
   KOL=0
   KOL_ODIR=0
   KOL_HDD=0
   I=1
   do while I<=MAC_KOL
      Response.Write  "<TR><TD><FONT size=1>" & I & "<FONT>"
'     Response.Write  "<TD>" & MAC_PODR(I)
      Response.Write  "<TD><A href='vid_all.asp?PODR_PAR=" & MAC_PODR(I) & "'>" & GET_SLV("s_PODR",MAC_PODR(I)) & "</A>" 
      call PUT_R(MAC_PODR_KOL(I))
      call PUT_R(MAC_ODIR(I))
      call PUT_R(MAC_HDD(I))
      call PUT_R(MAC_INV(I))
      KOL=KOL+MAC_PODR_KOL(I)
      KOL_ODIR=KOL_ODIR+MAC_ODIR(I)
      KOL_HDD=KOL_HDD+MAC_HDD(I)
      KOL_INV=KOL_INV+MAC_INV(I)
      I=I+1
   Loop
   Response.Write  "<TR><TD>&nbsp;<TD><FONT size=+1 color=red>Итого</FONT><TD>" & KOL & "<TD>" & KOL_ODIR & "<TD>" & KOL_HDD & "<TD>" & KOL_INV
   Response.Write  "</TABLE>"
end sub




if OPEN_DSN=0 then
   ZAPR_R = "Select * From LICO;"   
   Response.Write    "Исполняем запрос " & ZAPR_R & "<BR>"
   Set RS = DB.Execute(ZAPR_R)
'   Response.Write "запрос выполнен<BR>"

   MAC_KOL=0
   Do while NOT RS.EOF 
'      Response.Write  RS.Fields("KUZDANIE").value & " " & RS.Fields("KABINET").value & " " & RS.Fields("PODR").value & " " & RS.Fields("PRIM").value & "<BR>"
      N=NOMER_PODR(RS.Fields("KU").value)
      if RS.Fields("HDDodir").value&"~" ="~" then MAC_ODIR(N)=MAC_ODIR(N)+1
      if mid(RS.Fields("INVENTARN").value,1,1)="-" then MAC_INV(N)=MAC_INV(N)+1
      RS.MoveNext          
'      call PUT_RESULT
   Loop                    
   Set RS = Nothing
   Response.Write  "Всего " & MAC_KOL & "<HR>"

   call PUT_RESULT
end if
db.Close
Set db = Nothing
%>
<BR>
<A href="index.htm"><IMG src="назад.gif" border=0 alt="Выход" width="64" height="32"></A>
<FONT size=1 face="Arial" color=lightblue>Бобровский П.А. 10.11.2021 </FONT>
<BR>
</BODY>
</HTML>
