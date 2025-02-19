<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>ARH-stat.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="?"   


function Slv_Get(pSlv,znaRek)             ' возвращает из словаря код и расшифровку 
   A=znaRek 
   On Error Resume Next
   Set rsSLV = db.Execute("Select * From slv_" & pSlv & " where KD=" & znaRek & ";")
   if db.Errors.Count > 0 then
      A=A & " - не открылся"
   else
      if NOT rsSLV.EOF then A=A & " <FONT size=-1 color=DarkGreen>" & rsSLV.Fields("TX").value & "</FONT>"
      Set rsSLV = Nothing
   end if
   Slv_Get=A
end function


dim MAC(151)


sub MAC_CLEAR
   S=1
   Do while S<151
      MAC(S)=0
      S=S+1
   Loop                    

end sub



sub SPISOK_SHET
   call MAC_CLEAR
   ZAPROS = "Select * From ATOC_AI_1 order by rO;"
   On Error Resume Next          ' включает контроль ошибок
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка обращайтесь Довлетбаев 0550102614<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      Do while NOT Rs.EOF 
         S=rs.Fields("rO").value 
         if S<1 then S=150
         if S>150 then S=151
         MAC(S)=MAC(S)+1
         Rs.MoveNext          
      Loop                    
end sub



sub SPISOK_PRINT
   Response.Write "<TABLE><TR><TD>"
   Response.Write "<TABLE border=1>"
   S=1
   Do while S<152
      ZNA=MAC(S)
      if ZNA>0 then
         ZNA="<A href='arh-stat2.asp?O=" & S & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
         Response.Write "<TR><TD>" & Slv_Get("N1_ORG",S) & "<TD align=right>" & ZNA
      end if
      S=S+1
      if (S=14)or(S=51)or(S=81) then Response.Write "</TABLE><TD><TABLE border=1>"
   Loop                    
   Response.Write "</TABLE>"
   Response.Write "</TABLE>"
end sub




sub ORGAN_SHET(ORGAN)
   call MAC_CLEAR
   ZAPROS = "Select * From ATOC_AI_1 where rO='" & ORGAN & "' order by rG;"
   On Error Resume Next         
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка обращайтесь Довлетбаев 0550102614<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      Do while NOT Rs.EOF 
         S=rs.Fields("rG").value 
         if S<1 then S=150
         if S>150 then S=151
         MAC(S)=MAC(S)+1
         Rs.MoveNext          
      Loop                    
end sub






sub ORGAN_PRINT(ORGAN)
   Response.Write "<H2>" & Slv_Get("N1_ORG",ORGAN) & "</H2>"

   Response.Write "<TABLE border=1>"
   S=1
   Do while S<152
      ZNA=MAC(S)
      if ZNA>0 then
         ZNA="<A href='arh-poisk.asp?T=2&P1=" & ORGAN & "&P2=" & S & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
         Response.Write "<TR><TD>" & S & "<TD align=right>" & ZNA
      end if
      S=S+1
   Loop                    
   Response.Write "</TABLE>"
end sub




ORGAN = Request.QueryString("O") 
if ORGAN&"~" = "~" then 
   call SPISOK_SHET
   call SPISOK_PRINT
else 
   call ORGAN_SHET(ORGAN)
   call ORGAN_PRINT(ORGAN)
end if




%>
<BR>                  
<BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="назад.gif" border=0 alt="Выход"></A><BR>
<FONT size=1 Color=Tan>  просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.08.2020 </FONT><BR>
</BODY>
</HTML>

