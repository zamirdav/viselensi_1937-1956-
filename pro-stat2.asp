<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>ARH-stat.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
Set db = Server.CreateObject("ADODB.Connection") ' коннектимся к базе
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


dim MAC(999)


sub MAC_CLEAR    ' обнуляет массив
   S=1
   Do while S<999
      MAC(S)=0
      S=S+1
   Loop                    

end sub



sub SPISOK_SHET                   ' подсчитывает коло-во в органах поиском из базы
   call MAC_CLEAR
   ZAPROS = "Select * From ATOC_o2_4 order by rO;"  ' where AT_K_D >='20210801' and AT_K_D <='20210831' order by rO;"
   On Error Resume Next          ' включает контроль ошибок
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка исполнения запроса<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      Do while NOT Rs.EOF 
         S=rs.Fields("rO").value   ' по органам
         if S<1 then S=100
'         if S>150 then S=151
         MAC(S)=MAC(S)+1        ' кол-во в  матрице
         Rs.MoveNext  
	 response.flush        
      Loop                    
end sub



sub SPISOK_PRINT   ' выводит на экран таблицу подсчета по органам----------------
   Response.Write "<TABLE><TR><TD>"
   Response.Write "<TABLE border=1>"
   S=1
   Do while S<999
      ZNA=MAC(S)
      if ZNA>0 then

        ZNA="<A href='pro-stat2.asp?O=" & S & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
         Response.Write "<TR><TD>" & Slv_Get("N1_ORG",S) & "<TD align=right>" & ZNA
      end if
      S=S+1
      if (S=40)or(S=78)or(S=123)or(S=164)or(S=313) then Response.Write "</TABLE><TD><TABLE border=1>" ' раскидываем по столбцам
   Loop                    
   Response.Write "</TABLE>"
   Response.Write "</TABLE>"
end sub




sub ORGAN_SHET(ORGAN)     ' второй поиск подсчитывает кол-во по годам----------------
   call MAC_CLEAR
   ZAPROS = "Select * From ATOC_o2_4 where rO='" & ORGAN & "' and AT_K_D >='20210801' and AT_K_D <='20210831' order by rG;"
   On Error Resume Next         
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка исполнения запроса<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      Do while NOT Rs.EOF 
         S=rs.Fields("rG").value     ' по органам
         if S<1 then S=100
'        if S>150 then S=151
         MAC(S)=MAC(S)+1             ' заполняет матрицу
         Rs.MoveNext          
	 response.flush
      Loop                    
end sub






sub ORGAN_PRINT(ORGAN)          ' выводит на экран таблицу подсчета по органам----------------
   Response.Write "<H2>" & Slv_Get("N1_ORG",ORGAN) & "</H2>"

   Response.Write "<TABLE border=1>"
   Response.Write "<TR><TD>годы" & "<TD> количество" 
   S=1
   Do while S<152
      ZNA=MAC(S)
      if ZNA>0 then
         ZNA="<A href='\ATOC_2020-09-02\ATOC\all-poisk.asp?T=2&P1=" & ORGAN & "&P2=" & S & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
'         ZNA="<A href='pro-stat3.asp?rG=" & S & " &rO=" & ORGAN & "'>&nbsp;<FONT size=+1> " & ZNA & "</FONT>&nbsp;</A>"
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

