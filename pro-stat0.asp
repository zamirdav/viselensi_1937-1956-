<HTML>
<HEAD>
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
   <TITLE>pro-stat0.asp </TITLE>
</HEAD>
<BODY bgcolor=Tan>
<%
'Довлетбаев З.С
'07.07.2021 сделал статистику
'выводит на экран по строкам органы а по столбцам годы по выбранной таблице
'Response.Buffer = False 
'Response.Flush = false
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
Name_Zad="ATOC_o2_4"   


function Slv_Get(pSlv,znaRek) '---------- возвращает из словаря код и расшифровку 
   A=znaRek 
   On Error Resume Next
   Set rsSLV = db.Execute("Select * From slv_" & pSlv & " where KD=" & znaRek & " ;")
   if db.Errors.Count > 0 then
      A=A & " - не открылся"
   else
      if NOT rsSLV.EOF then A=A & " <FONT size=-1 color=DarkGreen>" & rsSLV.Fields("TX").value & "</FONT>"
      Set rsSLV = Nothing
   end if
   Slv_Get=A  ' возвращает из словаря код и расшифровку 
end function


dim MAC(999,999) ' обьявляем массив 

sub STAT_1    ' заполняет массив кол-ва органов по годам--------------------------
   ZAPROS = "Select * From "&Name_Zad&" order by rO,rG;" 'where AT_K_D >='2021.08.01' and AT_K_D <='2021.08.31' order by rO,rG;"
'      Response.Write ZAPROS & "<BR>"
   On Error Resume Next          ' включает контроль ошибок
   Set rs = db.Execute(ZAPROS)
   if db.Errors.Count > 0 then
      Response.Write "ошибка исполнения запроса<BR>"   
      Response.Write ZAPROS & "<BR>"
      exit sub
   end if
      KOL=0
      Do while NOT Rs.EOF 
         KOL = KOL + 1 
         S=rs.Fields("rO").value 
         K=rs.Fields("rG").value 
'         if S<1 then S=150    ' если 
         if K<1 then K=100
'         if S>150 then S=151
'         if K>150 then K=151
         MAC(S,K)=MAC(S,K)+1 ' заполняет массив кол-ва органов по годам
         Rs.MoveNext          
      Loop                    
end sub


sub PRINT(Z) ' выводит на экран таблицу----------------
    if Z=0 then Z=""
    Z="<A href='\ATOC_2020-09-02\ATOC\all-poisk.asp?T=2&P1=" & K & "&P2=" & S & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
    Response.Write "<TD>" & Z
end sub

   S=1              ' обнуляем массив---------------
   Do while S<999
      K=1
      Do while K<999
         MAC(S,K)=0
         K=K+1
      Loop                    
      S=S+1
   Loop                    



   call STAT_1
' обнуляем массив


' выводим таблицу органов по столбцам----------------
   Response.Write "<TABLE border=1><TR><TD>орган\годы"
   K=1
   Do while K<999                ' выводим номера колонок таблицы
      Response.Write "<TD>" & K 
      K=K+1
   Loop                    
   S=1
   Do while S<999                ' выводим строки таблицы
      DA=0                                    
      K=1
      Do while K<999
         if MAC(S,K)>0 then DA=1
         K=K+1
      Loop                    
      if DA>0 then
         Response.Write "<TR><TD>" & Slv_Get("N1_ORG",S) ' возвращает из словаря код и расшифровку
         K=1
         Do while K<999
           ZNA=MAC(S,K)
'            call PRINT(MAC(S,K))  ' выводим количество
           if ZNA=0 then ZNA=""
    	   ZNA="<A href='\ATOC_2020-09-02\ATOC\all-poisk.asp?T=2&P1=" & S & "&P2=" & K & "'>&nbsp;<FONT size=+1>" & ZNA & "</FONT>&nbsp;</A>"
'           Response.Write "<TR><TD>" & S & "<TD align=right>" & ZNA
	   Response.Write "<TD>" & ZNA
           K=K+1
        Loop                    
     end if
     S=S+1
   Loop                    
   Response.Write "</TABLE>"
' Response.Buffer = True 
%>
<BR>                  
<BR>
<BR>
<BR>
<A href="../index.htm"><IMG src="назад.gif" border=0 alt="Выход"></A><BR>
<FONT size=1 Color=Tan>  просмотр ATOC.mdb через ASP &nbsp; &nbsp;  &nbsp; &nbsp; версия 29.08.2020 </FONT><BR>
</BODY>
</HTML>

