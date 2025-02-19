
<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  

dim NAM

sub FILE_CRE
   Response.Write "<H2> делаем SLV_"&NAM&"</H2>"             
   ZAPR = "Create Table SLV_" & NAM & "(ID AutoIncrement CONSTRAINT ID Primary Key,"
   ZAPR=ZAPR & " KD   Int,"
   ZAPR=ZAPR & " TX   VarChar(31));"  
   Response.Write "команда=" & ZAPR & "<BR>"
   On Error Resume Next
   Set RS = db.Execute(ZAPR)  
   if Err.Number<>0 then
      Response.Write "ошибка создания SLV_"&NAM&"<BR>"
      Response.Write  Err.Number & ": " & Err.Description & "<HR>"
      Err.Clear
   end if
   Set RS = Nothing
end sub
                 

sub FILE_INS(KD,TX)
   ZAPR = "Insert into SLV_" & NAM & "(KD,TX) values("&KD&",'"&TX&"');"
   Response.Write "команда=" & ZAPR & "<BR>"
   On Error Resume Next
   Set RS = db.Execute(ZAPR)  
   if Err.Number<>0 then
      Response.Write "ошибка добавления в SLV_"&NAM&"<BR>"
      Response.Write  Err.Number & ": " & Err.Description & "<HR>"
      Err.Clear
   end if
   Set RS = Nothing
end sub
                 


NAM="N1_ORG"

call FILE_CRE
call FILE_INS(1,  "Ленинский")
call FILE_INS(2,  "Первомай")
call FILE_INS(3,  "Свердловский")
call FILE_INS(5,  "Кант")
call FILE_INS(8,  "Московский")
call FILE_INS(9,  "Сокулукский")
call FILE_INS(11, "Токмак")
call FILE_INS(12, "Чуйский")
call FILE_INS(55, "Октябрьский")
call FILE_INS(50, "Бишкек")
      


   Response.Write "<H2> закончили </H2>"             


%>
