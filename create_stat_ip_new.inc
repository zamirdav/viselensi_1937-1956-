<%
Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"   
   Response.Write "<H2> создание фиктивной таблицы </H2>"  

sub Create_ZAD(Tab,Ima,Ika)
    Response.Write "Ima-оригинал таблицы     <BR>"   
    Response.Write "Tab-имя фиктивной таблицы<BR>"   
    Response.Write "Ika-условие выборки      <BR>"   
    ZAPR_IND = "Create Table "&Tab & " Select * from "&Ima &" where "&Ika&";"   ' 
'    Response.Write"команда=" & ZAPR_Ind & "<BR>"
    Set rs = db.Execute(ZAPR_Ind)                        
    if db.Errors.Count > 0 then
      Response.Write "ошибка создания таблицы <BR>"   
      Response.Write ZAPR_IND & "<BR>"
      Response.Write "</CENTER>"
      exit sub
    end if
    Response.Write "команда=" & ZAPR_IND & "<BR>"
Set rs = Nothing

end sub

sub Delete_ZAD(Ima,Ika)
    Response.Write "Ima-оригинал таблицы     <BR>"   
    Response.Write "Tab-имя фиктивной таблицы<BR>"   
    Response.Write "Ika-условие выборки      <BR>"   
    ZAPR_DEL = "Delete from "&Ima &" where "&Ika&";"   ' 
'    Response.Write"команда=" & ZAPR_Ind & "<BR>"
    Set rs = db.Execute(ZAPR_DEL)                        
    if db.Errors.Count > 0 then
      Response.Write "ошибка создания таблицы <BR>"   
      Response.Write ZAPR_DEL & "<BR>"
      Response.Write "</CENTER>"
      exit sub
    end if
    Response.Write "команда=" & ZAPR_DEL & "<BR>"
Set rs = Nothing

end sub

%>
