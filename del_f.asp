<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>удаление </title>
</HEAD>
<BODY bgColor=Lavender>
<!--#INCLUDE FILE="edit.inc"-->

<%

'ID = Request.QueryString("ID") 


sub RECORD_DEL
   ZAPROS="delete from ATOC_sp_F where AT_1='" & ID & "' ;"

   On Error Resume Next  
   Set RS = DB.Execute(ZAPROS) 
   if db.Errors.Count > 0 then
      Response.Write "ошибка обращайтесь Довлетбаев 0550102614 " & ZAPROS & "<BR>" 
      Set RS = Nothing
   end if
   Response.Write ZAPROS&"<BR>"	
end sub                  


ID = Request.QueryString("ID")
'FIL_NAME = Request.QueryString("FIL_NAME")
'FIL_SIZE = Request.QueryString("FIL_SIZE")
'F_NAME   = Request.QueryString("F_NAME")

Response.Write "ID="&ID&"<BR>"
'Response.Write "FIL_NAME="&FIL_NAME&"<BR>"
'Response.Write "FIL_SIZE="&FIL_SIZE&"<BR>"
'Response.Write "F_NAME="&F_NAME&"<BR>"


call RECORD_DEL


'DB.Close
'Set DB = Nothing

Response.Write "<HR>"

Response.Write "<A href='sp--form.asp?T=1'>Выход</A><BR>"  ' ДЕЛАЛ НА УРОВЕНЬ ВВЕРХ
%>          





