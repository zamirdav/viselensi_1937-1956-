<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>Lico Import</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="init.inc"-->
<%
'Response.Write "Lico Import<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.27<BR>"

Dim RS
DIM MAC(3333), MAC_MAX, MAC_TEK
dim F1,F2,F3,DR, L_ID






sub FOTO_SAVE_(TIP,NAM,OPIS)    ----- удалить
    On Error Resume Next
    if Err.Number<>0 then
       Response.Write "FOTO_SAVE на входе Err.Number="& Err.Number & "<BR>"
       Err.Clear
       exit sub
    end if
    ZAPROS = "select * from FOTO;" ' where F_Id=0;"
'   Response.Write ZAPROS & "<BR>"
'   Dim RS As ADODB.Recordset
    Set RS = CreateObject("ADODB.Recordset")
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open ZAPROS, DB     ', 1, 3
'   RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
'''       exit sub
    end if
    Response.Write "запрос выполнен<BR>"
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
'   objStream.LoadFromFile(Server.MapPath("./") & "test.jpg" )
    RS.AddNew
    RS.Fields("LIC").Value      = L_ID
    RS.Fields("TIP").Value      = TIP
    RS.Fields("FIL_NAME").Value = NAM
'   RS.Fields("FOT").Value = objStream.Read
    RS.Update
    Response.Write "----- фото " & L_ID & " записано " & TIP & " " & NAM & " " & OPIS & "<BR>" & chr(10)
    objStream.Close
    Set objStream = Nothing
    Set RS = Nothing
end sub



function LICO_ID(F1,F2,F3,DR)
    On Error Resume Next
    LICO_ID=0
    ZAPROS = "select L_ID from LICO where F1 like '"&F1&"%' and F2 like '"&F2&"%' and F3 like '"&F3&"%' and DR like '"&DR&"%';"
'   Response.Write ZAPROS & "<BR>"
    Set RSL = DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit function
    end if
    if not RSL.eof then 
       LICO_ID = RSL.Fields("L_ID").value 
    else
       Response.Write "не найден " & F1 & "<BR>"
    end if
    Set RSL = Nothing
end function




sub LICO_SAVE
    if LICO_ID(F1,F2,F3,DR) <> 0 then
       Response.Write "Lico уже есть в базе<BR>"
       exit sub
    end if
    On Error Resume Next
    ZAPROS = "INSERT INTO LICO(F1,F2,F3,DR) values('"&F1&"','"&F2&"','"&F3&"','"&DR&"');"
    Response.Write ZAPROS & "<BR>"
    DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "Не выполнился " & ZAPROS & "<BR>"
       Err.Clear
       exit sub
    end if
'   Response.Write "запрос выполнен<BR>"
    L_ID=LICO_ID(F1,F2,F3,DR)
    Response.Write "L_ID=" & L_ID & " новое<BR>"
    if L_ID=0 then Response.Write "Lico не записалось в базу<BR>"
end sub




function LINE_RED
   LINE_RED=MAC(MAC_TEK)
'   Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & MAC(MAC_TEK) & "<BR>"  & chr(10)
   MAC_TEK=MAC_TEK+1
end function

sub LINE_BACK
   MAC_TEK=MAC_TEK-1
end sub




sub GET_FOTO
    FIL  = LINE_RED
    OPIS = LINE_RED
    X    = LINE_RED
    TIP  = LINE_RED
    X    = LINE_RED
    X    = LINE_RED
    JPG  = LINE_RED
    X    = LINE_RED
    Response.Write "----- " & TIP & " " & FIL & " <FONT size=1>" & OPIS & "</FONT><BR>" & chr(10)
    if TIP="ФАС" then TIP=1
    if TIP="ПРОФИЛЬ" then TIP=2
    call FOTO_SAVE(L_ID,TIP,"C:\TEMP\kir-roz\",FIL,OPIS)
    X = LINE_RED
end sub

 
sub GET_LICO
    F1 = SET_FIO(LINE_RED)
    F2 = SET_FIO(LINE_RED)
    F3 = SET_FIO(LINE_RED)
    X  = LINE_RED
    DR = mid(X,9,2) & mid(X,4,2) & mid(X,1,2)
    Response.Write "Лицо " & F1 & " " & F2 & " " & F3 & " " & DR & "<BR>" & chr(10)
    call LICO_SAVE
    X = LINE_RED
    X = LINE_RED
    X = LINE_RED
    X = LINE_RED
    A = LINE_RED
    if A = "//ФOB_PHOTO" then call GET_FOTO
    A = LINE_RED
    if A = "//ФOB_PHOTO" then 
       call GET_FOTO
    else
       LINE_BACK
    end if
   X = LINE_RED
end sub



sub RAZBOR_LEGENDA
    NAME_FILE = "C:\TEMP\kir-roz\кир-роз.txt"
    Response.Write " на входе " & NAME_FILE & "<BR>" & chr(10)

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set RSO  = oFSO.OpenTextFile(NAME_FILE, 1)
    A = RSO.ReadLine
    Response.Write A & "<BR>"  & chr(10)
    MAC_MAX=0
    Do while (not RSO.AtEndOfStream)
       MAC_MAX=MAC_MAX+1
       MAC(MAC_MAX) = RSO.ReadLine
    Loop  
    Set RSO = Nothing
    Response.Write "прочитано " & MAC_MAX & " строк." & "<BR>"  & chr(10)
    MAC_TEK=1
    Do while MAC_TEK<MAC_MAX 
      A=LINE_RED
'      Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & A & "<BR>"  & chr(10)
      if A = "//ВO_LICO"   then call GET_LICO
    Loop  
end sub



if OPEN_DSN=0 then
'   call LICO_EXPORT
   call RAZBOR_LEGENDA
else

end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
