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


PATH_FILE="C:\TEMP\WWW\������\"    ' <---------- 
'PATH_FILE="C:\TEMP\WWW\������-����\"    ' <---------- 
'NAME_FILE =NAME_FILE & ".xml"


SPISOK_ERROR = ""




sub NEW_LICO_EL_FILE(F1) 
    Z_VVOD = "INSERT INTO LICO(F1,PL,KU) values('" & F1 & "',1,101);"     ' 100 ��� 101 ���
    Response.Write " ������ " & Z_VVOD & "<BR>"
    DB.Execute(Z_VVOD) 
end sub                               






function LICO_TEST(F1)
    LICO_TEST=0
    On Error Resume Next
    ZAPROS = "select * from LICO where F1 like '" & F1 & "%';"
'   Response.Write ZAPROS & "<BR>"
    Set RS=DB.Execute(ZAPROS)
    if Err.Number<>0 then
       Response.Write "�� ���������� " & ZAPROS & "<BR>"
       Err.Clear
       exit function
    end if
    N=0
    Do while NOT RS.EOF 
       N = N + 1     
       ID=RS.Fields("L_ID").value 
'       Response.Write "ID=" & ID & "<BR>"
       RS.MoveNext          
    Loop  
    Set RS = Nothing
    if N>0 then LICO_TEST=-1
    if N=1 then LICO_TEST=ID
end function





sub LICO_ISK(Nam)
    P=InStr(NAM&" "," ")
    F1=mid(Nam,1,P-1)
    ID=LICO_TEST(F1)
    if ID<0 then
       Response.Write "��� " & F1 & ".<BR>"
       SPISOK_ERROR = "��� " & SPISOK_ERROR & NAM & "<BR>"
       exit sub
    end if
    if ID=0 then
'       Response.Write F1 & " - �����<BR>"
       call NEW_LICO_EL_FILE(F1) 
       ID=LICO_TEST(F1)
       if ID=0  then
          Response.Write "�������, " & F1 & " �� �������.<BR>"
       end if
    end if
    call FOTO_SAVE(ID,1,PATH_FILE,NAM,NAM)
'   Response.Write "<A href='lico.asp?T=A'>�������� ����</A><BR>"
end sub







sub RAZBOR_DIR
   Set oFSO  = CreateObject("Scripting.FileSystemObject")
   Set oFolders = oFSO.GetFolder(PATH_FILE)
   Set colFiles = oFolders.Files

   for Each  oFile in colFiles
       Nam=oFile.Name
       call LICO_ISK(Nam)
'       Response.Write F1 & " " & Nam & "<BR>"
   Next
'   Set colFiles = nill
'   Set oFolders = nill
'   Set oFSO  = nill
end sub


if OPEN_DSN=0 then
   call RAZBOR_DIR
   Response.Write "<H2>" & SPISOK_ERROR & "</H2>"
else

end if

'DB.Close
'Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> ����� </A><BR>
