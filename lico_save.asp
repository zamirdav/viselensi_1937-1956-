<html>
<head>
   <meta http-equiv="content-Type" content="text/html; charset=windows-1251">
   <title>Lico Save</title>
</HEAD>
<BODY bgcolor=lightblue>
<!--#INCLUDE FILE="init.inc"-->
<%
'Response.Write "Lico Save<BR>"
'Response.Write "BobrPA<BR>"
'Response.Write "2021.10.25<BR>"



Dim ID_ZAP, RS_T, KOL_IZM

ID_ZAP = Request.QueryString("ID")

NAM_TAB = "LICO"
NAM_ID  = "L_Id"


sub Write_Rek_0(nRek,znaRekPar)  
    znaRek    = trim(RS_T.Fields(nRek).value)
    Response.Write nREK & " сравниваем " & znaRek  & " и " & znaRekPar & "&nbsp;"
    if znaRek&"~" <> znaRekPar&"~" then
       Z_VVOD = "update "&NAM_TAB&" set " & nRek & "='" & znaRekPar & "' where "&NAM_ID&"=" & ID_ZAP & ";"
       Response.Write " запрос " & Z_VVOD 
       DB.Execute(Z_VVOD) 
       KOL_IZM=KOL_IZM+1
    end if
    Response.Write "<BR>"
end sub                               


sub Write_Rek_CIF0(nRek,znaRekPar)  
    znaRek    = trim(RS_T.Fields(nRek).value)
    Response.Write nREK & " сравниваем " & znaRek  & " & " & znaRekPar & "&nbsp;"
    if znaRek&"~" <> znaRekPar&"~" then
       Z_VVOD = "update "&NAM_TAB&" set " & nRek & "=" & znaRekPar & " where "&NAM_ID&"=" & ID_ZAP & ";"
       Response.Write " запрос " & Z_VVOD 
       DB.Execute(Z_VVOD) 
       KOL_IZM=KOL_IZM+1
    end if
    Response.Write "<BR>"
end sub                               



sub Write_Rek(nRek)  
    A=trim(Request.QueryString(nRek))
    if A&"~"<>"~" then call Write_Rek_0(nRek,A)  
end sub                               


sub Write_Rek_CIF(nRek)  
    A=trim(Request.QueryString(nRek))
    if A&"~"<>"~" then call Write_Rek_CIF0(nRek,A)  
end sub                               




sub Write_FIO(nRek)  
    Z=trim(Request.QueryString(nRek))
'    Z    = trim(RS_T.Fields(nRek).value)
    call Write_Rek_0(nRek,SET_FIO(Z))    
end sub                               



sub Write_IP  
    IP_ADR = Request("REMOTE_ADDR")
'   call Write_Rek_0("IP",IP)  
end sub                               


sub Write_DAT_KORR  
'    DAT2=DATA_LOW(Date())
'    call Write_Rek_0("K_DAT",DAT2)  
'    call Write_Rek_0("K_IP",IP_ADR)  
end sub                               







sub SAVE_LICO
   ZAPR_T = "Select * From "&NAM_TAB&" where "&NAM_ID&"=" & ID_ZAP & ";" 
   Response.Write "ZAPR_T " & ZAPR_T & "<BR>"
   Set RS_T = DB.Execute(ZAPR_T)

   if RS_T.eof then
      Response.Write "не нашлась запись с ID " & ID_ZAP & "<BR>"
      Break
   end if

   KOL_IZM=0

   Call Write_FIO("F1")
   Call Write_FIO("F2")
   Call Write_Rek("F3")
   Call Write_Rek_CIF("PL")
   Call Write_Rek("MR")
   Call Write_Rek_CIF("KU")

   if KOL_IZM>0 then Call Write_DAT_KORR

   Set RS_T = Nothing
end sub





sub SAVE_FOTO_INFO
   NAM_TAB = "FOTO"
   NAM_ID  = "F_Id"

   ZAPR_T = "Select * From "&NAM_TAB&" where "&NAM_ID&"=" & ID_ZAP & ";" 
   Response.Write "ZAPR_T " & ZAPR_T & "<BR>"
   Set RS_T = DB.Execute(ZAPR_T)

   if RS_T.eof then
      Response.Write "не нашлась запись с ID " & ID & "<BR>"
      Break
   end if

   KOL_IZM=0

   Call Write_FIO("TIP")
   Call Write_FIO("FIL_NAME")
   Call Write_Rek_CIF("FIL_SIZE")

   if KOL_IZM>0 then Call Write_DAT_KORR

   Set RS_T = Nothing
end sub






if OPEN_DSN=0 then
   if ID_ZAP&"~" = "~" then 
      Response.Write "нету параметра.<BR>"
   else
      TABL = Request.QueryString("TABL")
      if TABL = "FOTO" then
         call SAVE_FOTO_INFO
      else
         call SAVE_LICO
      end if
   end if
end if

DB.Close
Set DB = Nothing

Response.Write "<HR>"
%>
<A href="index.htm"> Выход </A><BR>
