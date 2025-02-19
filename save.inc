<%         ' добавление в счётчик - AutoIncrement !

Set db = Server.CreateObject("ADODB.Connection") 
db.Open "ATOC1"                                  
  
dim RS, Name_Zad, ADRES, RS2, nRea, KOL_IZM, RZ_Z


sub Write_Rek(pA, pRek)                            ' из первого файла
   znaRek    = trim(rs.Fields("r"&pRek).value)
   znaRekPar = trim(Request(pRek))
   Response.Write "<TR><TD>" & pA & "<TD>" & znaRek  & "<TD>" & znaRekPar & "<TD>&nbsp;"
   if znaRek&"~" <> znaRekPar&"~" then
      Z_VVOD1 = "update ATOC_" & Name_Zad & "_1 set r" & pRek & "='" & znaRekPar & "' "
      Z_VVOD2 = "where AT_0=" & ADRES & ";"
      Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
      Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
      Set RZ_Z = Nothing
      KOL_IZM=KOL_IZM+1
   end if
end sub                               


sub Write_Mp(pA, pRek, pNomer)                       '  мультиполе
   Set rs2 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & ADRES & ";")
   Kol=0
   Do while NOT Rs2.EOF                             ' корректируем старые
      AdresRec=rs2.Fields("AT_0").value
      Kol=Kol+1
      znaRekPar = Request(pRek & "_" & Kol)
      znaRek=rs2.Fields("r"&pRek).value
      Response.Write "<TR><TD>" & pA & "<TD>" & znaRek  & "<TD>" & znaRekPar & "<TD>&nbsp;"
      if znaRek&"~" <> znaRekPar&"~" then      
         Z_VVOD1 = "update ATOC_" & Name_Zad & "_" & pNomer & " set r" & pRek & "='" & znaRekPar & "' "
         Z_VVOD2 = "where AT_0=" & AdresRec & ";"
         Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
         Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
         Set RZ_Z = Nothing
      end if
      Rs2.MoveNext          
   Loop 
   Set rs2 = Nothing
   Kol=Kol+1
   znaRekPar = Request(pRek & "_" & Kol)            ' добавляем новое если есть
   Response.Write "<TR><TD>" & pA & "<TD>-<TD>" & znaRekPar & "<TD>&nbsp;"
   if (znaRekPar&"~"<>"~")and(znaRekPar<>"0") then                          ' ----- 2020.08.29
      Z_VVOD1 = "insert into ATOC_" & Name_Zad & "_" & pNomer & " (AT_1,r" & pRek & ") "  ' счётчик ! 
      Z_VVOD2 = "values (" & ADRES & ",'" & znaRekPar & "');"
      Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
      Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
      Set RZ_Z = Nothing
   end if
end sub                               


sub Write_RekM(pA, pRek)                                '   мультигруппа
   znaRek    = RS2.Fields("r"&pRek).value
   znaRekPar = Request(nRea & "_" & pRek)
   Response.Write "<TR><TD>" & pA & "<TD>" & znaRek & "<TD>" & znaRekPar & "<TD>&nbsp;"
   if znaRek&"~" <> znaRekPar&"~" then
      Z_VVOD1 = "update ATOC_" & Name_Zad & "_" & nFmulti & " set r" & pRek & "='" & znaRekPar & "' "
      Z_VVOD2 = "where AT_0=" & AdresRecM & ";"
      Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
      Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
      Set RZ_Z = Nothing
   end if
end sub                               


sub Write_MgMp(pA, pRek, pNomer)                        '  мультиполе в мультигруппе
   Set rs3 = db.Execute("Select * From ATOC_"&Name_Zad&"_" & pNomer & " where AT_1=" & AdresRecM & ";")
   Kol=0
   Do while NOT Rs3.EOF                             ' корректируем старые
      AdresRec=rs3.Fields("AT_0").value
      Kol=Kol+1
      znaRekPar = Request(nRea & "_" & pRek & "_" & Kol)
      znaRek    = rs3.Fields("r"&pRek).value
      Response.Write "<TR><TD>" & pA & "<TD>" & znaRek  & "<TD>" & znaRekPar & "<TD>&nbsp;"
      if znaRek&"~" <> znaRekPar&"~" then
         Z_VVOD1 = "update ATOC_" & Name_Zad & "_" & pNomer & " set r" & pRek & "='" & znaRekPar & "' "
         Z_VVOD2 = "where AT_0=" & AdresRec & ";"
         Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
         Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
         Set RZ_Z = Nothing
      end if
      Rs3.MoveNext          
   Loop 
   Set rs3 = Nothing
   Kol=Kol+1
   znaRekPar = Request(nRea & "_" & pRek & "_" & Kol)            ' добавляем новое если есть
   Response.Write "<TR><TD>" & pA & "<TD>-<TD>" & znaRekPar & "<TD>&nbsp;"
   if znaRekPar&"~"<>"~" then
      Z_VVOD1 = "insert into ATOC_" & Name_Zad & "_" & pNomer & " (AT_1,r" & pRek & ") "  ' счётчик !
      Z_VVOD2 = "values (" & AdresRecM & ",'" & znaRekPar & "');"
      Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
      Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
      Set RZ_Z = Nothing
   end if
end sub                               



KOD="0123456789ABCDEF"

Function Trans(byVal STR)
A=Str
P=1
do while P>0
   P=inStr(A,"%")
   if P>0 then 
      Z1=inStr(KOD,mid(A,P+1,1))-1
      Z2=inStr(KOD,mid(A,P+2,1))-1
      if (Z1>-1)and(Z2>-1) then
         A=mid(A,1,P-1) & chr(Z1*16+Z2) & mid(A,P+3,999)
      else
         A=mid(A,1,P-1) & "_" & mid(A,P+1,999)
      end if
   end if
Loop
Trans=A
End Function



sub OPEN_BAZA(pNam)
   Name_Zad=pNam                'Name_Zad=Session("Name_Zad") 
   Response.Write "запись в " & Name_Zad 
   
   ADRES = Request("ADRES")     
   
   ZAPROS = " Select * From ATOC_"&Name_Zad&"_1 where AT_0=" & ADRES & ";"
   Response.Write " <FONT color=white>" & ZAPROS & "</FONT><BR>"
   Set RS = db.Execute(ZAPROS)
end sub
   

sub CLOSE_BAZA
   Set rs = Nothing
   db.Close
   Set db = Nothing
   VID_SAVE = Request("VID_SAVE")     
   Response.Write "VID_SAVE=" & VID_SAVE & "<BR>"
   if VID_SAVE&"~" = "~" then Response.Redirect "VVO-vid-" & Name_Zad & ".asp?N=" & ADRES 
   Response.Write chr(13) & "<BR>" & Trans(request.ServerVariables("QUERY_STRING")) & "<BR>"
   Response.Write chr(13) & "<A href='index.htm'> Выход </A><BR>"

end sub


sub DATA_KORR(FIL)
    if FIL=0 then 
       KOL_IZM=0
    else
      IP  = Request("REMOTE_ADDR") ' ПРОЧИТАТЬ АДРЕС С КАКОГО ИДЕТ корректировка 
      DAT=Date()
      if len(DAT)<10 then DAT="0" & DAT
      DAT=mid(DAT,7,4) & mid(DAT,4,2) & mid(DAT,1,2) 
      Z_VVOD1 = "update ATOC_" & Name_Zad & "_" & FIL & " set AT_K_D='" & DAT & "'," & "AT_K_KIM='"& IP & "' " 
      Z_VVOD2 = "where AT_0=" & ADRES & ";"
      Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
      Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
      Set RZ_Z = Nothing
      TIM=Time()
      if len(TIM)<8 then TIM="0" & TIM
      TIM=mid(TIM,1,2) & mid(TIM,4,2) & mid(TIM,7,2) 
      Z_VVOD1 = "update ATOC_" & Name_Zad & "_" & FIL & " set AT_K_V='" & TIM & "' " 
      Z_VVOD2 = "where AT_0=" & ADRES & ";"
      Response.Write Z_VVOD1 & Z_VVOD2 & "<BR>"
      Set RZ_Z = DB.Execute(Z_VVOD1&Z_VVOD2) 
      Set RZ_Z = Nothing
    end if
end sub

%>

