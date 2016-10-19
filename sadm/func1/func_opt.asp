

<%

Function Get_TypeLable(xModID)
  Dim strCode,r
  strCode = Eval(xModID&"TypeCode")
  If strCode<>"" Then
	r = Eval(xModID&"TypeLable")
  Else
    r = "<a href='"&Config_Path&"smod/type/type_list.asp?ModID="&xModID&"' target='_blank'>类别</a>"
  End If
  Get_TypeLable = r
End Function
Function Get_TypeName(xModID,xType)
  Dim strCode,r,sql
  strCode = Eval(xModID&"TypeCode")
  If strCode<>"" Then
    If Left(strCode,5)= "(Sys)" Then
	  sql = "SELECT TypName FROM WebType WHERE TypID='"&xType&"' ORDER BY TypID "
	  r = rs_Val("",sql) 
	Else
	  ' ?? r = rs_Val(conn,"SELECT TypName FROM WebTyps WHERE TypLayer='"&xType&"'")
	End If
  Else
	  r = rs_Val("","SELECT TypName FROM WebTyps WHERE TypLayer='"&xType&"'")
  End If
  If r="" Then
    r = "<font color=gray>(无类别)</font>"
  End If
  Get_TypeName = r
End Function
Function Get_TypeOpt(xModID,xType) 
  Dim strCode,fLable,sOpt,sList
  Dim aCode,aName,aLay,i,fSel,sql
  sOpt = "" : sList = "" 
  strCode = Eval(xModID&"TypeCode")
  If strCode<>"" Then
    fLable = Eval(xModID&"TypeLable")
	If Left(strCode,5)= "(Sys)" Then
	  sql = "SELECT TypID,TypName FROM WebType WHERE TypMod='"&Mid(strCode,6)&"' ORDER BY TypID "
	  sOpt = Get_rsOpt(conn,sql,xType)
	  sOpt = "<option value=''>[不限]</option>"&sOpt
	Else
      ' ??sOpt = Get_SOpt(strCode,Eval(xModID&"TypeName"),xType,"")
    End If
  Else
    sOpt = Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&xModID&"' ORDER BY TypLayer ",xType,"Lay")
  End If
Get_TypeOpt = sOpt
End Function 


Function Get_Typ2Name(xModID,xTyp2)
  Dim strCode,sCode,sName,r,sql
  strCode = Eval(xModID&"Typ2Code")
  If strCode<>"" Then
    If strCode = "(Depart)" Then
	  sCode = Replace(strDepCode,"|",";")
    ElseIf Left(strCode,4)= "(rt)" Then
	  r = rs_Val("","SELECT TypName FROM WebTyps WHERE TypMod='"&Mid(strCode,5)&"' AND TypID='"&xTyp2&"' ")
    ElseIf Left(strCode,4)= "(rs)" Then
	  r = rs_Val("","SELECT TypName FROM WebType WHERE TypMod='"&Mid(strCode,5)&"' AND TypID='"&xTyp2&"' ")
    ElseIf Left(strCode,5)= "(Mem)" Then
	  r = rs_Val("","SELECT TypName FROM "&Mid(strCode,6)&" WHERE TypMod='"&xModID&"' AND TypID='"&xTyp2&"' AND LogAUser='"&Session("MemID")&"' ORDER BY TypID ") 
	Else
	  r = Get_SOpt(strCode,Eval(xModID&"Typ2Name"),xTyp2,"Val")
	End If
  End If
  Get_Typ2Name = r
End Function 
Function Get_Typ2Opt(xModID,xTyp2) 
  Dim strCode,fLable,sOpt,sList
  Dim aCode,aName,aLay,i,fSel,sql
  sOpt = "" : sList = "" 
  strCode = Eval(xModID&"Typ2Code")
  If strCode<>"" Then
    sOpt = "<select name='InfTyp2' id='InfTyp2'>"&vbcrlf&"<option value=''>[不限]</option>"&"($List$)"&vbcrlf&"</select>"
    fLable = Eval(xModID&"Typ2Lable")
	If strCode = "(Depart)" Then
      aCode = Split(strDepCode,"|")
      aName = Split(strDepName,"|")
      aFlag = Split(strDepFlag,"|")
      For i=0 To uBound(aCode)-1
	     If (inStr(Session(UsrPStr),"("&aCode(i))>0 AND Len(aCode(i))>4) Or (inStr(Session(UsrPStr),"{Admin}")>0) Then
           fSel = " " : If aCode(i) = xTyp2 Then fSel = " selected "
	       sList = sList&vbcrlf&"<option value="&aCode(i)&fSel&">"&aName(i)&"</option>"
         End If  
      Next
	ElseIf Left(strCode,4)= "(rt)" Then
	  sql = "SELECT TypID,TypName,TypLayer FROM WebTyps WHERE TypMod='"&Mid(strCode,5)&"' ORDER BY TypLayer,TypID "
	  sList = Get_rsTree(conn,sql,xTyp2,"") 
	  sOpt = Replace(sOpt,"<option value=''>[不限]</option>","")
    ElseIf Left(strCode,4)= "(rs)" Then
	  sql = "SELECT TypID,TypName FROM WebType WHERE TypMod='"&Mid(strCode,5)&"' ORDER BY TypID "
	  sList = Get_rsOpt(conn,sql,xTyp2)
    ElseIf Left(strCode,5)= "(Mem)" Then
	  sql = "SELECT TypID,TypName FROM "&Mid(strCode,6)&" WHERE TypMod='"&xModID&"' AND LogAUser='"&Session("MemID")&"' ORDER BY TypID "
	  sList = Get_rsOpt(conn,sql,xTyp2)
	Else
      sList = Get_SOpt(strCode,Eval(xModID&"Typ2Name"),xTyp2,"")
    End If
	sOpt = Replace(sOpt,"($List$)",sList)
	sOpt = fLable & sOpt
  End If
Get_Typ2Opt = sOpt
End Function 

Function Get_rsOpt(xconn,xSql,xDef)  
   Dim cnFPub,rsOpt,objOpt,sql,sID,sName,sTemp,fSel
   xType =UCase(xType)
   Set cnFPub = Server.CreateObject("ADODB.Connection")
   cnFPub.Open xconn
   Set rsOpt=Server.Createobject("Adodb.Recordset")
   rsOpt.Open xSql,cnFPub,1,1    
   objOpt = ""
   Do While NOT rsOpt.EOF 
     sID   = Trim(""&rsOpt(0))
	 sName = Trim(""&rsOpt(1))
	   fSel = ""
	 If xDef=sID Then
	   fSel = " selected "
	 End If
	 objOpt   = objOpt&vbcr& "<option value="&sID&fSel&">"&sName&"</option>"
     rsOpt.movenext
   Loop 
   rsOpt.close()
   set rsOpt = nothing	
   cnFPub.Close()
   SET cnFPub = Nothing
   Get_rsOpt = objOpt  
End Function 
Function Get_rsOpt2(xconn,xSql,xDef)  
   Dim cnFPub,rsOpt,objOpt,sql,sID,sName,sTemp,fSel
   xType =UCase(xType)
   Set cnFPub = Server.CreateObject("ADODB.Connection")
   cnFPub.Open xconn
   Set rsOpt=Server.Createobject("Adodb.Recordset")
   rsOpt.Open xSql,cnFPub,1,1    
   objOpt = ""
   Do While NOT rsOpt.EOF 
     sID   = Trim(""&rsOpt(0))
	 sName = Trim(""&rsOpt(1))
	   fSel = ""
	 If inStr(xDef,sID) > 0 Then
	   fSel = " selected "
	 End If
	 objOpt   = objOpt&vbcr& "<option value="&sID&fSel&">"&sName&"</option>"
     rsOpt.movenext
   Loop 
   rsOpt.close()
   set rsOpt = nothing	
   cnFPub.Close()
   SET cnFPub = Nothing
   Get_rsOpt2 = objOpt  
End Function 

Function Get_rsCBox(xconn,xSql,xDef)  
   Dim cnFPub,rsCBox,objCBox,sql,sID,sName,sTemp,fSel
   xType =UCase(xType)
   Set cnFPub = Server.CreateObject("ADODB.Connection")
   cnFPub.Open xconn
   Set rsCBox=Server.Createobject("Adodb.Recordset")
   rsCBox.Open xSql,cnFPub,1,1    
   objCBox = ""
   Do While NOT rsCBox.EOF 
     sID   = Trim(""&rsCBox(0))
	 sName = Trim(""&rsCBox(1))
	 If inStr(xDef,sID) > 0  Then
	   fSel = " checked "
	 Else
	   fSel = ""
	 End If
	 objCBox   = objCBox&vbcr& "<input type='checkbox' name='[ChkBox01]' value='"&sID&"'"&fSel&">"&sName&"<br>"
     rsCBox.movenext
   Loop 
   rsCBox.close()
   set rsCBox = nothing	
   cnFPub.Close()
   SET cnFPub = Nothing
   Get_rsCBox = objCBox  
End Function 

Function Get_rsTree(xconn,xSql,xDef,xStyle)  
Dim i,j,k,s1,rs,TypID,TypCode,TypLayer,TypName,nLayer,OutTypID
s1 = "<option value='"&xRoot&"'>┯ [ROOT]</option>"
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open xSql,xconn,1,1
  Do While Not rs.eof 
    TypID = rs("TypID")
	TypLayer = rs("TypLayer")
    TypName = rs("TypName")
	If xStyle="ID2" Or xStyle="Lay2" Then
	 TypName = rs("TypNam2") 
	End If
		nLayer = Len(TypLayer)-Len(Replace(TypLayer,";",""))
		nLayStr = ""
		If nLayer > 0 Then
		  For j = 1 To nLayer-1
		    nLayStr = nLayStr & "│ "
		  Next
		End If
		nLayStr = nLayStr & "├─"
    If xStyle="Lay" Or xStyle="Lay2" Then  ' ID  
	  OutTypID = TypLayer
	Else
	  OutTypID = TypID
	End If
	fSel = ""
    If (TypID=xDef)Or(TypLayer=xDef) Then 
	  fSel = " selected "
	End If
    s1 = s1&vbcrlf&"<option value='"&OutTypID&"'"&fSel&">"&nLayStr&TypName&"</option>"
  rs.movenext
  loop  
rs.close()
set rs = nothing
Get_rsTree = s1
End Function 

%>
