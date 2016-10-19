
<%

Function bbsDelID(xID,xRep) 
  Dim rs2,sql,iID
  Call rs_DoSql(conn,"DELETE FROM BBSInfo WHERE KeyID='"&xID&"' ")
  Call del_sfDir(ModTab,xID) 
  If xRep<>"Rep" Then '不指定，检查是否有回复
	 sql = "SELECT KeyID FROM BBSInfo WHERE KeyRE='"&xID&"' "
	 SET rs2=Server.CreateObject("Adodb.Recordset")
	 rs2.Open sql,conn,1,1 
	 Do While NOT rs2.EOF
 	   iID = rs2("KeyID") 
	   Call del_sfDir(ModTab,iID)
	 rs2.MoveNext
	 Loop
	 rs2.Close()
	 SET rs2=Nothing
  End If
End Function 

Function bbsChkUser(xID,xUsr,xFlg) 
  Dim f,u,t
  f = false
  If xFlg="Self" Then
    u = rs_Val("","SELECT LogAUser FROM BBSInfo WHERE KeyID='"&xID&"'")
	If u=xUsr Then
	  f = true
	End If
  End If
  If xFlg="Master" Then
    t = rs_Val("","SELECT InfType FROM BBSInfo WHERE KeyID='"&xID&"'")
	u = " SELECT TypNam2 FROM WebTyps WHERE TypLayer='"&t&"' "
	If inStr(u,xUsr) Then
	  f = true
	End If
  End If
  bbsChkUser = f
End Function 

%>