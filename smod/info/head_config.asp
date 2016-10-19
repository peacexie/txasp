<%

Function hAdd(xTab,xID)
  Dim sql,KeyMod
  Set rs=Server.Createobject("Adodb.Recordset")
  rs.Open "SELECT * FROM "&xTab&" WHERE KeyID='"&xID&"' ",conn,1,1 
  if NOT rs.eof then 
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfSubj = RequestSafe(rs("InfSubj"),"C",120)
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
InfCont = Show_HText(rs("InfCont"),480)
InfCont = RequestSafe(InfCont,"C",480)
AutoID = Get_AutoID(24)
sql = " INSERT INTO InfoHead (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyRe" 
sql = sql& ", KeyMod" 
sql = sql& ", InfType" 
sql = sql& ", InfTyp2" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", SetSubj" 
sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & AutoID &"'" 
sql = sql& ", '" & xID &"'" 
sql = sql& ", '" & KeyMod &"'" 
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & "" &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & InfCont &"'" 
sql = sql& ", '" & "" &"'" 
sql = sql& ", '" & "" & "'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & LogAUser &"'" 
sql = sql& ", '" & LogATime &"'" 
sql = sql& ")"
Call rs_Dosql(conn,sql)	
  end if 
  rs.Close()
  Set rs = Nothing
  hAdd = AutoID
End Function 

''// 头条类别
Function hOpt(xDef)
  hOpt = Get_rsOpt(conn,"SELECT TypID,TypName FROM WebType WHERE TypMod='InfHead' ORDER BY TypTop,TypID",xDef)
End Function 

''// 头条资料源
OrgTCode = "InfN124;PicS124"
OrgTName = "新闻中心;产品图片"
'"InfoNews|InfoPics"


%>

