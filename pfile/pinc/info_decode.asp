<%

ID = RequestS("KeyID",3,48) 
ModTab = rel_IDTab(ID)


Dim aImgs
rs.Open "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"' AND SetShow='Y' ",conn,1,3 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = Show_Text(rs("InfSubj"))
InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
SetRead = rs("SetRead")+1
LogATime = rs("LogATime")
xxxCont = Show_Html(rs("InfCont"))
ImgName = rs("ImgName")&"" 
aImgs = Show_sfImgs(ImgName,KeyID)
rs("SetRead") = SetRead
rs.Update() 
  else
KeyID = ""
  end if 
rs.Close()


If KeyID = "" Then 
  Response.End() 
End If


MD = KeyMod '"PicS124"'
TP = InfType '"S110048;S120104;"'
TPLay = TP
MDName = GetMName(MD)
If TP="" Then
 TPName = vPMsg_TName
Else
 TPName = GetNLay(MD,TP,"","","")'GetTName(MD,Replace(TP,";",""),"")'
End If

'rmUrl = "<a href='porder.asp?ObjID="&KeyID&"&ObjSubj="&Server.URLEncode(InfSubj)&"'>["&vPic_Order&"]</a>"
'rmUrl = "<a href='pitems.asp?ID="&KeyID&"&Act=Join&iPrice="&RequestSafe(InfPrice,"N",0)&"&iSubj="&Server.URLEncode(InfSubj)&"'>["&vPic_OAdd&"]</a>"
rmUrl = "<a href='remark.asp?ModID="&MD&"&ObjID="&KeyID&"&ObjSubj="&Server.URLEncode(InfSubj)&"'>["&vInf_Remark&"("&rs_Count(conn,"GboSend WHERE KeyCode='"&ID&"'")&")]</a>"

If Request("_testTempID") <> "" Then
  tmpShow = Request("_testTempID")
Else
  tmpShow = get_TmpID("Show",MD&";"&TP) 
  'tmpShow = "2Pics" '用于测试
  'Response.Write sqlK&ModTab&tmpShow&TypMD
End If



%>

