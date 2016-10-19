<%
 
DefMD = "InfN124" 
DefTP = "" 
DefDP = ""

MD = RequestS("ModID",3,48)
DP = RequestS("DepID",3,48) 
TP = RequestS("TypID",3,240)
If DP<>"" AND MD="" Then
  MD="InfD124"
End If

If MD="InfA124" Then 
 DefTP="A110012"
End If
If MD="InfC124" Then 
 DefTP="C110024"
 DefDP = "ClsTyp212"
End If
If MD="InfD124" Then 
 DefTP="D110024"
 DefDP = "DepTech"
End If
 
KW = RequestS("KeyWD",3,24)
Page = RequestS("Page","N",1)
Flag = Request("Flag")
TPLay = Request("TPLay") ' 折叠菜单（全部）
TypMD = Request("TypMD") ' 折叠菜单(部分)
ID = RequestS("KeyID",3,48)

If MD="" AND DefMD&""<>"" Then 
  MD=DefMD
End If
'If DP="" AND DefDP&""<>"" Then 
  'DP=DefDP
'End If
MDName = GetMName(MD)
ModTab = rel_ModTab(MD)
DPName = Get_Typ2Name(MD,DP)
'Response.Write MD&"-"&DP&"-"&TP&"-"&DPName


If TP="" AND DefTP&""<>"" Then 
  'TP = rs_Val("","SELECT TOP 1 TypID FROM WebTyps WHERE TypMod='"&MD&"' ORDER BY TypID ")
  'aTP = Split(Eval("s"&MD&"Code"),"|") : TP = aTP(0)
  TP = DefTP
End If
If TP="" Then
 TPName = vPMsg_TName
Else
 'TPName = GetTName(MD,TP,"")  ' 参数["Lay",""] 得到 类别名称 或 Lay字串
 TPName = GetNLay(MD,TP,"","","")     '                得到 类别名称 含级别：衣(服饰)系列 >> 童装 >> 
End If
'Response.Write " :: "&MDName&" &gt;&gt; "&TPName&" &gt;&gt; "&MD&" &gt;&gt; "&TP

sqlK = " WHERE ( KeyMod='"&MD&"' ) AND SetShow='Y' " 
If KW&"" <> "" Then
  sqlK = sqlK & " AND ( InfSubj LIKE '%"&KW&"%' ) " 
End If
If TP&"" <> "" Then
  sqlK = sqlK & " AND ( InfType LIKE '%"&TP&"%' ) " 
End If
If DP&"" <> "" Then
  sqlK = sqlK & " AND ( InfTyp2='"&DP&"' ) " 
End If

If Request("_testTempID") <> "" Then
  tmpList = Request("_testTempID")
  '&_testTempID=News
Else
  tmpList = get_TmpID("List",MD&";"&TP) 'Cont,Next,FAQ,NList,PList,News,Pics,Jobs,Vdos
  'tmpList = "PicC" 
  'Response.Write sqlK&ModTab&tmpList&TypMD
End If

%>

