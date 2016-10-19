<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/func_sfile.asp" -->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.inTab td {
	padding:3px 8px;
}
.combox {
	width:82px;
	font-family:"Courier New", Courier, monospace;
	float:left;
	overflow:inherit;
}
</style>
</head>
<body>
<%

'Reply="(File Read Error!)" Then iReply=""
'Call rs_DoSql(conn,"UPDATE GboInfo SET InfReply='' WHERE InfReply='(File Read Error!)'")

Dim defYear,defRoot,tabPath,tabTabs,aPath,aTabs

defYear = Year(Now())
defRoot = Config_Path&"upfile/"

yAct = Request("yAct")
sPath = Request("sPath")
sTabs = Request("sTabs")
sYear1 = Request("sYear1") : If sYear1="" Then sYear1=defYear-5
sYear2 = Request("sYear2") : If sYear2="" Then sYear2=defYear
fYear = get_fYear(sYear1,sYear2) '"OK"

sTemp = "" ';#dbf#;myfile;myftp;sys;myTest
tabPath = Config_upDir&sTemp
tabTabs = Config_dbTab&sTemp
aPath = Split(Replace(tabPath,"|",";"),";")
aTabs = Split(Replace(tabTabs,"|",";"),";")

shwPath = get_CBox("Path")
shwTabs = get_CBox("Tabs")

'Response.Write "<br>"&sPath&":"&sTabs
'If inStr(sPath,sDir)<=0 Then Response.End() '???


If yAct="DData" Then
  Call rs_DoSql(conn,"DELETE FROM WebTyps WHERE TypMod NOT IN (SELECT SysID FROM AdmSyst)") 'Del Types
  Call rs_DoSql(conn,"DELETE FROM InfoNews WHERE KeyMod NOT IN (SELECT SysID FROM AdmSyst)") 'Del News
  Call rs_DoSql(conn,"DELETE FROM InfoPics WHERE KeyMod NOT IN (SELECT SysID FROM AdmSyst)") 'Del Pics
  'Call rs_DoSql(conn,"DELETE FROM VotItem WHERE VSID NOT IN (SELECT VSID FROM VotSubj)") 'Del Vote
  Msg = "[DData: 清理DB数据] 成功！"  
ElseIf yAct="DDir" Then
  If sPath="" Then
    Msg = "[DDir: 删除目录: 失败！请设置目录！！！ "
  Else
    a = Split(Replace(sPath," ",""),",")
    For i=0 To uBound(a)
      Call fold_del(defRoot&a(i))
    Next 
    Msg = "[DDir: 删除目录("&sPath&")] 成功！!!!刷新后生效!!! "
  End If
ElseIf yAct="SDir" Then
  If sPath="" Or fYear="" Then
    Msg = "[SDir: 清理空目录: 失败！请设置目录和年份！！！ "
  Else
    a = Split(Replace(sPath," ",""),",")
	n0 = 0
    For i=0 To uBound(a)
	For j=sYear1 To sYear2
      n0 = n0 + fSDirs(a(i),j)
	  'Response.Write "<br>"&a(i)&":"&j
    Next 
	Next 
    Msg = "[SDir: 清理空目录("&n0&"个:"&sPath&":"&sYear1&"~"&sYear2&")] 成功！ "
  End If
ElseIf yAct="SFile" Then
  If sTabs="" Or fYear="" Then
    Msg = "[SFile: 删除垃圾文件: 失败！请设置数据表和年份！！！ "
  Else
    a = Split(Replace(sTabs," ",""),",")
	n0 = 0
    For i=0 To uBound(a)
	For j=sYear1 To sYear2
	  n0 = n0 + fSFile(a(i),j)
	  'Response.Write "<br>"&a(i)&":"&j
    Next 
	Next 
    Msg = "[SFile: 删除垃圾文件("&n0&"个:"&sYear1&"~"&sYear2&")] 成功！ "
  End If
ElseIf yAct="CFile" Then
  If sTabs="" Or fYear="" Then
    Msg = "[CFile: 数据转移DB-=>File: 失败！请设置数据表和年份！！！ "
  Else
    a = Split(Replace(sTabs," ",""),",")
	n0 = 0
    For i=0 To uBound(a)
	For j=sYear1 To sYear2
	  n0 = n0 + fConv(a(i),j,yAct)
	  'Response.Write "<br>"&a(i)&":"&j 'fConv(sDir,sTab,sYear,yAct)
    Next 
	Next 
    Msg = "[CFile: 数据转移DB-=>File("&n0&"记录:"&sYear1&"~"&sYear2&")] 成功！ "
  End If
ElseIf yAct="CToDB" Then
  If sTabs="" Or fYear="" Then
    Msg = "[CToDB: 数据转移File-=>DB: 失败！请设置数据表和年份！！！ "
  Else
    a = Split(Replace(sTabs," ",""),",")
	n0 = 0
    For i=0 To uBound(a)
	For j=sYear1 To sYear2
	  n0 = n0 + fConv(a(i),j,yAct)
	  'Response.Write "<br>"&a(i)&":"&j 'fConv(sDir,sTab,sYear,yAct)
    Next 
	Next 
    Msg = "[CToDB: 数据转移File-=>DB("&n0&"记录:"&sYear1&"~"&sYear2&")] 成功！ "
  End If
ElseIf yAct="MOnes" Then
	sql = "SELECT KeyID FROM [GboInfo] WHERE KeyMod='GboU224' "
	Set rs=Server.Createobject("Adodb.Recordset")
	rs.Open sql,conn,1,1 
	cID = rs.RecordCount
	Do While NOT rs.EOF
 	 iID = rs("KeyID") 
	 Call del_sfCont("GboInfo",iID)
	rs.MoveNext
	Loop
	rs.Close()
	Set rs = Nothing
	Msg = "[MOnes: 清理个人留言笔记("&cID&"条)] 成功！ "
ElseIf yAct="^Test^" Then
  Msg = yAct&":"&sPath&":"&sTabs&":"&sYear1&"~"&sYear2&")] ！ "
Else
  Msg = "---"
End If

%>
<br>
<a name="TimTest"></a>

<table width="640"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF">
    <td width="20%"><strong>操作项目</strong></td>
    <td width="70%" align="left"><div style="float:right"> &nbsp; <a href="?">重载本页</a> | <a href="upd.asp">返回&gt;&gt;&gt;</a> &nbsp; </div>
      <strong>说明/设置</strong></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="2"><table width="620" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#F0F0F0" class="inTab">
        <form name="fmScan" method="post" action="?#TimTest">
          <tr>
            <td align="right" bgcolor="#FFFFFF">操作类型</td>
            <td colspan="2" align="left" bgcolor="#FFFFFF"><select name="yAct" id="yAct" style="width:450px" onChange="SetCBox(this)">
                
                <option value="<%=yAct%>" selected><%=yAct%>: [当前]</option>
                <option value="^Test^">[^Test^]</option>  
                <option value="MOnes">MOnes: 个人留言笔记 ------ 不用设置参数</option>
                <option value="DData">DData: 清理DB数据 ------ 不用设置参数</option>
                <option value="DDir">DDir: 删除目录 ------ 选目录</option>
                <option value="SDir">SDir: 清理空目录 ------ 选目录+设置年份</option>
                <option value="SFile">SFile: 删除垃圾文件 ------ 设置数据表+年份</option>
                <option value="CFile">CFile: 数据转移DB-=>File ------ 设置数据表+年份</option>
                <option value="CToDB">CToDB: 数据转移File-=>DB ------ 设置数据表+年份</option>
                
              </select></td>
          </tr>
          <tr id="idDirs">
            <td align="right" bgcolor="#FFFFFF">目标目录 </td>
            <td colspan="2" align="left" bgcolor="#FFFFFF"><%=shwPath%></td>
          </tr>
          <tr id="idTabs">
            <td align="right" bgcolor="#FFFFFF">数据表 </td>
            <td colspan="2" bgcolor="#FFFFFF"><%=shwTabs%></td>
          </tr>
          <tr align="left">
            <td width="12%" align="right" bgcolor="#FFFFFF">设置年份</td>
            <td bgcolor="#FFFFFF"><input name="sYear1" type="text" id="sYear1" value="<%=sYear1%>" size="12" style="width:120px">
              ~
            <input name="sYear2" type="text" id="sYear2" value="<%=sYear2%>" size="12" style="width:120px"></td>
            <td bgcolor="#FFFFFF"><input type="submit" name="button" id="button4" value="执行操作"></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="2" align="left">注意： 1. 当前设置：ConfigCont=<%=Config_Cont%>；<br>
      2.
      ConfigCont[内容存文件] 参数设置入口： 系统与设置  &gt;&gt; 配置 &gt;&gt; ConfigCont；<br>
      3. 提示：<span class="colF00"><%=Msg%></span><br>
      4. 仅限网站初始化时设置使用， 正常运行的系统，请不要随便执行该操作！否则后果自负！！！；</td>
  </tr>
</table>
<script type="text/javascript">

function getElmID(xID){
  return document.getElementById(xID);
}

var fScan = document.fmScan;
var cDirs = getElmID('idDirs');
var cTabs = getElmID('idTabs');

function SetCBox(e)
{
  var v = e.value;
  if((v=='SFile')||(v=='CFile')||(v=='CToDB')) {
	 cDirs.style.display = 'none';
	 cTabs.style.display = '';
  } else if(v=='^Test^') {
	 cDirs.style.display = '';
	 cTabs.style.display = '';
  } else {
	 cDirs.style.display = '';
	 cTabs.style.display = 'none';
  }
}
SetCBox(fScan.yAct);
</script>
<%

Function fConv(xTab,xYear,xAct) 

  Dim xDir,sql0,sql1,n1 : sql0 = "" : n1 = 0
  xDir = Get_SOpt(tabTabs,tabPath,xTab,"Val")
  sql1 = " WHERE KeyID LIKE '"&xDir&"-"&xYear&"%' "
  sql0 = sql0&" SELECT KeyID,InfCont"
  If xTab="GboInfo" Then
    sql0 = sql0&" ,InfReply "
  End If
  sql0 = sql0&"  FROM "&xTab&" "&sql1
  sql0 = Replace(sql0,"/","-")
  'Response.Write sql0&"<br>"&defRoot

  Dim kstr,rs,iKey,iMod : kstr = ""
  Set rs=Server.CreateObject("Adodb.Recordset")
  rs.Open sql0,conn,1,3 
  Do While NOT rs.EOF
   n1 = n1 + 1
   iKey = rs("KeyID")
   iPath = Replace(iKey,"-","/")
   If xAct="CToDB" Then '//////////////////////////////////////// 
     'ReadFile,UpdateDB,DelFile
     If xTab="GboInfo" Then
       iCont = File_Read(defRoot&iPath&".org.htm","utf-8")
       iReply = File_Read(defRoot&iPath&".rep.htm","utf-8")
	   If iReply="(File Read Error!)" Then iReply=""
       Call fil_del(defRoot&iPath&".org.htm")
       Call fil_del(defRoot&iPath&".rep.htm")
       rs("InfCont") = iCont
       rs("InfReply") = iReply
     ElseIf xTab="GboSend" Then
       iCont = File_Read(defRoot&iPath&".out.htm","utf-8")
       Call fil_del(defRoot&iPath&".out.htm")
       rs("InfCont") = iCont
     Else
       iCont = File_Read(defRoot&iPath&"/fcont.htm","utf-8")
       Call fil_del(defRoot&iPath&"/fcont.htm")
       rs("InfCont") = iCont
     End If
   End If
   If xAct="CFile" Then '//////////////////////////////////////// 
     'AddFile,UpdateDB
     If xTab="GboInfo" Then
       Call fold_add9(Config_Path&"upfile/",iKey,1) 
       iCont = rs("InfCont")
       iReply = rs("InfReply")
       Call File_Add2(defRoot&iPath&".org.htm",iCont,"utf-8")
       Call File_Add2(defRoot&iPath&".rep.htm",iReply,"utf-8")
       rs("InfCont") = ""
       rs("InfReply") = ""
     ElseIf xTab="GboSend" Then
       Call fold_add9(Config_Path&"upfile/",iKey,1)
       iCont = rs("InfCont")
       Call File_Add2(defRoot&iPath&".out.htm",iCont,"utf-8")
       rs("InfCont") = ""
     Else
       Call fold_add9(Config_Path&"upfile/",iKey,0)
       iCont = rs("InfCont")
       Call File_Add2(defRoot&iPath&"/fcont.htm",iCont,"utf-8")
       rs("InfCont") = ""
     End If
   End If
   rs.Update()
  rs.MoveNext
  Loop
  rs.Close()
  Set rs=Nothing
  ' Response.End()
  
  fConv = n1
End Function

Function fSFile(xTab,xYear) 

  Dim xDir,sql0,sql1 : sql0 = ""
  xDir = Get_SOpt(tabTabs,tabPath,xTab,"Val")
  sql1 = " WHERE KeyID LIKE '"&xDir&"-"&xYear&"%' "
  If xTab="GboInfo" Or xTab="GboSend" Then
    sql0 = sql0&" SELECT KeyID FROM [GboInfo] "&sql1
    sql0 = sql0&" union "
    sql0 = sql0&" SELECT KeyID FROM [GboSend] "&sql1
  Else '
    sql0 = sql0&" SELECT KeyID FROM ["&xTab&"] "&sql1
  End If
  sql0 = Replace(sql0,"/","-")
  'Response.Write "<br>"&sql0&xDir
  'Response.End()

  Dim kstr,rs,iKey,iMod : kstr = ""
  Set rs=Server.CreateObject("Adodb.Recordset")
  rs.Open sql0,conn,1,1 
  Do While NOT rs.EOF
   iKey = rs("KeyID")
   kstr = kstr&iKey&";"
  rs.MoveNext
  Loop
  rs.Close()
  Set rs=Nothing
  'Response.Write Replace(kstr,";",";<br>")
  ' Response.End()
  
  Dim sYear,sMonth,sDir,p0,n0,n,d,fso,dir
  sYear = xYear
  sDir = xDir
  p0 = defRoot&sDir&"/"&sYear&"/"
  n0 = 0
  If fold_exist(defRoot&sDir,sYear) Then
    set fso = CreateObject("Scripting.FileSystemObject")
    set dir = fso.GetFolder(file_fPath(p0))
    for each d in dir.SubFolders
      n = d.Name
      'Response.Write "<hr>1---"&n&"<br>"&defRoot&sDir&"/"&sYear&"/"&n
      n0 = n0 + fSFil2(defRoot&sDir&"/"&sYear&"/"&n&"/",kstr) 
    next  
    set dir = Nothing
    set fso = Nothing
  End If
  
  fSFile = n0
End Function

Function fSFil2(xPath,xKeys)
    dim fso,dir,f0,d0,f1,f2,sKeys,n1
    n1 = 0
    sKeys = Replace(xKeys,"-","/")&";" ':Response.Write "<br>5-"&k
    set fso = CreateObject("Scripting.FileSystemObject")
    set dir = fso.GetFolder(file_fPath(xPath))
    for each d0 in dir.SubFolders
      f1 = Right(xPath,14)&Left(d0.Name,10)
      If inStr(sKeys,f1&";")<=0 Then
        'Response.Write "<br>3---"&xPath&d0.Name&"/"
        Call fold_del(xPath&d0.Name&"/")
        n1 =  n1 + 1
      End If
    next   
    for each f0 in dir.Files
      f2 = Right(xPath,14)&Left(f0.Name,10)
      If inStr(sKeys,f2&";")<=0 Then
        'Response.Write "<br>4---"&xPath&f0.Name
        Call fil_del(xPath&f0.Name)
        n1 =  n1 + 1
      End If
    next 
    set dir = Nothing
    set fso = Nothing
    fSFil2 = n1
End Function

Function fSDirs(xDir,xYear)
  Dim sYear,sDir,p0,n0,n,d,fso,dir
  sYear = xYear
  sDir = xDir
  p0 = defRoot&sDir&"/"&sYear&"/"
  n0 = 0
  If fold_exist(defRoot&sDir,sYear) Then
    set fso = CreateObject("Scripting.FileSystemObject")
    set dir = fso.GetFolder(file_fPath(p0))
    for each d in dir.SubFolders
      n = d.Name '6Q
	  n0 = n0 + fSDir2(p0,n)
      If fold_Size(defRoot&sDir&"/"&sYear&"/"&n&"/")=0 Then
        Call fold_del(defRoot&sDir&"/"&sYear&"/"&n&"/")
        n0 = n0 + 1
        'Response.Write "<br>-"&defRoot&sDir&"/"&sYear&"/"&n
      End If
    next  
    set dir = Nothing
    set fso = Nothing
  End If
  fSDirs = n0
End Function

Function fSDir2(xPar,xExt)
  Dim sYear,sDir,p0,n0,n,d,fso,dir
    set fso = CreateObject("Scripting.FileSystemObject")
    set dir = fso.GetFolder(file_fPath(xPar&xExt&"/"))
    for each d in dir.SubFolders
      n = d.Name 'J8WF.7GMVU
      If fold_Size(xPar&xExt&"/"&n&"/")=0 Then
        Call fold_del(xPar&xExt&"/"&n&"/")
        n0 = n0 + 1
        'Response.Write "<br>-"&xPar&xExt&"/"&n&"/"
      End If
    next  
    set dir = Nothing
  fSDir2 = n0
End Function

Function get_CBox(xFlag)
Dim i,iFlag,nSel,fSel,s
  For i=0 To uBound(aPath)
    iFlag = fold_exist(defRoot,aPath(i))
	fDis = "" : If NOT iFlag Then fDis=" disabled "
	nSel = Eval("inStr(s"&xFlag&",a"&xFlag&"(i))")
	fSel = "" : If Int(nSel)>0 Then fSel=" checked "
	If xFlag="Path" Then 
      s = s&vbcrlf&"<div class='combox'><input type='checkbox' name='sPath' value='"&aPath(i)&"'"&fSel&""&fDis&">"&aPath(i)&"</div>"
	ElseIf xFlag="Tabs" Then
      s = s&vbcrlf&"<div class='combox'><input type='checkbox' name='sTabs' value='"&aTabs(i)&"'"&fSel&""&fDis&">"&aTabs(i)&"</div>"
	End If
  Next
  get_CBox = s
End Function

Function get_fYear(sYear1,sYear2) 
  Dim fYear : fYear = "OK"
  If (NOT isNumeric(sYear1)) Or (NOT isNumeric(sYear2)) Then
    fYear = ""
  ElseIf sYear1<1900 Then
    fYear = ""
  ElseIf Int(sYear2)<Int(sYear1) Then
    fYear = ""
  End If
  get_fYear = fYear
End Function


%>
</body>
</html>
