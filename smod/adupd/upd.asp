<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/cch_Class.asp"-->
<!--#include file="../../upfile/sys/para/rkeyid.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
send = Request("send")

T0Flag = "0.000000"
TMN = Request("TMN") : If TMN = "" Then TMN = 100
CD1 = Request("CD1") : CT1 = Request("CT1")
CD2 = Request("CD2") : CT2 = Request("CT2")
CT1T = CT1 : If CT1 = "" Then CT1T = "[Set Test Code]" 
CT2T = CT2 : If CT2 = "" Then CT2T = "[Set Test Code]"

If send="TimTest" Then
  TF1 = false : TF2 = false
  If CD1<>"" Then
    TN01 = runTimer(CD1,TMN)
	TM01 = getTimeM(TN01,TMN)
	TF1 = true
  End If
  If CD2<>"" Then
    TN02 = runTimer(CD2,TMN)
	TM02 = getTimeM(TN02,TMN)
	TF2 = true
  End If
  If TF1 AND TF2 Then
    If cStr(TN01)=T0Flag AND cStr(TN02)=T0Flag Then
	  TR12 = "0:0"
	  TR21 = "0:0"
	ElseIf cStr(TN01)=T0Flag Then
	  TR12 = "0:999999"
	  TR21 = "999999:0"
	ElseIf cStr(TN02)=T0Flag Then
	  TR12 = "999999:0"
	  TR21 = "0:999999"
	Else
	  TR12 = FormatNumber(TN01/TN02,3,-1)
	  TR21 = FormatNumber(TN02/TN01,3,-1)
	End If
  End If
End If

Function getTimeM(xTime,rTime)
  Dim r,n,nTime
  nTime = FormatNumber(xTime/rTime,6,-1)
  If cStr(nTime)=T0Flag Or cStr(nTime)="0" Then
    r = "0.000001"
	n = "MAX"
  ElseIf Left(cStr(nTime),7)="0.00000" Then
    n = 1000000
  ElseIf Left(cStr(nTime),6)="0.0000" Then
    n = 100000
  ElseIf Left(cStr(nTime),5)="0.000" Then
    n = 10000
  ElseIf Left(cStr(nTime),4)="0.00" Then
    n = 1000
  ElseIf Left(cStr(nTime),3)="0.0" Then
    n = 100
  ElseIf Left(cStr(nTime),2)="0." Then
    n = 10
  Else
    n = 1
  End If
  If cStr(n)<>"MAX" Then
    r = nTime*n
    r = FormatNumber(r,6,-1)
  End If
  getTimeM = n&"x: "&r
End Function

Function runTimer(rData,rTime)
  Dim r,rT1,i
  r = ""
  r = r&vbcrlf&"rT1 = Timer()"
  r = r&vbcrlf&"For i=1 To "&rTime&""
  r = r&vbcrlf&"  "&rData
  r = r&vbcrlf&"Next"
  r = r&vbcrlf&"rT1 = Timer() - rT1"
  Execute(r) 'Eval(r)
  runTimer = FormatNumber(rT1,6,-1)
End Function

fisExpert = Config_Mode

%>
<br>
<a name="TimTest"></a>

<table width="640"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="#FFFFFF">
    <th width="20%">刷新项目</th>
    <td width="70%" align="left"><strong>说明</strong> <span class="colF00">刷新后,可能出现一些提示信息,可以不于理会。</span></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="?send=(Home)"><font color="#FF00FF">刷新首页</font></a></td>
    <td align="left">
    <%If inStr(sUpdItems,"Type")>0 Then%>
    缓存默认 <%=cchTime%>分钟 自动刷新一次；本操作可立即刷新缓存。 
    <%Else%>
    本系统不需要刷新！！！
	<%End If%>
    --- <a href="../../sadm/func2/cch_Admin.asp">缓存管理</a></td>
  </tr>
  <%If inStr(sUpdItems,"Type")>0 Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="../type/type_list.asp?send=upd">刷新类别</a></td>
    <td align="left">当 新闻，图片的大小类别有更改过后使用</td>
  </tr>
  <%End If%>
  <%If inStr(sUpdItems,"Para")>0 Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="../../sadm/system/upd_para.asp">刷新参数</a></td>
    <td align="left">当 参数 有更改过后使用</td>
  </tr>
  <%End If%>
  <%If inStr(sUpdItems,"Link")>0 Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="../link/info_upd.asp">刷新友情连接</a></td>
    <td align="left">当 友情连接 有更改过后使用</td>
  </tr>
  <%End If%>
  <%If inStr(sUpdItems,"Docs")>0 Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="../../doc/dadm/adm_list.asp?yAct=UpdList">刷新内部公文</a></td>
    <td align="left">当 内部公文 的分组和人员 有更改过后使用</td>
  </tr>
  <%End If%>
  <%If inStr(sUpdItems,"Advert")>0 Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td class="col999">刷新广告(图片,文字)</td>
    <td align="left">当 文字/图片广告 有更改过后使用，<span class="colF00">请先进入广告管理页</span>，直接点页面的[更新]即可</td>
  </tr>
  <%End If%>
  <%If fisExpert="isExpert" Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="../type/type_data.asp?Group=Inf"><span class="col00F">数据处理/移动</span></a></td>
    <td align="left">同表不同模块之间，不同类别之间的资料移动</td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="upd_data.asp"><span class="col00F">数据转化/清理</span></a></td>
    <td align="left">清理DB数据,删除目录,清理空目录,删除垃圾文件,DB&lt;-=&gt;File数据转移</td>
  </tr>
  <%End If%>
  <%If send="TimTest" Then%>
  <tr align="center" bgcolor="#FFFFFF">
    <td>效率测试<br>
      <a href="?"><font color="#FF00FF">关闭测试</font></a></td>
    <td align="left"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#F0F0F0">
        <form name="fmTimTest" method="post" action="?#TimTest">
          <tr>
            <td colspan="2" bgcolor="#FFFFFF">Code1
              <select name="CT1" id="CT1" style="width:80px" onChange="CDSet(this,1)">
                <option value="<%=CT1%>"><%=CT1T%></option>
                <option value="Excute(DB)">Excute(DB)</option>
                <option value="Select(DB)">Select(DB)</option>
                <option value="Reade File">Reade File</option>
                <option value="Check File">Check File</option>
                <option value="Server.Execute">Server.Execute</option>
                <option value="[Self Setting]">[Self Setting]</option>
              </select>
              <br>
              <textarea name="CD1" id="CD1" cols="56" rows="3" wrap="off"><%=CD1%></textarea></td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="#FFFFFF">Code2
              <select name="CT2" id="CT2" style="width:80px" onChange="CDSet(this,2)">
                <option value="<%=CT2%>"><%=CT2T%></option>
                <option value="Excute(DB)">Excute(DB)</option>
                <option value="Select(DB)">Select(DB)</option>
                <option value="Reade File">Reade File</option>
                <option value="Check File">Check File</option>
                <option value="Server.Execute">Server.Execute</option>
                <option value="[Self Setting]">[Self Setting]</option>
              </select>
              <br>
              <textarea name="CD2" id="CD2" cols="56" rows="3" wrap="off"><%=CD2%></textarea></td>
          </tr>
          <tr>
            <td width="30%" bgcolor="#FFFFFF">TimeN
              <input name="TMN" type="text" id="TMN" value="<%=TMN%>" size="12"></td>
            <td align="center" bgcolor="#FFFFFF"><input type="submit" name="button2" id="button2" value="效率测试">
              <input name="send" type="hidden" id="send" value="TimTest">
              <input type="button" name="button3" id="button3" value="交换代码" onClick="CDExch()"></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="2"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#F0F0F0">
        <tr>
          <td width="20%" align="right" nowrap bgcolor="#FFFFFF">Code1</td>
          <td width="20%" align="right" nowrap bgcolor="#FFFFFF">TimeN: <%=TN01%></td>
          <td width="20%" align="right" nowrap bgcolor="#FFFFFF"><%=TM01%></td>
          <td width="20%" align="right" nowrap bgcolor="#FFFFFF">R1/R2: <%=TR12%></td>
          <td width="20%" align="center" nowrap bgcolor="#FFFFFF"><%=CT1%></td>
        </tr>
        <tr>
          <td align="right" nowrap bgcolor="#FFFFFF">Code2</td>
          <td align="right" nowrap bgcolor="#FFFFFF">TimeN: <%=TN02%></td>
          <td align="right" nowrap bgcolor="#FFFFFF"><%=TM02%></td>
          <td align="right" nowrap bgcolor="#FFFFFF">R2/R1: <%=TR21%></td>
          <td align="center" nowrap bgcolor="#FFFFFF"><%=CT2%></td>
        </tr>
      </table></td>
  </tr>
  <%Else%>
  <tr align="center" bgcolor="#FFFFFF">
    <td><a href="?send=TimTest"><span class="col00F">效率测试</span></a></td>
    <td align="left"><p>Excute(DB), Select(DB), Reade File, Check File, Server.Execute, Self Setting</p></td>
  </tr>
  <form name="fmDir" method="post" action="">
    <tr align="center" bgcolor="#FFFFFF">
      <td>连接隐藏或扩展的<br>
      管理项目</td>
      <td align="left"><input name="sDir" type="text" id="sDir" value="#../../sadm/system/para_text.asp?ModID=ParText" size="56" maxlength="255">
        <br>
        <select name="select" id="select" onChange="javascript:document.fmDir.sDir.value=this.value" style="width:320px">
          <option value="#">[Items for Manager]</option>
          <option value="../../sadm/system/para_text.asp?ModID=ParText"> ### Text Para ### </option>
          <option value="../../tools/help/xhelp.asp"> ### Help File ### </option>
          <option value="../../smod/type/type_center.asp"> *** 信息类别管理 *** </option>
          <%If fisExpert="isExpert" Then%>
          <option value="../../tools/help/vhelp.asp"> *** Readme *** </option>
          <option value="../../tools/tools.asp?Act=Admin"> *** Admin Links *** </option>
          <option value="../../sadm/admin/sys_config.asp"> *** - Web Config - *** </option>
          <%End If%>
          <%If Chk_PermSP() Then%>
          <option value="../../sadm/user/index.asp?SwhShowSpace=Y"> *** - 显示占用空间 - *** </option>
          <%End If%>
        </select>
        <input type="button" name="button" id="button" value="提交" onClick="goTools()"></td>
    </tr>
  </form>
  <%End If%>
  <tr align="center" bgcolor="#FFFFFF">
    <td colspan="2" align="left" valign="top">说明：&quot;刷新(数据)&quot;的目的是为了提高效率！ 一些更新比较慢的数据，设置好“刷新”后直接从文件读取，比从数据库读取 <span class="colF00">快(效率高)至少10倍</span>！</td>
  </tr>
</table>
<script type="text/javascript">

var fmTest = document.fmTimTest;
function CDExch()
{
  var t = fmTest.CD1.value;
  fmTest.CD1.value = fmTest.CD2.value;
  fmTest.CD2.value = t;
}
function CDSet(e,n)
{
  var o = eval("fmTest.CD"+n);
  var v = e.value; 
  var C = "";
  if(v=="Excute(DB)") { C = 'Call rs_DoSql(conn,"UPDATE AdmSyst SET LogETime=\'<%=Now()%>\' WHERE SysID=\'ModSystem\'")'; }
  if(v=="Select(DB)") { C = 'Call rs_DoSql(conn,"SELECT SysType FROM AdmSyst WHERE SysID=\'ModSystem\'")'; }
  if(v=="Reade File") { C = 'Call File_Read("../../upfile/sys/config/Config.asp","utf-8")'; }
  if(v=="Check File") { C = 'Call fil_exist("../../upfile/sys/config/Config.asp)'; }
  if(v=="Server.Execute") { C = 'Call Server.Execute("../../ext/xtest/uInclude.asp")'; }
  o.value = C;
}

function goTools()
{
  document.fmDir.action = document.fmDir.sDir.value;
  document.fmDir.submit();
}
</script>
<%

If send="(Home)" Then
  s="" 
  sql = "SELECT TOP 8 * FROM [InfoPics] WHERE KeyMod='PicInner' AND InfType LIKE '%In10056%' AND LEN(ImgName)>15 ORDER BY SetTop,LogATime DESC " ' AND SetHot='Y'
  Set rs=Server.CreateObject("Adodb.Recordset")
  rs.Open sql,conn,1,1 
  Do While NOT rs.EOF
   InfSubj = rs("InfSubj")
   ImgName = Replace(Config_Path&rs("ImgName") ,"//","/")
   LogATime = FormatDateTime(rs("LogATime"),2)
   InfSpeci = rs("InfSpeci")
   s = s&vbcrlf&"<item itmDate='"&LogATime&"' itmSubj='"&InfSubj&"' itmUrl='"&InfSpeci&"'>"&ImgName&"</item>"
  rs.MoveNext
  Loop
  rs.Close()
  Set rs=Nothing
  xData =             "<?xml version='1.0' encoding='UTF-8'?>"
  xData = xData&vbcrlf&"<PeacePics>"
  xData = xData&vbcrlf&s
  xData = xData&vbcrlf&"</PeacePics>"
  
  Call cchClear()
  
ElseIf send="TestAddNews" Then
  Call tAddNews()
ElseIf send="updFolder" Then
 'fStr = "../../upfile/test/"
 'Call Upd_fYYYY(fStr)
 fStr = Config_Path&"upfile/" 
 Call Upd_fYYYYMM(fStr)
 'fStr = "../../upfile/test/" 
 'Call Upd_fYYYYMMDD(fStr)
ElseIf send="TestDelNews" Then 
  Call rs_DoSql(conn,"DELETE FROM InfoNews WHERE LogAddIP='(IP)'")
ElseIf send="TestAddPics" Then
  Call tAddPics()
ElseIf send="TestDelPics" Then 
  Call rs_DoSql(conn,"DELETE FROM InfoPics WHERE LogAddIP='(IP)'")
ElseIf send="DelRub2" Then 
 '//
ElseIf send="DelRubbish" Then 
  'Del Types
  Call rs_DoSql(conn,"DELETE FROM WebTyps WHERE TypMod NOT IN (SELECT SysID FROM AdmSyst)")
  'Del News
  Call rs_DoSql(conn,"DELETE FROM InfoNews WHERE KeyMod NOT IN (SELECT SysID FROM AdmSyst)")
  'Call rs_DoSql(conn,"DELETE FROM InfoNews WHERE InfType NOT IN (SELECT TypLayer FROM WebTyps)")
  'Del Pics
  Call rs_DoSql(conn,"DELETE FROM InfoPics WHERE KeyMod NOT IN (SELECT SysID FROM AdmSyst)")
  'Call rs_DoSql(conn,"DELETE FROM InfoPics WHERE InfType NOT IN (SELECT TypLayer FROM WebTyps)")
  'Del Vote
  'Call rs_DoSql(conn,"DELETE FROM VotItem WHERE VSID NOT IN (SELECT VSID FROM VotSubj)")
End If


%>
</body>
</html>
