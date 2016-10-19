<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../himg/tstyle.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="../himg/tconfig.asp"-->
<!--#include file="oConfig.asp"-->
<script src="../_func/FuncJS.js" type="text/javascript"></script>
<%
Call Chk_Perm1("HomCount","") 
send  = Request("send") 
sAct = Request("sAct") :sAct=Left(sAct,12)
 
sVal = Request("sVal")
If not isNumeric(sVal) Then 
  sVal = 1
End if 

sDate = Request("sDate")
If NOT isDate(sDate) Then
  sDate = "1099-12-31"
End if 

If send = "ins" Then

ElseIf send = "Clear" Then
  If sAct="OnlPage" Then
    sql = " DELETE FROM [OnlPage] WHERE pTime<"&cfgTimeC&""&sDate&""&cfgTimeC&" " 
  Else
    sql = " DELETE FROM [OnlView] WHERE vTime<"&cfgTimeC&""&sDate&""&cfgTimeC&" AND vID<'9999' "  
  End If
	Call rs_DoSql(conn,sql)
	Msg = "清理成功!"
ElseIf send = "Reset" Then
    'sAct = AllView,OnlMax,OnlStart
	If Request("sDate")<>"" Then dSql=",vTime="&cfgTimeC&""&sDate&""&cfgTimeC&""
	sql = " UPDATE [OnlView] SET vNum="&sVal&" "&dSql&" WHERE vID='"&sAct&"'" 
	Call rs_DoSql(conn,sql)
	Msg = "重设成功!"
ElseIf send = "Edi2" Then
	sql = " UPDATE [OnlView] SET vNum=vNum+"&sVal&" WHERE vID='"&sDate&"'" 
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
ElseIf send = "Style" Then
	sVal = Left(Request("sVal"),2)
	sql = " UPDATE [AdmPara] SET ParText='"&sVal&"' WHERE ParCode='Config_NStyle'"
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"
	Response.Write "<IFRAME name=mainFrame src='../../sadm/admin/sys_config.asp?yAct=upd' frameBorder=0 width='0' scrolling='yes' height='1'></IFRAME> "
End If


sDate = DateAdd("d",(-7),Date())
Config_NStyle = rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='Config_NStyle'")
sAllView = rs_Val("","SELECT vNum FROM OnlView WHERE vID='AllView'") 'RequestSafe(,"N",0)
If NOT isNumeric(sAllView) Then 
  sAllView="0"
End If
s = "" ': Response.Write sAllView
If Config_NStyle="00" Then
 s = sAllView
Else
 For i=1 To Len(sAllView)
  c = Mid(sAllView,i,1)
  s = s&"<img src='"&Config_Path&"img/inum/"&Config_NStyle&"0"&c&".gif' border=0>"
 Next
End If

%>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr bgcolor="#FFFFFF">
    <td align="left"><table border="0" cellspacing="1" cellpadding="0">
        <form name="ff2" method="post" action="?">
          <tr>
            <td>清理
              <input name="send" type="hidden" id="send" value="Clear"></td>
            <td><select name="sAct" id="sAct" style="width:150px;">
                <option value="OnlPage" selected="selected">OnlPage.页面记录</option>
                <option value="OnlView">OnlView.日期记录</option>
              </select>
              日期
              <input name="sDate" type="text" id="sDate" value="<%=DateAdd("d",(-7),Date())%>" size="12" />
              (删除此日期之前) </td>
            <td><input type="submit" name="button" id="button" value="修改" /></td>
            <td><a 
          href="oView.asp">查看</a></td>
          </tr>
        </form>
        <form name="ff3" method="post" action="?">
          <tr>
            <td>重设
              <input name="send" type="hidden" id="send" value="Reset"></td>
            <td><select name="sAct" id="sAct" style="width:150px;">
                <option value="AllView" selected="selected">AllView.总访问量</option>
                <option value="OnlMax">OnlMax.最高在线</option>
                <option value="OnlStart">OnlStart.开始日期</option>
              </select>
              日期
              <input name="sDate" type="text" id="sDate" size="12" />
              值=
              <input name="sVal" type="text" id="sVal" value="0" size="6" /></td>
            <td><input type="submit" name="button2" id="button2" value="修改" /></td>
            <td><a 
          href="oTest.asp">测试</a></td>
          </tr>
        </form>
        <form name="ff4" method="post" action="?">
          <tr>
            <td>修改
              <input name="send" type="hidden" id="send" value="Edi2"></td>
            <td><select name="sAct" id="sAct" style="width:150px;">
                <option value="---" selected="selected">增加某访问量</option>
              </select>
              日期
              <input name="sDate" type="text" id="sDate" value="<%=DateAdd("d",(-1),Date())%>" size="12" />
              值+
            <input name="sVal" type="text" id="sVal" value="1" size="6" /></td>
            <td><input type="submit" name="button2" id="button2" value="修改" /></td>
            <td>&nbsp;</td>
          </tr>
        </form>
        <form name="ff5" method="post" action="?">
          <tr>
            <td>风格
              <input name="send" type="hidden" id="send" value="Style"></td>
            <td><input name="sVal" type="text" id="sVal" value="<%=Config_NStyle%>" size="6" />
              <%=s%></td>
            <td><input type="submit" name="button2" id="button2" value="修改" /></td>
            <td><a xhref='../../sadm/admin/sys_config.asp?yAct=upd'>刷新</a></td>
          </tr>
        </form>
    </table></td>
  </tr>
</table>
<div style="line-height:8px;">&nbsp;</div>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">

  <tr>
    <th bgcolor="#FFFFFF">风格</th>
    <th bgcolor="#FFFFFF">样式 - <a href="?">重载本页</a></th>
    <th bgcolor="#FFFFFF">&nbsp;</th>
  </tr>
  <tr>
    <td align="center" bgcolor="#FFFFFF">00</td>
    <td bgcolor="#FFFFFF"><%=sAllView%> (纯文字) 当前计数</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
<%
Randomize
For j=1 To 12
 t = Right("0"&j,2)
 s = "" : n = Int((9999 - 1000 + 1) * Rnd + 1000)
 For i=1 To Len(n)
  c = Mid(n,i,1)
  s = s&"<img src='"&Config_Path&"img/inum/"&t&"0"&c&".gif' border=0>"
 Next
%>  
  <tr>
    <td align="center" bgcolor="#FFFFFF"><%=t%></td>
    <td bgcolor="#FFFFFF"><%=s%> (随机)</td>
    <td bgcolor="#FFFFFF"><%=n%></td>
  </tr>
<%
Next
%>
  <tr>
    <th bgcolor="#FFFFFF">调用</th>
    <td colspan="2" align="left" bgcolor="#FFFFFF">&lt;script src='<%=Config_Path%>tools/base/Online.asp?send=sCount' type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</td>
  </tr>
</table>



</body>
</html>
