<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%

TPMax = 20
TPName = rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")

If TPName="" Then
  Response.Write "<br>参数错误！"
  Response.End()
End If

'Response.Write " <hr>更新成功！"
send  = Request("send") 
TypID = Trim(RequestS("TypID",3,48))
DefID = RequestS("DefID",3,48) 

If send = "ins" Then
  TypID = Left(DefID,13)&Right(TypID,2)
  Msg = ""
  sql = "SELECT * FROM [TradeType] WHERE TypID='"&TypID&"' "
  exF2 = rs_Exist(conn,sql)
  If exF2 = "YES" Then
    Msg = "新增失败！["&TypID&"]已经存在！"
  Else
	sql = "INSERT INTO [TradeType] (TypID,TypMod,TypName,TypFlag,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&TypID&"','"&ModID&"','"&RequestS("TypName",3,48)&"','"&RequestS("TypFlag",3,24)&"','"&Get_CIP()&"','"&Session("MemID")&"','"&Now()&"')"
	Call rs_DoSql(conn,sql)
	Msg = "["&TypID&"]新增成功！"
  End If
ElseIf send = "del" Then
    sql = " DELETE FROM [TradeType] WHERE TypID='"&TypID&"' AND LogAUser='"&Session("MemID")&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"
ElseIf send = "edt" Then
    sql = " UPDATE [TradeType] SET TypName='"&RequestS("TypName",3,48)&"',TypFlag='"&RequestS("TypFlag",3,48)&"' "
    sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&Session("MemID")&"',LogETime='"&Now()&"' "
	sql = sql & " WHERE TypID='"&TypID&"' AND LogAUser='"&Session("MemID")&"' "
	Call rs_DoSql(conn,sql)
	Msg = "修改成功!"

ElseIf send="upd" Then
Set rs=Server.Createobject("Adodb.Recordset")

  OldMod = ""
  s1="" : s2="" : s3="" : s4="" :s8="" :s9=""
  s="" : t=""
  
  sql7 = " SELECT * FROM [TradeType] WHERE LogAUser='"&Session("MemID")&"' ORDER BY TypMod,TypID"
  rs.Open sql7,conn,1,1 '3
  Do While NOT rs.EOF 
   TypID = rs("TypID")
   TypName = rs("TypName")
   TypFlag = rs("TypFlag")
   TypMod = rs("TypMod")
   'If TypFlag="Link" Or TypFlag="LInn" Then
     'TypFlag=TypResume
   'End If
   If OldMod="" Then 
     OldMod=TypMod
	 s8=s8&TypMod&"|"
	 s9=s9&rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&TypMod&"'")&"|"
   End If
   If TypMod<>OldMod Then
	 s8=s8&TypMod&"|"
	 s9=s9&rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&TypMod&"'")&"|"
	 s=s&vbcrlf
	 s=s&vbcrlf&"var s"&OldMod&"Code='"&s1&"';"
	 s=s&vbcrlf&"var s"&OldMod&"Name='"&s3&"';"
	 s=s&vbcrlf&"var s"&OldMod&"Flag='"&s4&"';"
     s1="" : s2="" : s3="" : s4=""
	 OldMod=TypMod
   End If 
   s1=s1&TypID&"|"
   s3=s3&TypName&"|"
   s4=s4&TypFlag&"|"
   rs.Movenext
  Loop 
  rs.Close()
  
	 s=s&vbcrlf
	 s=s&vbcrlf&"s"&OldMod&"Code='"&s1&"';"
	 s=s&vbcrlf&"s"&OldMod&"Name='"&s3&"';"
	 s=s&vbcrlf&"s"&OldMod&"Flag='"&s4&"';"
	 
  If rs_Exist(conn,"SELECT ParCode FROM TradePara WHERE ParFlag='Tra_Typ' AND LogAUser='"&Session("MemID")&"' ")="YES" Then
  sql = " UPDATE TradePara SET ParRem='"&Replace(s,"'","''")&"' WHERE ParFlag='Tra_Typ' AND LogAUser='"&Session("MemID")&"' "
  Else
  sql = " INSERT INTO [TradePara] (" 
  sql = sql& "  ParCode,ParMod,ParFlag"  
  sql = sql& ", ParRem" 
  sql = sql& ",LogAddIP,LogAUser,LogATime" 
  sql = sql& ")VALUES(" 
  sql = sql& "  '" & rs_TraTID() &"','UserTemp','Tra_Typ'" 
  sql = sql& ", '" & Replace(s,"'","''") &"'" 
  sql = sql& " ,'"&Get_CIP()&"','"&Session("MemID")&"','"&Now()&"'"
  sql = sql& ")"
  End If
  Call rs_Dosql(conn,sql)
  'Response.Write sql
 
  Response.Write "<hr>"&s&" <hr>更新成功！"
  Response.End()

Set rs = Nothing
Else
    '
End If

TypID = ""

    sql = " SELECT * FROM [TradeType] WHERE TypMod='"&ModID&"' "
	sql =sql& " AND LogAUser='"&Session("MemID")&"' ORDER BY TypID " 
   Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
   rs.Open Sql,conn,1,1
%>
<br>

<table width="720" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="6" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="1" cellspacing="0">
        <tr align="center">
          <td width="30%" rowspan="2" align="center"><strong>[<%=TPName%>]</strong> 类别设置</td>
          <td align="center" nowrap><!--#include file="set_menu.asp"--></td>
        </tr>
        <tr align="center">
          <td align="left" nowrap>&nbsp;<font color="#FF0000">提示:<%=msg%></font></td>
          <form name="form1" method="post" action="?">
          </form>
        </tr>
    </table></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#e8e8e8">NO</td>
    <td height="27" align="center" bgcolor="#e8e8e8">类别代码</td>
    <td height="27" align="center" bgcolor="#e8e8e8">类别名称</td>
    <td align="center" bgcolor="#e8e8e8">选项</td>
    <td height="27" align="center" bgcolor="#e8e8e8">修改</td>
    <td height="27" align="center" bgcolor="#e8e8e8">删除</td>
  </tr>

  <%
  if not rs.eof then
  i = 0
  Do While NOT rs.EOF
  i = i + 1
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f4f4f4"
		end if
		TypID = rs("TypID")
		TypName = Show_Text(rs("TypName"))
		TypFlag = rs("TypFlag")
	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
      <%=i%></td>
      <td><%=Right(TypID,2)%>
        <input name="TypID" type="hidden" id="TypID" value="<%=TypID%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td><input name="TypName" type="text" id="TypName" value="<%=TypName%>" size="24" maxlength="48">
      </td>
      <td align="center"><select name="TypFlag" id="TypFlag">
        <%If ModID="TraA124" Then%>
        <option value="Cont" <%If TypFlag="Cont" Then Response.Write("selected")%>>单信息页</option>
        <option value="FAQ"  <%If TypFlag="FAQ"  Then Response.Write("selected")%>>多信息组</option>
        <option value="NTab" <%If TypFlag="NTab" Then Response.Write("selected")%>>信息列表</option>
        <option value="PTab" <%If TypFlag="PTab" Then Response.Write("selected")%>>图片阵列</option>
        <!-- 
        <option value="Link" <%If TypFlag="Link" Then Response.Write("selected")%>>.外连接.</option>
        <option value="LInn" <%If TypFlag="LInn" Then Response.Write("selected")%>>.内连接.</option>
        -->
        <%Else%>
        <option value="List" selected >[默认]</option>
        <%End If%>
      </select></td>
      <td align="center"><input type="submit" name="Submit" value="修改" >
      </td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TypID=<%=TypID%>&ModID=<%=ModID%>','<%=TypName%>')">
      </td>
    </tr>
  </form>
  <%
  rs.movenext
  Loop
  end if
	  
	  rs.close()
	  set rs = nothing

Function rs_TMemID(xTab,xID,xMD) 
Dim tSql,tVal,sVal
ChrMD = "" 'Mid(xMD,4,1)
Chr12 = Left(xID,13) '12
  If xID="" Then
    tVal = Left(Get_AutoID(12),10)&Base_32(Session.SessionID(),16,3,"Right")&ChrMD&"08" 
	tSql = "SELECT TypID FROM "&xTab&" WHERE TypID='"&tVal&"' "
	If rs_Exist(conn,tSql)="YES" Then
	  tVal = rs_TMemID(xTab,xID,xMD)
	End If
  Else
    tVal = Chr12&Next_ID(Right(xID,2),"08",4) 
  End If
rs_TMemID = tVal
End Function

DefID = rs_TMemID("TradeType",TypID,ModID) '"_"&Rnd_ID("KEY",3)

If i=0 Then
  If ModID="TraA124" Then 
    sName = "企业介绍;在线答疑;建站公告;企业图片;联系我们"
    sFlag = "Cont;   FAQ;    NTab;   PTab;   Cont;企业介绍"
  ElseIf ModID="TraN124" Then 
    sName = "公司新闻;行业新闻"
    sFlag = "List;   List;新闻中心"
  ElseIf ModID="TraJ124" Then 
    sName = "业务部门;技术部门"
    sFlag = "List;   List;职位招聘"
  ElseIf ModID="TraT124" Then 
    sName = "产品一类;产品二类;产品三类;产品四类"
    sFlag = "List;   List;   List;   List;产品供求"
  End If
  aName = Split(sName,";")
  aFlag = Split(sFlag,";")
  For i=0 To uBound(aName)
    iCode = 8 + i*4
    iCode = Left(DefID,13)&Right("0"&iCode,2)
    sql = "INSERT INTO [TradeType] (TypID,TypMod,TypName,TypFlag,LogAddIP,LogAUser,LogATime)VALUES"
    sql =sql& "('"&iCode&"','"&ModID&"','"&aName(i)&"','"&Trim(aFlag(i))&"','"&Get_CIP()&"','"&Session("MemID")&"','"&Now()&"')"
    Call rs_DoSql(conn,sql)
  Next

 If rs_Exist(conn,"SELECT ParCode FROM TradePara WHERE ParText='"&ModID&"' AND LogAUser='"&Session("MemID")&"'")="EOF" Then
  If ModID="TraJ124" Then
    sTemp = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmpInfJ124'")
  ElseIf ModID="(xxxx)" Then
    sTemp = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmpInfxxxx'")
  Else
    sTemp = ""
  End If
  If sTemp<>"" Then
   sTemp = Replace(sTemp,"'","")
   sql = " INSERT INTO [TradePara] (" 
   sql = sql& " ParCode, ParName, ParText, ParFlag, ParRem" 
   sql = sql& ",LogAddIP,LogAUser,LogATime" 
   sql = sql& ")VALUES(" 
   sql = sql& " '"&Left(DefID,12)&Rnd_ID("KEY",3)&"模版', '" & aFlag(i) &"', '" & ModID &"', 'UserTemp', '" &aFlag(i)&"模版：会员后台：[供求与商务>企业资料-参数>关联"&ModID&"]中设置"&sTemp&"'" 
   sql = sql& " ,'"&Get_CIP()&"','"&Session("MemID")&"','"&Now()&"'"
   sql = sql& ")" ':Response.Write sql
   Call rs_DoSql(conn,sql)
   sMsg = "\n导入默认模版成功，你可以根据需要到[参数]中自己设置！"
  Else
   sMsg = ""
  End If
 End If
  
   Response.Write js_Alert("导入系统默认类别成功！"&sMsg,"Redir","?ModID="&ModID)
  
  
End If
	  

	  
	  %>

  <form name="ff" method="post" action="?">
    <tr bgcolor="#e8e8e8">
      <td align="right" bgcolor="#e8e8e8"><input name="send" type="hidden" id="send" value="ins">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="DefID" type="hidden" id="DefID" value="<%=DefID%>">
        新增 </td>
      <td bgcolor="#e8e8e8"><input name="TypID" type="text" id="TypID" value="<%=Right(DefID,2)%>" size="18" maxlength="2"></td>
      <td><input name="TypName" type="text" id="TypName" size="24" maxlength="48"></td>
      <td align="center"><select name="TypFlag" id="TypFlag">
        <%If ModID="TraA124" Then%>
        <option value="Cont" >单信息页</option>
        <option value="FAQ"  >多信息组</option>
        <option value="NTab" selected >信息列表</option>
        <option value="PTab" >图片阵列</option>
        <!--
        <option value="Link" >.外连接.</option>
        <option value="LInn" >.内连接.</option>
        -->
        <%Else%>
        <option value="List" selected >[默认]</option>
<%End If%>
            </select></td>
      <td colspan="2"><input type="button" name="Button" value="新增类别" onClick="chkData()" <%If (i>TPMax) Then Response.Write("disabled")%>>
      </td>
    </tr>
  </form>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="8"><span class="colF03">注意</span>: 最多可以设置<span class="col00F">20</span>个类别！请分别设置如下类别： <a 
            href="?ModID=TraA124" >企业介绍</a> | <a 
            href="?ModID=TraN124" >行业新闻</a> | <a 
            href="?ModID=TraT124" >产品供求</a> | <a 
            href="?ModID=TraJ124" >职位招聘</a><br>
    建议使用自动编号; 代码按大小排列,可在中间插入类别代码; </td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  
	 if (document.ff.TypID.value.length==0)
           { alert('[错误]\n 类别代码不能为空!');
             document.ff.TypID.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TypName.value.length==0)
           { alert('[错误]\n 请输入名称!');
             document.ff.TypName.focus();
             errflag=0;
             break;
     }
	 //var tID = document.ff.TypID.value; 
	 //if ( (tID.substring(0,13)!="<%=Left(DefID,13)%>") )
           //{ alert('[错误]\n 代码请用"<%=Left(DefID,13)%>"开头!');
             //document.ff.TypID.focus();
             //errflag=0;
             //break;
     //}
        }
          if (errflag==1)
          {    
		  document.ff.submit()
          }
}
  
                </script>
</body>
</html>
