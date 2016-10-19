<!--#include file="_config.asp"-->
<%

'PrmFlag,ModID,

ModID = RequestS("ModID","C",48) 
TelID = Trim(RequestS("TelID",3,48))
If ModID="" Then ModID="sysTels"

'yAct = Request("yAct")
send = Request("send") 
Page = RequestS("Page","N",1)
TelID = RequestS("TelID","C",96)

KW = RequestS("KW","C",96)
KT = RequestS("KT","C",96)
TP = RequestS("TP","C",96)

If ModID="sysTels" Then
  ModName = "(系统)群发号码组"
ElseIf ModID="userTels" Then
  ModName = "(用户)群发号码组"
End If 
sqlK = " TelMod='"&ModID&"' "
if KW&"" <> "" then
  If KT="Name" Then
    sqlK = sqlK& " AND ( TelName LIKE '%"&KW&"%' ) "
  ElseIf KT="Num" Then
    sqlK = sqlK& " AND ( TelNum LIKE '%"&KW&"%' ) "
  ElseIf KT="User" Then
    sqlK = sqlK& " AND ( LogAUser LIKE '%"&KW&"%' ) "
  Else
    sqlK = sqlK& " AND ( TelID LIKE '%"&KW&"%' ) " 
  End If
end if

If send="add" Then

  act = "ins"
  acName = "群发号码组 - 增加"
  
ElseIf send="ins" Then 'form

  aNums = getTGrp("tNumb") 'smpFmtTNum(RequestS("TelNums",3,120123)) :echo sNums
  sql = "INSERT INTO [SmsTelq] (TelID,TelMod,TelName,TelNums,TelCount,LogAddIP,LogAUser,LogATime)VALUES"
  sql =sql& "('"&Get_AutoID(24)&"','"&ModID&"','"&RequestS("tName",3,48)&"','"&aNums(2)&"',"&aNums(1)&",'"&Get_CIP()&"','"&PrmUser&"','"&Now()&"')"
  'echo sql
  Call rs_DoSql(conn,sql)
  Msg = "新增成功！"
  acName = "群发号码组 - 增加"

ElseIf send = "edit" Then 'form

  act = "eddo"
   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open "SELECT * FROM [SmsTelq] WHERE TelID='"&TelID&"'",conn,1,1
	if NOT rs.EOF then
		tName = rs("TelName")
		tNumb = rs("TelNums")
		tCount = rs("TelCount")
	end if
  rs.close()
  set rs = nothing
  acName = "群发号码组 - 修改"

ElseIf send = "eddo" Then
  aNums = getTGrp("tNumb")
  sql = " UPDATE [SmsTelq] SET TelName='"&RequestS("tName",3,48)&"',TelNums='"&aNums(2)&"',TelCount='"&aNums(1)&"' "
  sql = sql & " ,LogEditIP='"&Get_CIP()&"',LogEUser='"&PrmUser&"',LogETime='"&Now()&"' "
  sql = sql & " WHERE TelID='"&TelID&"' "
  Call rs_DoSql(conn,sql)
  Msg = "修改成功!"
  acName = "群发号码组 - 修改"

ElseIf send = "tset" Then

  act = "tsdo"
  acName = "电话薄 - 批量增加"
  
ElseIf send="tsdo" Then 

  sqlUBak = " And LogAUser='(user)' "
  If PrmFlag="(Mem)" Then 
	sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
  Else
	sqlUsr2 = ""
  End If

  acName = "电话薄 - 批量增加"
  aNums = getTGrp("tNumb") 
  a = Split(aNums(2),vbcrlf)
  iType = RequestS("TelType",3,48)
  iAuto = Get_AutoID(15)
  For i=0 To uBound(a)
	iExt = Get_9999ID(i+1,6)&Rnd_ID("",3)
	p = inStr(a(i),"[")
	If p>8 Then
	  iTel = Left(a(i),p-1)
	  iName = Mid(a(i),p+1)
	Else
	  iTel = a(i)
	  iName = iExt
	End If
	iName = Replace(iName,"]","")
	sql = "INSERT INTO [SmsTels] (TelID,TelMod,TelType,TelName,TelNum,LogAddIP,LogAUser,LogATime)VALUES"
	sql =sql& "('"&iAuto&iExt&"','"&ModID&"','"&iType&"','"&iName&"','"&iTel&"','"&Get_CIP()&"','"&PrmUser&"','"&Now()&"')"
	If rs_Val(conn,"SELECT TelNum FROM SmsTels WHERE TelNum='"&iTel&"'"&sqlUsr2)="" Then
	  Call rs_DoSql(conn,sql)
	End If
  Next
  Response.Redirect "tels.asp?ModID="&ModID&"&PrmFlag="&PrmFlag&""

ElseIf send = "del" Then
    sql = " DELETE FROM [SmsTelq] WHERE TelID='"&TelID&"' "
	Call rs_DoSql(conn,sql)
	Msg = "删除成功!"

Else

  send="list"
    '
End If

tCount = RequestSafe(tCount,"N",0)

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
<script src="../inc/send_check.js?r=<%=Rnd_ID("",8)%>" type="text/javascript"></script>
</head>
<body>
<%If inStr("(add,edit,tset,tsdo)",send)>0 Then%>
<div class="line15">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <form name="fmGroup" method="post" action="?">
    <tr>
      <td colspan="2" align="center"><%=acName%></td>
    </tr>
    <%If inStr("(add,edit)",send)>0 Then%>
    <tr>
      <td width="15%" align="center">群组名称</td>
      <td><input name="tName" type="text" id="tName" value="<%=tName%>" size="18" maxlength="24">
        共<%=tCount%>个号码</td>
    </tr>
    <%else%>
    <tr>
      <td width="15%" align="center">所属类别</td>
      <td><select name="TelType" id="TelType"><%
			    sqlUBak = " And LogAUser='(user)' "
				If PrmFlag="(Mem)" Then 
				  sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
				Else
				  sqlUsr2 = ""
				End If
			  %>
              <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TelType)%>
        </select>
      </td>
    </tr>
    <%end if%> 
    <tr>
      <td align="center">号码表</td>
      <td><textarea name="tNumb" cols="40" rows="12" id="tNumb" onBlur="fmtTels()"><%=tNumb%></textarea></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" name="button" id="button" value="提交">
        &nbsp;
        <input type="reset" name="button2" id="button2" value="取消">
        <input name="send" type="hidden" id="send" value="<%=act%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
        <input name="TelID" type="hidden" id="TelID" value="<%=TelID%>">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
    </tr>
    <tr>
      <td colspan="2">说明：<br>
        1. 一行一个,姓名用半角[]括起来; 格式:137-1274-8216[Name];<br>
        2. 可从Excel表中直接复制号码表粘贴, <a href="../../upfile/<%=Server.URLEncode("#dbf#")%>/ysWeb_Demo_Sms.xls" target="_blank">下载Excel模版(点击下载或打开)</a>; </td>
    </tr>
  </form>
</table>
<script type="text/javascript">
var fm = document.fmGroup;
function chkData()
{   
  errflag=1
        for(i=0;i<1;i++)
        {  

	 if (document.ff.TelName.value.length==0)
           { alert('[错误]\n 请输入名称!');
             document.ff.TelName.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TelNum.value.length<=5)
           { alert('[错误]\n 号码 太短!');
             document.ff.TelNum.focus();
             errflag=0;
             break;
     }

        }
          if (errflag==1)
          {    
		  document.ff.submit()
          }
}
</script>
<%Else%>
<%
sqlUBak = " And LogAUser='(user)' "
If PrmFlag="(Mem)" Then
  sqlK = sqlK&Replace(sqlUBak,"(user)",PrmUser)
End If
If Left(ModID,4)="user" And PrmFlag<>"(Mem)" Then
  sqlO = " ORDER BY TelID DESC "
Else
  sqlO = " ORDER BY TelName,TelID ASC "
End If
TelID = ""

    sql = " SELECT * FROM [SmsTelq] WHERE "&sqlK&sqlO
   Set rs=Server.Createobject("Adodb.Recordset") ':Response.Write sql
   rs.Open Sql,conn,1,1
   rs.PageSize = 13
   nRecs = rs.RecordCount
if int(Page)>rs.PageCount or int(Page)<1 Then
  Page = 1
End If
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr align="center" bgcolor="E0E0E0">
    <td height="21" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <form action="?" method="post" name="fsch9">
          <tr bgcolor="#FFFFFF">
            <td width="40%" rowspan="2" align="center" nowrap bgcolor="#FFFFFF"><strong><%=ModName%></strong></td>
            <td colspan="2" align="right" nowrap bgcolor="#FFFFFF"><%If PrmFlag="(Mem)" Then%>
              (用户)电话薄
              <%Else%>
              <a 
            href="?ModID=sysTels">(系统)群发号码组</a> | <a 
            href="?ModID=userTels">(用户)群发号码组</a></a>
              <%End If%>
              &nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="center" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font>
              <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>">
              <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
            <td width="30%" align="right" nowrap bgcolor="#FFFFFF"><select name="TP" id="TP">
                <option value="">(所有)</option>
                <%
			  If PrmFlag="(Mem)" Or Left(ModID,3)="sys" Then
			    If PrmFlag="(Mem)" Then 
				  sqlUsr2 = Replace(sqlUBak,"(user)",PrmUser)
				Else
				  sqlUsr2 = ""
				End If
			  %>
                <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM SmsType WHERE TypMod='"&ModID&"' "&sqlUsr2&" ORDER BY TypID",TP)%>
                <%End If%>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="8" maxlength="12">
              <select name="KT" id="KT">
                <option value="">(所有)</option>
                <%=Get_SOpt("Name;Num;User","名称;号码;User",KT,"")%>
              </select>
              <input type="submit" name="Submit3" value="查询"></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <%
  if not rs.eof then
  %>
  <tr bgcolor="#FFFFFF">
    <td colspan="8" align="right"><%=RS_Page(rs,Page,"?send=pag&KW="&KW&"&KT="&KT&"&ModID="&ModID&"&PrmFlag="&PrmFlag&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#e8e8e8">NO</td>
    <td height="27" align="center" bgcolor="#e8e8e8">群组名称</td>
    <td height="27" align="center" bgcolor="#e8e8e8">号码提示</td>
    <td align="center" bgcolor="#e8e8e8">个数</td>
    <td align="center" bgcolor="#e8e8e8">User</td>
    <td align="center" bgcolor="#e8e8e8">Mod</td>
    <td height="27" align="center" bgcolor="#e8e8e8">修改</td>
    <td height="27" align="center" bgcolor="#e8e8e8">删除</td>
  </tr>
  <%
  rs.AbsolutePage = Page
  for i = 1 to rs.PageSize
		if i mod 2 = 1 then
		  col = "#ffffff"
		else
		  col = "#f8f8f8"
		end if
		TelID = rs("TelID")
		TelMod = rs("TelMod")
		TelName = rs("TelName")
		TelNums = rs("TelNums")
		If Len(TelNums)>40 Then
		  TelNums = Left(TelNums,36)&"..."
		End If
		TelNums = Replace(TelNums,vbcrlf,",")
		TelCount = rs("TelCount")
		LogAUser = rs("LogAUser")
		
	  %>
  <form name="ff<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right"><input name="send" type="hidden" id="XXsnd" value="edt">
        <input name="TP" type="hidden" id="XXsnd2" value="<%=TP%>">
        <%=i%></td>
      <td nowrap><input name="TelName" type="text" id="TelName" value="<%=TelName%>" size="18" maxlength="24">
        <input name="TelID" type="hidden" id="TelID" value="<%=TelID%>"></td>
      <td><input name="TelNums" type="text" id="TelNums" value="<%=TelNums%>" size="48" maxlength="24"></td>
      <td align="center"><%=TelCount%></td>
      <td align="center"><%=LogAUser%>
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
      <td align="center"><%=TelMod%>
        <input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td align="center"><a href="?send=edit&ModID=<%=ModID%>&TelID=<%=TelID%>">修改</a></td>
      <td align="center"><input type="button" name="Button" value="删除" 
			    onClick="Del_YN('?send=del&TelID=<%=TelID%>&ModID=<%=ModID%>&KT=<%=KT%>&KW=<%=KW%>&TP=<%=TP%>','<%=Show_jsStr(TelName)%>')"></td>
    </tr>
  </form>
  <%
    rs.movenext
    if rs.eof then exit for
    next
  Else
%>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="27" align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="27" align="center" bgcolor="#FFFFFF">无号码组</td>
    <td height="27" align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="27" align="center" bgcolor="#FFFFFF">&nbsp;</td>
    <td height="27" align="center" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <%
  end if
  rs.Close
  set rs = nothing
	  
	  %>
  <tr bgcolor="#e8e8e8">
    <td align="right" bgcolor="#e8e8e8">&nbsp;</td>
    <td bgcolor="#e8e8e8">&nbsp;</td>
    <td>&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td colspan="2" align="center"><%If nRecs<=120 Then %>
      <a href="?ModID=<%=ModID%>&PrmFlag=<%=PrmFlag%>&send=add">增加</a>
      <%Else%>
      <input type="button" name="Button" value="多于120个，不能新增" disabled>
      <%End If%></td>
  </tr>
  <tr align="left" bgcolor="#FFFFFF">
    <td colspan="10">说明: </td>
  </tr>
</table>
<script type="text/javascript">
function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

	 if (document.ff.TelName.value.length==0)
           { alert('[错误]\n 请输入名称!');
             document.ff.TelName.focus();
             errflag=0;
             break;
     }
	 if (document.ff.TelNum.value.length<=5)
           { alert('[错误]\n 号码 太短!');
             document.ff.TelNum.focus();
             errflag=0;
             break;
     }

        }
          if (errflag==1)
          {    
		  document.ff.submit()
          }
}
</script>
<%End If%>
</body>
</html>
