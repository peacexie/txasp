<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<%

ID = RequestS("ID",3,48)
send = Request("send")

If send = "send" Then
sql = " INSERT INTO [VoteItem] (" 
sql = sql& "  KeyID, KeyMod" 
sql = sql& ", InfSubj, SetTop, SetVote" 
sql = sql& ", LogAddIP, LogAUser, LogATime" 
sql = sql& " )VALUES( "
sql = sql& "  '" & Get_AutoID(24) &"'" 
sql = sql& ", '" & ID &"'"  'ModID
sql = sql& ", '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ", " & RequestS("SetTop","N",5120) &"" 
sql = sql& ", 0" 'SetVote
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("UsrID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
  Call rs_DoSql(conn,sql) 
  Response.Redirect "?ID="&ID
ElseIf send="Del" Then
  iImg = RequestS("iImg",3,120)
  If iImg<>"" Then
    Call fil_del("../../upfile/vote/"&iImg)
  End If 
  Call rs_DoSql(conn,"DELETE FROM VoteItem WHERE KeyID='"&RequestS("iID",3,48)&"'") 
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
bKeyID = rs("KeyID")
bKeyCode = rs("KeyCode")
bInfSubj = Show_Form(rs("InfSubj"))
bInfTime1 = rs("InfTime1")
bInfTime2 = rs("InfTime2")
bInfCard = rs("InfCard")
bInfVNum = rs("InfVNum")
bInfNum1 = rs("InfNum1")
bInfNum2 = rs("InfNum2")
rs.Close()
SET rs=Nothing 
  end if

'If bKeyID = "" Then Response.Redirect("info_list.asp")

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <tr bgcolor="#FFFFFF">
    <td colspan="2" align="left" nowrap bgcolor="#FFFFFF"><FIELDSET style="padding:5px;">
      <LEGEND>
      <table border="0" cellpadding="0" cellspacing="1">
        <tr>
          <td align="center" nowrap><strong><%=bInfSubj%>[<%=bKeyCode%>] </strong> </td>
          <td align="center" nowrap> 　　　　<a href="info_list.asp">返回&lt;&lt;&lt;</a>&nbsp;</td>
        </tr>
      </table>
      </LEGEND>
      <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
        <tr bgcolor="#FFFFFF">
          <td colspan="2" align="center" bgcolor="#FFFFFF"></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td align="center" nowrap bgcolor="#FFFFFF">投票主题</td>
          <td bgcolor="#FFFFFF"><%=bInfSubj%> </td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td align="center" nowrap bgcolor="#FFFFFF">投票日期</td>
          <td bgcolor="#FFFFFF"><%=bInfTime1%> ~ <%=bInfTime2%> 　　　　可选项目 <%=bInfNum1%> ~ <%=bInfNum2%></td>
        </tr>
      </table>
      </FIELDSET></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" align="left" nowrap bgcolor="#FFFFFF"><FIELDSET style="padding:5px;">
      <LEGEND>投票选项管理</LEGEND>
      <table width="100%" border="0" align="center" cellspacing="01" bgcolor="#CCCCCC">
        <tr>
          <td width="10%" align="right" bgcolor="#FFFFFF">No&nbsp;</td>
          <td bgcolor="#FFFFFF">&nbsp;主题</td>
          <td width="5%" align="center" bgcolor="#FFFFFF">选票</td>
          <td width="5%" align="center" bgcolor="#FFFFFF">图片</td>
          <td width="10%" align="center" bgcolor="#FFFFFF">固定位置</td>
          <td width="8%" align="center" bgcolor="#FFFFFF">编辑</td>
          <td width="8%" align="center" bgcolor="#FFFFFF">删除</td>
        </tr>
        <%
    sql = " SELECT * FROM [VoteItem] "
	sql =sql& " WHERE KeyMod='"&ID&"' "
	sql =sql& " ORDER BY SetTop,KeyID " ': Response.Write sqlk'SetTop,
  Set rs=Server.Createobject("Adodb.Recordset")
  i=0
  rs.Open Sql,conn,1,1
  Do While NOT rs.EOF
i=i+1
KeyID = rs("KeyID")
InfSubj = rs("InfSubj")
SetTop = rs("SetTop")
SetVote = rs("SetVote")
ImgName = rs("ImgName")&""
If ImgName<>"" Then
 sImg = "<img src='../../upfile/vote/"&ImgName&"' height=18>"
Else
 sImg = "<font color=red>无图</font>"
End If
%>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><%=i%>&nbsp;</td>
          <td bgcolor="#FFFFFF"><%=InfSubj%></td>
          <td align="center" bgcolor="#FFFFFF"><%=SetVote%></td>
          <td align="center" bgcolor="#FFFFFF"><%=sImg%></td>
          <td align="center" bgcolor="#FFFFFF"><%=SetTop%></td>
          <td align="center" bgcolor="#FFFFFF"><a href="item_edit.asp?ID=<%=ID%>&iID=<%=KeyID%>">编辑</a></td>
          <td align="center" bgcolor="#FFFFFF"><a onClick="Del_YN('?ID=<%=ID%>&iID=<%=KeyID%>&iImg=<%=ImgName%>&send=Del','确认删除?')" href="#" >删除</a> </td>
        </tr>
        <%
  rs.MoveNext()
  Loop
  rs.Close()
  Set rs = Nothing
%>
        <form action="?" method="post" name="fm01">
          <tr>
            <td align="right" bgcolor="#FFFFFF">&nbsp;增加&nbsp;</td>
            <td align="left" bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" size="48" maxlength="120"></td>
            <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF"><input name="SetTop" type="text" id="SetTop" value="5120" size="5" maxlength="4"></td>
            <td colspan="2" align="center" bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send">
              <input type="submit" name="button" id="button" value="增加选项">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>"></td>
          </tr>
        </form>
      </table>
      </FIELDSET></td>
  </tr>
</table>
<script type="text/javascript">


 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj.focus();
     eflag = 1; break;
   }
 
	document.fm01.InfRem1.value = Editor01.getHTML(); ;
 if (document.fm01.InfRem1.value.length==0) 
   {   
     alert(" 内容 不能为空！"); 
     eflag = 1; break;
   }  
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
