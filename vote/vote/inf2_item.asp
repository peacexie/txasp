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
sql = sql& ", InfSubj, SetTop, ImgName" 
sql = sql& ", LogAddIP, LogAUser, LogATime" 
sql = sql& " )VALUES( "
sql = sql& "  '" & Get_AutoID(24) &"'" 
sql = sql& ", '" & ID &"'"  'ModID
sql = sql& ", '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ", " & RequestS("SetTop","N",5120) &"" 
sql = sql& ", '" & RequestS("ImgName","C",24) &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Session("UsrID") &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ")"
  Call rs_DoSql(conn,sql) 
  Response.Redirect "?ID="&ID
ElseIf send="Edt" Then

eID = RequestS("eID",3,48)

sql = " UPDATE [VoteItem] SET " 
sql = sql& " InfSubj = '" & RequestS("InfSubj",3,255) &"'" 
sql = sql& ",ImgName = '" & RequestS("ImgName",3,24) &"'" 
sql = sql& ",InfRem = '" & RequestS("InfRem",3,1020) &"'" 
sql = sql& ",SetTop = '" & RequestS("SetTop","N",5120) &"'" 
sql = sql& " WHERE KeyID='"&eID&"' "
  Call rs_DoSql(conn,sql) 

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
      <LEGEND>调查选项管理 --- <strong><%=bInfSubj%>[<%=bKeyCode%>]</strong> --- 调查日期:<%=bInfTime1%> ~ <%=bInfTime2%> </LEGEND>
      <table width="100%" border="0" align="center" cellspacing="01" bgcolor="#CCCCCC">
        <tr>
          <td align="right" nowrap bgcolor="#FFFFFF">No&nbsp;</td>
          <td bgcolor="#FFFFFF">&nbsp;调查项目 / 具体选项</td>
          <td align="center" bgcolor="#FFFFFF">类型</td>
          <td width="10%" align="center" bgcolor="#FFFFFF">固定位置</td>
          <td width="8%" align="center" bgcolor="#FFFFFF">编辑</td>
          <td width="8%" align="center" bgcolor="#FFFFFF">删除</td>
        </tr>
          <tr>
            <td colspan="6" align="right" style="line-height:1px; background-color:#FFFFFF;">&nbsp;</td>
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
		If i mod 2 = 0 Then
		  col = "#ffffff"
		Else
		  col = "#F0F0F0"
		End If
KeyID = rs("KeyID")
InfSubj = rs("InfSubj")
SetTop = rs("SetTop")
SetVote = rs("SetVote")
InfRem = rs("InfRem")&""
ImgName = rs("ImgName")&""
%>
        <form name="form1<%=i%>" method="post" action="?">
        <tr bgcolor="<%=col%>">
          <td align="right" nowrap><%=i%>&nbsp;</td>
          <td><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="56" maxlength="120">
            <a href="item_edit.asp?ID=<%=ID%>&iID=<%=KeyID%>">
            <input name="eID" type="hidden" id="eID" value="<%=KeyID%>">
            </a></td>
          <td align="center">
            <select name="ImgName" id="ImgName">
              <option value="Select" <%If ImgName="Select" Then Response.Write("selected")%>>[选择]</option>

            </select>          </td>
          <td align="center"><input name="SetTop" type="text" id="SetTop" value="<%=SetTop%>" size="5" maxlength="4"></td>
          <td align="center"><a href="item_edit.asp?ID=<%=ID%>&iID=<%=KeyID%>">
            <input name="ID" type="hidden" id="ID" value="<%=ID%>">
            <input name="send" type="hidden" id="send" value="Edt">
            <input type="submit" name="button2" id="button2" value="编辑">
            </a><a href="item_edit.asp?ID=<%=ID%>&iID=<%=KeyID%>">            </a></td>
          <td align="center"><a onClick="Del_YN('?ID=<%=ID%>&iID=<%=KeyID%>&iImg=<%=ImgName%>&send=Del','确认删除?')" href="#" >删除</a> </td>
        </tr>
        <tr bgcolor="<%=col%>">
          <td align="right" nowrap>选项</td>
          <td colspan="5"><input name="InfRem" type="text" id="InfRem" value="<%=InfRem%>" size="110" maxlength="255" <%If ImgName="FBlank" OR ImgName="GBlank" Then Response.Write("disabled style='background-color:#E0E0E0'")%>></td>
          </tr>
        </form>
        <%
  rs.MoveNext()
  Loop
  rs.Close()
  Set rs = Nothing
%>
          <tr>
            <td colspan="6" align="right" style="line-height:1px; background-color:#FFFFFF;">&nbsp;</td>
          </tr>
        <form action="?" method="post" name="fm01">
          <tr>
            <td align="right" nowrap bgcolor="#FFFFFF">&nbsp;增加&nbsp;</td>
            <td align="left" bgcolor="#FFFFFF"><input name="InfSubj" type="text" id="InfSubj" size="48" maxlength="120"></td>
            <td align="center" bgcolor="#FFFFFF"><select name="ImgName" id="ImgName">
              <option value="Select" >选择</option>

            </select></td>
            <td align="center" bgcolor="#FFFFFF"><input name="SetTop" type="text" id="SetTop" value="5120" size="5" maxlength="4"></td>
            <td colspan="2" align="center" bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send">
              <input type="submit" name="button" id="button" value="增加">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
              <input type="button" name="button3" id="button3" value="预览"></td>
          </tr>
    <tr bgcolor="#FFFFFF">
      <td colspan="6" align="left" bgcolor="#FFFFFF"> &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;<span class="colF0F">注意</span>: 先增加调查项，再编辑具体选项，<span class="col00F">具体选项之间用</span><span class="colF0F"> | </span><span class="col00F">分开</span>；一旦调查投入使用，<span class="col00F">任何增加调查选项或删除调查选项或修改固定顺序的操作都影响调查结果</span>，甚至出错（可以修改选项）！</td>
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
