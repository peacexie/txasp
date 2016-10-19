<!--#include file="_config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../inc/spub.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

  ReEnd = Request("ReEnd")

If Request("send")="ins" Then		
  MemID = RequestS("MemID",C,48)
  MemUrl = Replace(Replace(RequestS("MemUrl","C",255),"http://",""),"https://","")
sql = " INSERT INTO [SmsMember] (" 
sql = sql& "  MemID" 
sql = sql& ", MemCode" 
sql = sql& ", MemName" 
sql = sql& ", MemMobile" 
sql = sql& ", MemBalance" 
sql = sql& ", MemUrl" 
sql = sql& ", MemFlag" 
sql = sql& ", LogAddIP,LogAUser,LogATime"
sql = sql& ")VALUES(" 
sql = sql& "  '" & MemID &"'" 
sql = sql& ", '" & RequestS("MemCode","C",96) &"'" 
sql = sql& ", '" & RequestS("MemName","C",96) &"'" 
sql = sql& ", '" & RequestS("MemMobile","C",48) &"'" 
sql = sql& ", " & RequestS("MemBalance","N",0) &"" 
sql = sql& ", '" & MemUrl &"'" 
sql = sql& ", 'Y'" 
sql = sql& ", '"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"'" 
sql = sql& ")"

   sqlE = "SELECT MemID FROM [SmsMember] WHERE MemID='"&MemID&"' "
   'Response.Write sql
 If rs_Exist(conn,sqlE) = "EOF" Then
   Call rs_DoSql(conn,sql)
   Msg = "增加成功！"
 Else
   Msg = "增加失败，证书编号已经存在！"
 End If
 
  If ReEnd="Y" Then
    Response.Write js_Alert(Msg,"Redir","madd.asp?KW="&KW&"&PG="&PG&"&CrdType="&CrdType&"&ReEnd="&ReEnd&"") 
  Else
    Response.Write js_Alert(Msg,"Redir","member.asp?KW="&KW&"&PG="&PG&"&CrdType="&CrdType&"&ReEnd="&ReEnd&"")   
  End If
 
End If
'Get_rsOpt2(conn,sqli,MemGrade)

'DefID = "Crd"&Year(Now)&"_"&Get_FmtID("mdhnsx","")

%>
<br>
<table width="540" align="center" cellpadding="2" cellspacing="1">
  <form name='fm01' method='post' action='?'>
    <tr>
      <td colspan="2" bgcolor="#999999"></td>
    </tr>
    <tr bgcolor="f4f4f4">
      <td height="24" colspan="2"><strong>&nbsp;短信会员 增加</strong> <font color="#FF0000"><%=Msg%></font></td>
    </tr>
    <tr>
      <td>帐号</td>
      <td><input name='MemID' type='text' size='36' maxlength='24'>
        <input type='button' name='Reset2' value='检查'></td>
    </tr>
    <tr>
      <td>SN</td>
      <td><input name='MemCode' type='text' id="MemCode" value="<%=Get_GUID("2.2","")%>" size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td>名称</td>
      <td><input name='MemName' type='text' size='36' maxlength='48'></td>
    </tr>
    <tr>
      <td>手机</td>
      <td><input name='MemMobile' type='text' id="MemMobile" value="13" size='36' maxlength='24'></td>
    </tr>
    <tr>
      <td>余额</td>
      <td><input name='MemBalance' type='text' id="MemBalance" value="0" size='36' maxlength='6'>
        (条,数字)</td>
    </tr>
    <tr>
      <td>Url</td>
      <td><input name='MemUrl' type='text' id="MemUrl" value="<%=MemUrl%>" size='36' maxlength='120'>
        (外部调用Url)</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input name="ReEnd" type="radio" id="ReEnd1" value="N" <%If ReEnd="N" Then Response.Write("checked")%>>
        添加资料后返回列表
        &nbsp;&nbsp;&nbsp;
        <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
        添加资料后继续</td>
    </tr>
    <tr bgcolor="f4f4f4">
      <td><a href="member.asp">返回</a></td>
      <td bgcolor="f4f4f4"><input name='Button' type='button' class="btn60" onClick="chkData()" value='提交'>
        <input name='Reset' type='reset' class="btn60" value='重设'>
        <a href="?">重载</a>
        <input name='send' type='hidden' id='send' value='ins'></td>
    </tr>
    <tr>
      <td colspan="2">说明: Url中, http:// 或 https:// 省略;<br>
        <span class="fntF00">认证只在Member表中....</span></td>
    </tr>
  </form>
  <tr>
    <td colspan="2" bgcolor="#999999"></td>
  </tr>
</table>
<script type="text/javascript">

 function chkData()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////

 if (document.fm01.MemID.value.length<4) 
   {   
     alert("帐号 至少4位");
	 document.fm01.MemID.focus();
	 eflag = 1; break;
   }

 if (document.fm01.MemName.value.length==0) 
   {   
     alert(" 姓名 不能为空！"); 
     document.fm01.MemName.focus();
     eflag = 1; break;
   }

 //tmv = chkF_Date(document.fm01.CrdTime,"日期 不规范！");
 //if (tmv=='ER') { eflag = 1; break; }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>