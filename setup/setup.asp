<!--#include file="../sadm/func1/func1.asp"-->
<!--#include file="../sadm/func2/func2.asp"-->
<!--#include file="../sadm/func1/func_file.asp"-->
<!--#include file="../sadm/admin/sys.funcs.asp"-->
<!--#include file="../sadm/func1/md5_func.asp"-->
<%

Function SysMFile(xPath,xOrg,xObj)
Dim fso,fObj,fOrg
  Set fso = Server.CreateObject("Scripting.FileSystemObject")	
   If fil_exist(xPath&xOrg) Then
     fObj=Server.MapPath(xPath&xObj)
     fOrg=Server.MapPath(xPath&xOrg)
     fso.MoveFile fOrg,fObj
	 'echo "MFile.1"
   Else
     'echo "MFile.2"
	 'Response.End()
   End If
  Set fso = Nothing
End Function 


'' // 公共值
uNow = Get_vPath(0)
uNow = Mid(uNow,1,inStr(uNow,"/setup/setup.asp"))
'Response.Write "<br>"&uNow

pNow = Get_vPath(0) 'Request.ServerVariables("URL")
pNow = Mid(pNow,1,inStr(pNow,"/setup/setup.asp"))
'Response.Write "<br>"&pNow


Act = Request("Act")
UsrID = RequestS("UsrID","C",96)
UsrPW = RequestS("UsrPW","C",96)
UsrPR = RequestS("UsrPR","C",96)
If Act="Setup" Then
  '' 取出原始值：Config_Code,AppRandom,Config_RAdm,
  If UsrPR<>UsrPW Then
    Response.Write "错误:密码不一致!"
	Response.End()
  End If
  'Response.End()
  OrgConfg = Config_Code
  OrgRead = AppRandom
  OrgApp = AppRand12
  OrgAdm = Config_RAdm
  '' 取出设置值：Set_Path,Set_Confg,Set_CName,Set_Url,Rnd_Read,Rnd_Admin,UsrID,UsrPR,UsrPW,
  Set_Path  = RequestS("Set_Path","C",96)
  Set_Confg = RequestS("Set_Confg","C",96)
  Set_CName = RequestS("Set_CName","C",96)
  Set_Url   = RequestS("Set_Url","C",255)
  Rnd_Read  = RequestS("Rnd_Read","C",96)
  Rnd_App  = RequestS("Rnd_App","C",96)
  Rnd_Admin = RequestS("Rnd_Admin","C",96)
  '' 更改数据库信息
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Set_Path&"'  WHERE ParCode='Config_Path'")
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Set_Confg&"' WHERE ParCode='Config_Code'")
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Set_CName&"' WHERE ParCode='Config_Name'")
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Set_Url&"'   WHERE ParCode='Config_URL'")
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Rnd_Read&"'  WHERE ParCode='AppRandom'")
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Rnd_App&"'   WHERE ParCode='AppRand12'")
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Rnd_Admin&"' WHERE ParCode='Config_RAdm'")
  '' 刷新配置文件
  SET rs=Server.CreateObject("Adodb.Recordset") 
  s = SysConfig()
  Set rs = Nothing 
  s="<"&"%" &s&vbcrlf& "%"&">"
  Call SysMFile("../sadm/",""&OrgAdm&".asp",""&Rnd_Admin&".asp")
  Call SysMFile("../member/","mu_"&OrgRead&".asp","mu_"&Rnd_Read&".asp")
  Call SysMFile("../member/","mu_"&OrgApp&".asp","mu_"&Rnd_App&".asp")
  If cfgDBType = "Access" Then
    '' 更改文件名 'ysWeb_285146793.Peace!DB
    Call SysMFile(Config_Path&"upfile/#dbf#/","ysWeb_"&OrgConfg&".Peace!DB","ysWeb_"&Set_Confg&".Peace!DB")
  End If
  Call File_Add2("../upfile/sys/config/Config.asp",s,"UTF-8")
  '' 下一步，继续更新帐号,密码
  Response.Redirect "setup.asp?Act=SetID&UsrID="&UsrID&"&UsrPW="&UsrPW&"&id1="&OrgConfg&"&id2="&Set_Confg&""
ElseIf Act="SetID" Then
  '' AddLogs
  Call rs_DoSql(conn,"DELETE FROM [AdmLogs]")
  Call Add_Log(conn,Session.SessionID,"初始化安装程序！","[Setup]","安装程序")
  '' Set User IS
  Call rs_DoSql(conn,"DELETE FROM [AdmUser"&Adm_aUser&"] WHERE UsrID IN('"&UsrID&"','peace','websa','sa')")
  eStr = MD5_Adm(UsrPW&UsrID) ':Response.Write UsrPW&"---"&UsrID
  sql = "INSERT INTO [AdmUser"&Adm_aUser&"] (UsrID,UsrName,UsrPW,UsrType,UsrPerm,UsrExp,LogAddIP,LogAUser,LogATime)VALUES"
  sql =sql& "('"&UsrID&"','Setup','"&eStr&"','Admin','{Admin}','2100-12-31','"&Get_CIP()&"','Setup','"&Now()&"')"
  Call rs_DoSql(conn,sql) 
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParDate='"&Date()&"' WHERE ParCode='dateSetup'") '安装日期
  '' 设置成功信息
  Rnd_Admin = Config_RAdm
  Msg = "安装成功！本页处于保护状态……不要刷新本页面。"
  Dis = "disabled"
Else
  Msg = "帐号,密码区分大小写; 帐号&gt;2位,密码&gt;5位。"
  Dis = "---"
End If

fLink = fil_exist(Config_Path&"upfile/#dbf#/ysWeb_"&Config_Code&".Peace!DB")
'Response.Write fLink
  
rRead = lCase(Rnd_ID("0",6))
rApp = lCase(Rnd_ID("0",6))
DefPW = "xyz"&Rnd_ID("0",5)

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>(ASP 2.x)和平鸽建站系统 --- 安装程序</title>
<style type="text/css">
<!--
body, td, th {
	font-size: 12px;
}
a:link {
	color: #00F;
	text-decoration: none;
}
a:visited {
	color: #00F;
	text-decoration: none;
}
a:hover {
	color: #006;
	text-decoration: underline;
}
a:active {
	color: #F00;
	text-decoration: none;
}
-->
</style>
</head>
<body>
<table width="640" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <form name="fm01" method="post" action="?">
    <%If NOT fLink Then %>
    <tr>
      <th colspan="3" align="center" bgcolor="#FFFFFF" style="color:#F00">提示:数据库连接不正确，请先配置 <br>
        /upfile/sys/config/Config.asp <br>
        的 Config_Path 参数</th>
    </tr>
    <%Else%>
    <tr>
      <td align="center" bgcolor="#FFFFFF">提示:</td>
      <td colspan="2" bgcolor="#FFFFFF" style="color:#F00"><%=Msg%>
        <input name="Act" type="hidden" id="Act" value="Setup"></td>
    </tr>
    <%If Dis = "disabled" Then%>
    <tr>
      <td align="right" bgcolor="#FFFFFF">后台地址：</td>
      <td colspan="2" bgcolor="#FFFFFF"><a href="../sadm/<%=Rnd_Admin%>.asp" target="_blank">../sadm/<%=Rnd_Admin%>.asp</a></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFFFF">前台地址：</td>
      <td colspan="2" bgcolor="#FFFFFF"><a href="../" target="_blank">前台</a></td>
    </tr>
    <tr>
      <td colspan="3" align="left" bgcolor="#FFFFFF" style="line-height:180%;">注意：<br>
        1. 请从以上地址登陆后台，分别进入[系统与设置&gt;&gt;配置] 和 [系统与设置&gt;&gt;参数&gt;&gt;系统] 设置更多详细参数, 设置好后分别点 参数页的“刷新”和配置页的“刷新配置”生效。<br>
        2. 在后台，进入[系统与设置&gt;&gt;工具&gt;&gt;安全处理与应用&gt;&gt;开发者说明文件]，得到详细的“开发者说明文件”。<br>
        3. 请记录好你的设置值(如后台地址，管理帐号密码等)，并请删除目录setup；如果今后要重新安装，可以恢复本目录，并把setup.aspPeace文件改名为setup.asp后，重新运行即可。<br>
        4. 系统与设置 &gt;&gt; 配置 &gt;&gt; 1Key配置 一步到位配置一些常用项目。</td>
    </tr>
    <%Else%>
    <tr>
      <th width="33%" bgcolor="#FFFFFF">项目</th>
      <th width="33%" bgcolor="#FFFFFF">原始值</th>
      <th width="33%" bgcolor="#FFFFFF">设置值</th>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">系统路径</td>
      <td bgcolor="#FFFFFF"><input name="txt_01" type="text" id="txt_01" value="<%=Config_Path%>" size="24" maxlength="24"></td>
      <td bgcolor="#FFFFFF"><input name="Set_Path" type="text" id="Set_Path" value="<%=Config_Path%>" size="24" maxlength="48"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">初始化ID</td>
      <td bgcolor="#FFFFFF"><input name="txt_02" type="text" id="txt_02" value="<%=Config_Code%>" size="24" maxlength="24"></td>
      <td bgcolor="#FFFFFF"><input name="Set_Confg" type="text" id="Set_Confg" value="<%=Get_GUID(Config_Vers,"PEACE"&Rnd_ID("0",1)&"ASP"&Rnd_ID("0",1)&"")%>" size="24" maxlength="48"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">站点中文名</td>
      <td bgcolor="#FFFFFF"><input name="txt_03" type="text" id="txt_03" value="<%=Config_Name%>" size="24" maxlength="24"></td>
      <td bgcolor="#FFFFFF"><input name="Set_CName" type="text" id="Set_CName" value="(请填写您的网站名称)" size="24" maxlength="48"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">URL地址</td>
      <td bgcolor="#FFFFFF"><input name="txt_04" type="text" id="txt_04" value="<%=Config_URL%>" size="24" maxlength="24"></td>
      <td bgcolor="#FFFFFF"><input name="Set_Url" type="text" id="Set_Url" value="<%=uNow%>" size="24" maxlength="48"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">防注_Read文件名</td>
      <td bgcolor="#FFFFFF">/member/mu_<br>
        <input name="txt_05" type="text" id="txt_05" value="<%=AppRandom%>" size="18" maxlength="24">
        .asp</td>
      <td bgcolor="#FFFFFF">/member/mu_<br>
        <input name="Rnd_Read" type="text" id="Rnd_Read" value="read_<%=rRead%>" size="18" maxlength="24">
        .asp</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">防注_App文件名</td>
      <td bgcolor="#FFFFFF">/member/mu_<br>
        <input name="txt_05app" type="text" id="txt_05app" value="<%=AppRand12%>" size="18" maxlength="24">
        .asp</td>
      <td bgcolor="#FFFFFF">/member/mu_<br>
        <input name="Rnd_App" type="text" id="Rnd_App" value="app_<%=rApp%>" size="18" maxlength="24">
        .asp</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">管理入口文件[无后缀]</td>
      <td bgcolor="#FFFFFF">/sadm/<br>
        <input name="txt_06" type="text" id="txt_06" value="<%=Config_RAdm%>" size="18" maxlength="24">
        .asp</td>
      <td bgcolor="#FFFFFF">/sadm/<br>
        <input name="Rnd_Admin" type="text" id="Rnd_Admin" value="adm<%=Rnd_ID("0",2)%>" size="18" maxlength="24">
        .asp</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">管理员帐号</td>
      <td bgcolor="#FFFFFF"><input name="UsrID" type="text" id="textfield4" value="sa<%=Rnd_ID("0",3)%>" size="18" maxlength="24">
        (帐号)</td>
      <td bgcolor="#FFFFFF">帐号,密码区分大小写</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF">管理员密码</td>
      <td bgcolor="#FFFFFF"><input name="UsrPR" type="text" id="Rnd_Read3" value="<%=DefPW%>" size="18" maxlength="24">
        (密码)</td>
      <td bgcolor="#FFFFFF"><input name="UsrPW" type="text" id="UsrPW" value="<%=DefPW%>" size="18" maxlength="24">
        (确认密码)</td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFFFF"><a href="?">重载本页</a></td>
      <td align="center" bgcolor="#FFFFFF"><input type="button" name="button" id="button" value="确认安装" <%=Dis%> onClick="chkData()"></td>
      <td align="center" bgcolor="#FFFFFF"><a href="../">前台</a></td>
    </tr>
    <%End If%>
    <%End If%>
  </form>
</table>
<script type="text/javascript">
 function chkData()
 {
       
	   var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.UsrPW.value.length==0) 
   {   
     alert(" 确认密码 不能为空！"); 
     document.fm01.UsrPW.focus();
     eflag = 1; break;
   }

 if (document.fm01.UsrPW.value!=document.fm01.UsrPR.value) 
   {   
     alert(" 确认密码 与 管理员密码 要相同！"); 
     document.fm01.UsrPW.focus();
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		 document.fm01.submit();
		 //Editor01.remoteUpload("document.fm01.submit();"); // 
		 }
}

</script>
</body>
</html>
<%
'' 设置保护状态,不能刷新
If Dis = "disabled" Then
  Call SysMFile("./","setup.asp","setup.aspPeace")
End If
%>
