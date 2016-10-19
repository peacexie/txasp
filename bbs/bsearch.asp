<!--#include file="binc/_config.asp"-->
<!--#include file="../sadm/func1/func_opt.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>帖子搜索 - <%=bbsName%></title>
<link rel="stylesheet" type="text/css" href="bimg/style.css">
</head>
<body>
<!--#include file="_itop.asp"-->
<div align="center" style="width:980px; height:auto; margin:auto; background-color:#F2F6FB">
  <table width="980" border="0" align="center" cellpadding="1" cellspacing="8">
    <tr>
      <td align="left" bgcolor="#FFFFFF">&nbsp;<img src="bimg/face1.gif" align="absmiddle" />&nbsp;<a href="../">首页</a> &gt;&gt; <a href="bind.asp"><%=bbsName%></a> &gt;&gt; 帖子搜索 &gt;&gt; </td>
      <td width="50%" align="right" bgcolor="#FFFFFF">&nbsp; &nbsp;<a href="bind.asp">返回首页</a>&nbsp;&nbsp;</td>
    </tr>
  </table>
  <table width="720"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <form name="fm01" id="fm01" action="blist.asp" method="post">
      <tr>
        <td align="center" bgcolor="#F2F6FB">关 键 字</td>
        <td align="left" bgcolor="#F2F6FB"><input name="KW" type="text" id="KW" size="36" maxlength="60" /></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">所属类别</td>
        <td align="left" bgcolor="#F2F6FB" ><select class="form" id="select" name="TypLay" >
          <%=Get_rsTree(conn,"SELECT * FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypLayer ",InfType,"Lay") %>
        </select></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">会员帐号</td>
        <td align="left" bgcolor="#F2F6FB"><input name="sUser" type="text" id="sUser" size="36" /></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#F2F6FB">&nbsp;</td>
        <td align="left" bgcolor="#F2F6FB"><input type="submit" name="Button" value="提交" />
          &nbsp;&nbsp;
          <input type="reset" name="Submit2" value="重置" />
          <input name="send" type="hidden" id="send" value="send" />
          <input name="TypID" type="hidden" id="TypID" value="($Search$)" /></td>
      </tr>
    </form>
  </table>
  <div style="line-height:8px;">&nbsp;</div>
  <table width="960" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#FFFFFF">
    <tr>
      <td height="24" align="center" bgcolor="#FFFFFF">&nbsp;<img src="/img/tool/email2.gif" width="16" height="16" align="absmiddle" /> 普通帖子&nbsp;&nbsp;&nbsp;<img src="/img/tool/icon_jian.gif" width="15" height="15" align="absmiddle" /> 推荐帖子&nbsp;&nbsp;&nbsp;<img src="/img/tool/new2.gif" width="30" height="10" align="absmiddle" /> 今日新帖&nbsp;</td>
    </tr>
  </table>
  <div class="line08">&nbsp;</div>
</div>
<!--#include file="_ibot.asp"-->
</body>
</html>
