<!--#include file="../himg/tconfig.asp"-->
<%Call Chk_Perm1(xPara,"")%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>开发者说明文件 -<%=Config_Name%></title>
<LINK href="../himg/hstyle.css" type=text/css rel=stylesheet>
</head>
<body>

<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD vAlign=top align=center bgColor=#ffffff class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 开发者说明文件</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
          <tr>
            <td valign="top" bgcolor="#FFFFFF">　　这是给 网站开发者查看使用的，<span class="FntF00">内容管理者请离开</span>！！！ 
            <span class="pILink">文件说明</span></td>
            <td width="25%" valign="top" bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"><!--#include file="vmenu.asp"--></td>
          </tr>
      </table></TD>
    </TR>
  </TBODY>
</TABLE>

<P class=dTitle><A id=FlagUTF6 name=FlagPicUP></A> 文件说明</P>
<table width="750" border="0" align="center" cellpadding="8" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="left" bgcolor="#FFFFFF"><ul>
        <li>\sadm\admin\sys_1key.asp 1键配置</li>
        <li>\sadm\admin\sys_dbacc.asp 数据库 备份压缩(Access)</li>
        <li>\sadm\admin\sys_config.asp 站点配置管理</li>
        <li>\sadm\admin\type_list.asp 类别设置管理</li>
        <li>\sadm\admin\typs_list.asp 类别设置管理</li>
        </ul>
      <ul>
        <li>\sadm\func1\FuncJS.js 常用js函数</li>
        <li>\sadm\func1\func_obj.asp 检查组件; (Do_Cmd,相关水印,发邮件函数)</li>
        <li>\sadm\func1\func_opt.asp 从DB资料中得到option列表函数</li>
        <li>\sadm\func1\func_rs.asp DB相关函数</li>
        <li>\sadm\func1\func_time.asp 时间ID相关函数</li>
        <li>\sadm\func1\func_vbs.asp 一些vbs常用函数</li>
        <li>\sadm\func1\md5_func.asp MD5加密函数</li>
        <li>\sadm\func1\rsa_fpea.asp rsa加密函数 *** </li>
        <li>\sadm\func1\up_class.asp 无组件上传组件</li>
      </ul>
      <ul>
        <li>\sadm\func2\cch_*.asp 缓存相关文件</li>
        <li>\sadm\func2\func_const.asp 重要配置文件，同 /upfile/sys/config/Config.asp（刷新过来的）一起使用</li>
        <li>\sadm\func2\func_page.asp 分页函数</li>
        <li>\sadm\func2\func_perm.asp 权限判断函数</li>
        <li>\sadm\func2\func_sfile.asp 内容存文件，显示模式等相关函数</li>
      </ul>
      <ul>
        <li>\sadm\system\logs_list.asp 系统记录</li>
        <li>\sadm\system\para_*.asp 参数设置相关系列页</li>
        <li>\sadm\system\system.asp 系统模块管理</li>
        <li>\sadm\system\ upd_para.asp 参数刷新</li>
      </ul>
      <ul>
        <li>\sadm\user\user.asp 系统用户管理</li>
        <li>\sadm\user\user_editpw.asp 用户密码修改</li>
        <li>\sadm\user\user_perm.asp 用户权限分配      </li>
      </ul>
      <ul>
        <li>\inc\home\check.asp Session检查刷新</li>
        <li>\inc\home\convert.js js简繁转化</li>
        <li>\inc\home\func3.asp 前台显示相关函数</li>
        <li>\inc\home\jsadv.js js广告公共函数</li>
        <li>\inc\home\jsBrows.js js浏览器判断</li>
        <li>\inc\home\jsdat1.js js中文农历日期</li>
        <li>\inc\home\jsdat2.js js英文农历日期</li>
        <li>\inc\home\jsdate.asp js简单日期/多语言版</li>
        <li>\inc\home\jsEvent.js  FireFox的event、event.* 属性</li>
        <li>\inc\home\jsInfo.js js公共函数</li>
        <li>\inc\home\jsload.js jsLoad函数</li>
        <li>\inc\home\jsPager.js js分页程序</li>
        <li>\inc\home\jsPlayer.js js程序器</li>
        <li>\inc\home\playflv.swf flv播放器</li>
        <li>\inc\home\playpic.asp 图片切换播放器（asp）</li>
        <li>\inc\home\playpic.swf 图片切换播放器</li>
        <li>\inc\home\RunActive.js Flash Player Detection</li>
      </ul></td>
  </tr>
</table>
<P><A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>

<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD align=left vAlign=top bgColor=#ffffff class="fSiz12"><a href="vhelp.asp"  class="pILink">基本说明</a> <a href="vfolder.asp" class="pILink">目录说明</a> <a href="vfile.asp" class="pILink">文件说明</a> <a href="vdb.asp" class="pILink">数据库说明</a> <a href="vfunc.asp" class="pILink">重要函数说明</a> </TD>
    </TR>
    <TR>
      <TD vAlign=top align=right bgColor=#ffffff class="dCode">更新 Peace[XieYS] 
        2009-04-24 ~ 2009-05-01 &nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
