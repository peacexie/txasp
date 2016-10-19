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
      <TD vAlign=top align=center bgColor='#ffffff' class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 开发者说明文件</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
          <tr>
            <td valign="top" bgcolor="#FFFFFF">　　这是给 网站开发者查看使用的，<span class="FntF00">内容管理者请离开</span>！！！ 
              <span class="pILink">目录说明</span>
            </td>
            <td width="25%" valign="top" bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"><!--#include file="vmenu.asp"--></td>
          </tr>
      </table></TD>
    </TR>
  </TBODY>
</TABLE>

<P class=dTitle><A id=FlagUTF name=FlagPicUP></A> 目录说明</P>
<table width="750" border="0" align="center" cellpadding="8" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="left" bgcolor="#FFFFFF"><ul>
        <li>bbs
        论坛前台文件</li>
        <li>doc 内部公文</li>
        <li>ext 扩展代码</li>
        <li>img 公共图片
          <ul>
            <li>face/file 表情，文件类型图标</li>
            <li>Head/inum 头像，数字</li>
            <li>logo logo</li>
            <li>msn/qq MSN， QQ表情</li>
            <li>rnd_nid/tool 圆角图片， 小工具图片</li>
            <li>tree/vote 树型目录， 投票</li>
            <li>xtest (测试文件)</li>
          </ul>
        </li>
        
        <li>inc 公共文件
          <ul>
            <li>adm_img/adm_inc 后台管理图片， 后台管理框架文件</li>
            <li>home 前台显示公共文件</li>
            <li>mem_img/mem_inc 会员管理图片，会员管理框架文件</li>
            <li>xtemp 临时文件(暂时不用)</li>
          </ul>
        </li>
        <li>member 会员前台
          <ul>
            <li>admin 会员管理</li>
            <li>ecard 通用查询</li>
            <li>info 会员资料</li>
          </ul>
        </li>
        <li>page 前台主要显示文件</li>
        <li>pfile 前台显示要调用的文件
（inc,图片，模版）          
  <ul>
            <li>pimg/pinc 前台图片，包含文件</li>
    <li>pub/temp 前台公共文件，模版</li>
  </ul>
        </li>
        <li>sadm
          后台管理主框架
          <ul>
            <li>admin/edfck 站点配置 初始化， 编辑器</li>
            <li>func1/func2 系统重要公用函数1， 系统重要公用函数2</li>
            <li>pcode/setup 认证码，安装备份</li>
            <li>system/user 系统记录 模块 参数，个人密码 系统管理员 权限</li>
          </ul>
        </li>
        <li>setup 安装文件</li>
        <li>smod 主信息管理后台
          <ul>
            <li>adupd/file 刷新杂项目录,文件上传</li>
            <li>gbook/info 留言，新闻/图片主资料管理</li>
            <li>link/type,vote 连接/广告，类别设置，投票</li>
          </ul>
        </li>
        <li>tools 工具，扩展
          <ul>
            <li>_admin/base 工具扩展公共文件，工具扩展主要文件</li>
            <li>help/himg 帮助，帮助图片</li>
            <li>out/peace 部连接抓取，peace收藏代码</li>
          </ul>
        </li>
        <li>trade 商务供求前台</li>
        <li>upfile
          刷新，附件等
          <ul>
            <li>#dbf# 数据库文件等</li>
            <li>defdt/defdt 留言文件目录, 杂项图片上传目录</li>
            <li>dtbbs/dtbus/dtdoc 论坛，供求，公文上传目录</li>
            <li>dtinf/dtpic 新闻，图片上传目录</li>
            <li>mufile/fyftp 自定义，FTP上传目录</li>
            <li>sys
              系统刷新等            
              <ul>
                <li>xadv/app 广告, 会员登陆</li>
                <li>config 主要配置，关键文件</li>
                <li>doc/para 公文， 参数</li>
              </ul>
            </li>
          </ul>
        </li>
        <li>vote 投票调查前台</li>
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
