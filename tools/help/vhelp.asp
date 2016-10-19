<!--#include file="../himg/tconfig.asp"-->
<%

Call Chk_Perm1(xPara,"")

Function readFile(xFile)
  Dim stm,str
  set stm=server.CreateObject("adodb.stream")
  stm.Type=2 
  stm.Mode=3 
  stm.Charset="utf-8"
  stm.Open
  stm.loadfromfile Server.MapPath(xFile)
  str=stm.ReadText
  stm.Close
  set stm=nothing
  readFile = str
End Function

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>开发者说明文件 -<%=Config_Name%></title>
<LINK href="../himg/hstyle.css" type=text/css rel=stylesheet>
<style type="text/css">
.mTab {
	background-color:#CCC;
}
table.mTab td {
	background-color:#FFF;
	padding:1px;
	line-height:150%;
}
table.mTab th {
	background-color:#E0E0E0;
	padding:1px;
	line-height:150%;
}
table.mTab td, table.mTab th {
	text-align:center;
}
.mCellV {
	color:#00F;
}
.mCellX {
	color:#999;
}
.mCellY {
	color:#999;
	font-size:12px;
}
.mCellW, table.mTab td.mCellW {
	background-color:#FFFFCC;
}
.mCLeft, table.mTab td.mCLeft {
	text-align:left;
}


.nTab {
	background-color:#CCC;
}
table.nTab td {
	background-color:#FFF;
	padding:1px;
	font-size:12px;
	line-height:150%;
}

td.mBot {    font-size:12px;	
}
td.mTop {    font-size:12px;	
}

.mBox {
	width:157px;
	height:128px;
	overflow-y:scroll;
	font-size:12px;
	line-height:150%;
	float:left;
	border:1px solid #999;
	margin:0 0 5px 12px;
	padding: 5px;
}
div.mBox ul {
	margin:0px;
	padding:2px 3px;
}
td.mTop, td.mBot {
    font-size:12px;	
}
td.mTop li {
	width:70px;
	height:20px;
	text-align:left;
	letter-spacing:2px;
	display:block;
	float:left;
	margin:0px;
	padding:2px 0px 3px 1px;
	overflow:hidden;
}

td.mCell, td.mCell div {
	text-align:left
}


</style>
</head>
<body>
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 border=0>
  <TBODY>
    <TR>
      <TD vAlign=top align=center bgColor=#ffffff class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 开发者说明文件</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
          <tr>
            <td valign="top" bgcolor="#FFFFFF">　　这是给 网站开发者查看使用的，<span class="FntF00">内容管理者请离开</span>！！！ <br>
              <span class="pILink">基本说明，概述</span> <a href="#FlagWho" class="pILink">&middot;发布说明</a><a href="#FlagReusme" class="pILink">&middot;系统模块概述</a> 
              <a href="#FlagDemo" class="pILink">&middot;参数显示-示例</a> 
            <a href="#FlagSetup" class="pILink">&middot;安装与配置</a> <a href="#FlagTips" class="pILink">&middot;特色应用与扩展</a><a href="#FlagBad" class="pILink">&middot;一起来抱怨</a> <a href="#FlagNotes" class="pILink">&middot;责权申明与鸣谢</a></td>
            <td width="25%" valign="top" bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"><!--#include file="vmenu.asp"--></td>
          </tr>
        </table></TD>
    </TR>
  </TBODY>
</TABLE>
<P class=dTitle><A id=FlagUTF4 name=FlagWho></A> 发布说明：</P>
<P><span class="pDot2">1. </span>本说明是给 网站开发者查看使用的，请确认你身份；如果你只是内容管理者，请关闭这页！<br>
  <span class="pDot2">2. </span>请确认你 的代码程序来源合法。<br>
  <span class="pDot2">3. </span>本说明只供开发或二次开发资料使用，不作技术支持和服务依据，但可以开发探讨。<br>
<span class="pDot2">4. </span>请用于正当用途。<br>
<span class="pDot2">5. </span>今天[2010-08-02]发布 <span class="Fnt00F">和平鸽(asp)2.0</span>！预计2.0所要的功能部分基本完成；剩下一些帮助文档资料待整理，如：目录说明,文件说明,数据库说明,重要函数说明,前台首页整理 等未整理；预计在2.1(正式版)版前完成。较之1.x版，此版本主要特征为：<span class="FntF00">内容存文件</span>，<span class="FntF00">FCK编辑器与上传和eWeb完美结合</span>，<span class="FntF00">使用缓存技术</span>，<span class="FntF00">显示模式由参数控制</span>等。<br>
<span class="pDot2">6.</span>今天是孩子出生100天的日子：《与宝宝赛跑》------《宝宝健康快乐的成长；我努力的工作，积极的生活》之见证。<br>
--- 2010-08-02 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF name=FlagReusme></A> 系统模块概述：(<span class="FntF0F">提示：查看 [<a href="vver2x.asp">版本信息</a>] 中的最新特征</span>) </P>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="mTab">
  <tr>
    <th>组别 </th>
    <th>模块 </th>
    <th>多级类别</th>
    <th>多语言版 </th>
    <th>会员发布 </th>
    <th>扩展模块</th>
    <th> 参数显示 </th>
    <th width="15%">备注</th>
  </tr>
  <tr>
    <td rowspan="3">新闻<br>
      与<br>
      介绍</td>
    <td>综合介绍</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">*注01</td>
  </tr>
  <tr>
    <td>新闻中心</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">含评论</td>
  </tr>
  <tr>
    <td>部门介绍</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">*注02</td>
  </tr>
  <tr>
    <td rowspan="5">图片<br>
      与<br>
      视频 </td>
    <td>产品图片</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">含订购(购物车)</td>
  </tr>
  <tr>
    <td>人物图片</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">含评论</td>
  </tr>
  <tr>
    <td>其它图片</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">含评论</td>
  </tr>
  <tr>
    <td>视频播放</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">含评论</td>
  </tr>
  <tr>
    <td>下载列表</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">仅一个列表</td>
  </tr>
  <tr>
    <td>留言</td>
    <td>访客留言</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">*注03</td>
  </tr>
  <tr>
    <th>组别 </th>
    <th>模块 </th>
    <th>多级类别</th>
    <th>多语言版 </th>
    <th>会员发布 </th>
    <th>扩展模块</th>
    <th> 参数显示 </th>
    <th>备注</th>
  </tr>
  <tr>
    <td rowspan="3">内部<br>
      系统</td>
    <td>内部会员</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">*注04</td>
  </tr>
  <tr>
    <td>内部公文</td>
    <td class="mCellV">V</td>
    <td class="mCellW">X</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">*注05</td>
  </tr>
  <tr>
    <td>内部论坛</td>
    <td class="mCellV">V</td>
    <td class="mCellW">X</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">*注06</td>
  </tr>
  <tr>
    <td rowspan="4">会员<br>
      系统</td>
    <td>会员中心</td>
    <td class="mCellX">-</td>
    <td class="mCellW">X</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">*注07</td>
  </tr>
  <tr>
    <td>公共论坛</td>
    <td class="mCellV">V</td>
    <td class="mCellW">X</td>
    <td class="mCellV">V</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">&nbsp;</td>
  </tr>
  <tr>
    <td>观点辩论</td>
    <td class="mCellV">V</td>
    <td class="mCellW">X</td>
    <td class="mCellV">V</td>
    <td class="mCellX">-</td>
    <td class="mCellX">-</td>
    <td class="fSiz12">论坛的扩充</td>
  </tr>
  <tr>
    <td>商务供求</td>
    <td class="mCellX">-</td>
    <td class="mCellW">X</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="mCellV">V</td>
    <td class="fSiz12">*注08</td>
  </tr>
  <tr>
    <th>组别 </th>
    <th>模块 </th>
    <th>多级类别</th>
    <th>多语言版 </th>
    <th>会员发布 </th>
    <th>扩展模块</th>
    <th> 参数显示 </th>
    <th>备注</th>
  </tr>
  <tr>
    <td rowspan="2">扩展<br>
    其它</td>
    <td class="mCellY">浮动广告</td>
    <td class="mCellY">图片广告</td>
    <td class="mCellY">文字广告</td>
    <td class="mCellY">RSS订阅</td>
    <td class="mCellY">投票调查</td>
    <td class="mCellY">访问统计</td>
    <td><span class="fSiz12">*注09</span></td>
  </tr>
  <tr>
    <td class="mCellY">DB管理器</td>
    <td class="mCellY">&nbsp;</td>
    <td class="mCellY">&nbsp;</td>
    <td class="mCellY">档案查询</td>
    <td class="mCellY">友情连接</td>
    <td class="mCellY">站点导航</td>
    <td><span class="fSiz12">*注10</span></td>
  </tr>
  <tr>
    <td rowspan="3">工具</td>
    <td class="mCellY"><span class="mCellX">木马扫描</span></td>
    <td class="mCellY"><span class="mCellX"> IP监控</span></td>
    <td class="mCellY"><span class="mCellX">超级素数</span></td>
    <td class="mCellY"><span class="mCellX">计数器</span></td>
    <td class="mCellY"><span class="mCellX">万年历</span></td>
    <td class="mCellY"><span class="mCellX">算24</span></td>
    <td class="fSiz12">&nbsp;</td>
  </tr>
  <tr>
    <td class="mCellY"><span class="mCellX">算素数</span></td>
    <td class="mCellY"><span class="mCellX">连接抓取</span></td>
    <td class="mCellY"><span class="mCellX">鸽子年历</span></td>
    <td class="mCellY"><span class="mCellX">颜色选框</span></td>
    <td class="mCellY"><span class="mCellX">效率测试</span></td>
    <td class="mCellY"><span class="mCellX"> 阿江探针</span></td>
    <td class="fSiz12">&nbsp;</td>
  </tr>
  <tr>
    <td class="mCellY"><span class="mCellX">进制转换</span></td>
    <td class="mCellY">安全中心</td>
    <td class="mCellY">ASCII码</td>
    <td class="mCellY">特殊字符</td>
    <td class="mCellY">圆角效果</td>
    <td class="mCellY">素数计算</td>
    <td class="fSiz12">&nbsp;</td>
  </tr>
  
  <tr>
    <td rowspan="2">典型<br>
      特征</td>
    <td class="mCellY">内容存文件</td>
    <td class="mCellY">FCK编辑器</td>
    <td class="mCellY">使用缓存技术</td>
    <td class="mCellY">utf-8编码</td>
    <td class="mCellY">水印开关</td>
    <td class="mCellY">图片认证码</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="mCellY">图片等比缩放</td>
    <td class="mCellY">权限测试</td>
    <td class="mCellY">&nbsp;</td>
    <td class="mCellY">&nbsp;</td>
    <td class="mCellY">js分页程序</td>
    <td class="mCellY">js繁体转化</td>
    <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td colspan="8" align="left" class="mCLeft fSiz12">
      <li>注A1：[类别级数] 可以设置无限制(但建议不超过5)类别级数；前台CSS控制显示，目前CSS可显示3级；如前台显示不支持，则不在次讨论范围。</li>
      <li>注A2：[多语言版] 复制一个目录，只设置一个参数(或少数)，列表 详情等页，就改变了显示语言的框架；</li>
      <li>注A3：[会员发布] 在会员后台加一个连接，会员就可发布本模块信息；</li>
      <li>注A4：[扩展模块] 可扩展无限多(但建议不超过20)个类似模块；</li>
      <li>注A5：[参数显示] 可另扩充最多24个参数；可设置：连接,列表,方式,详情等显示方式；</li>
      <li>注B1：上表：<span class="Fnt00F">[V]</span>表示支持；[X]表示不支持；<span class="Fnt999">[-]</span>无意义；</li>
      <li>注01：同一大栏目中，可设置单页介绍，上下页介绍，问答，列表，图片 等显示方式；</li>
      <li>注02：每一部门(如医院科室)，都可设置如[注01]的小栏目；</li>
      <li>注03：含站务笔记，私人秘密(笔记)，供管理员用；</li>
      <li>注04：同管理员同表；用于内部公文，内部论坛；</li>
      <li><span class="fSiz12">注05：含查阅记录 公文评论</span>；</li>
      <li><span class="fSiz12">注06：可同时配置，公开/内部 两个论坛</span>；</li>
      <li><span class="fSiz12">注07：会员后台 含会员反馈 通知 笔记 </span>等，从留言里面分离出来的；</li>
      <li><span class="fSiz12">注08：含 企业资料，企业介绍，新闻中心，供求交易，招聘信息</span>；可扩展无限多(但建议不超过5)个类似模块；</li>
      <li><span class="fSiz12">注09：[投票调查] 含 投票项目，问卷调查，简易调查</span>(两个)不同模式的 <span class="fSiz12">投票</span>/<span class="fSiz12">调查</span>；</li>
      <li><span class="fSiz12">注10：[档案查询] </span>可从Excel导入资料，实际应用中，需要根据实际对字段等作较多的调整；&nbsp;<A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></li>
    </td>
  </tr>
  
  <tr>
    <td class="mCellV">V</td>
    <td align="left">表示支持</td>
    <td align="left" class="mCLeft">&nbsp;</td>
    <td class="mCellX">-</td>
    <td align="left">无意义</td>
    <td align="left" class="mCLeft">&nbsp;</td>
    <td class="mCellW">X</td>
    <td align="left">不支持</td>
  </tr>
  <tr>
    <td colspan="8" align="left">
    后台默认模块（实际在不断更新中...）
    <table width="740" border="0" align="center" cellpadding="1" cellspacing="1">

        <tr>
          <td valign="top" class="mCell"><%

str = readFile("../../inc/adm_inc/bak_menu.asp")
str = Replace(str,"href=","url") 
str = Replace(str,"../../","../") 
str = Replace(str,"onClick=","") 
str = Replace(str,"<div class='left_end'></div>","")
str = Replace(str,"<li class='left_top_right right'></li>","") 
str = Replace(str,"<div","<div class='mBox'") 

Response.Write str
'str1 = Replace(str1,"<","&lt;")

%></td>
        </tr>
      </table>
      前台默认模块（实际在不断更新中...）
      <table width="740" border="0" align="center" cellpadding="1" cellspacing="1">
        <tr>
          <td valign="top" class="mTop"><%
	str = readFile("../../upfile/sys/config/menu_top.htm")
    str = Replace(str,"href=","url") 
	Response.Write str
	%></td>
        </tr>
        <tr>
          <td align="center" valign="top" class="mBot"> Home
            <%
	str = readFile("../../upfile/sys/config/sys_menu2.htm")
    str = Replace(str,"href=","url") 
	Response.Write str
	%></td>
        </tr>
      </table>
</td>
  </tr>
</table>

<P class=dTitle><A id=FlagUTF2 name=FlagDemo></A> 参数显示-示例：</P>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="1" class="nTab">
  <tr>
    <th width="10%">组别 </th>
    <th width="20%">编码</th>
    <th width="10%">名称</th>
    <th>示例地址</th>
    <th width="20%">说明</th>
  </tr>
  <tr>
    <td rowspan="7" align="center">连接</td>
    <td>(Null)</td>
    <td>无连接</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicR124&TypID=R110020">info.asp?ModID=PicR124&amp;TypID=R110020</a></td>
    <td>无[查看详情]连接</td>
  </tr>
  <tr>
    <td>(ImgNam1)</td>
    <td>连接图片</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicR124&TypID=R110024">info.asp?ModID=PicR124&amp;TypID=R110024</a></td>
    <td align="left">[点图片放大]</td>
  </tr>
  <tr>
    <td>(InfPar2)</td>
    <td>外部连接</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110026">info.asp?ModID=InfA124&amp;TypID=A110026</a></td>
    <td>[连接到本网站外]</td>
  </tr>
  <tr>
    <td>iview.asp?KeyID=(KeyID)</td>
    <td>iview.asp</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicR124">info.asp?ModID=PicR124</a></td>
    <td>默认iview.asp查看详情</td>
  </tr>
  <tr>
    <td>?KeyID=(KeyID)</td>
    <td>本页</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110020">info.asp?ModID=InfA124&amp;TypID=A110020</a></td>
    <td>本页查看详情</td>
  </tr>
  <tr>
    <td>nview.asp?KeyID=(KeyID)</td>
    <td>---</td>
    <td>---</td>
    <td>扩充...</td>
  </tr>
  <tr>
    <td>pview.asp?KeyID=(KeyID)</td>
    <td>---</td>
    <td>---</td>
    <td>扩充...</td>
  </tr>
  <tr>
    <td rowspan="14" align="center">列表</td>
    <td>Cont</td>
    <td>单页</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110012">info.asp?ModID=InfA124&amp;TypID=A110012</a></td>
    <td>用于介绍</td>
  </tr>
  <tr>
    <td>Next</td>
    <td>上下页</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110020">info.asp?ModID=InfA124&amp;TypID=A110020</a></td>
    <td>用于多页介绍</td>
  </tr>
  <tr>
    <td>FAQ</td>
    <td>问答</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110016">info.asp?ModID=InfA124&amp;TypID=A110016</a></td>
    <td>用于FAQ等</td>
  </tr>
  <tr>
    <td>NTab</td>
    <td>列表</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110024">info.asp?ModID=InfA124&amp;TypID=A110024</a></td>
    <td>用于列表新闻/无翻页</td>
  </tr>
  <tr>
    <td>PTab</td>
    <td>图片</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110028">info.asp?ModID=InfA124&amp;TypID=A110028</a></td>
    <td>用于列表图片/无翻页</td>
  </tr>
  <tr>
    <td>PicA</td>
    <td>宽4:3</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicR124&_testTempID=PicA">info.asp?ModID=PicR124&amp;_testTempID=PicA</a></td>
    <td>2x3列行图片</td>
  </tr>
  <tr>
    <td>PicB</td>
    <td>长3:4</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicR124&_testTempID=PicB">info.asp?ModID=PicR124&amp;_testTempID=PicB</a></td>
    <td>3x2列行图片</td>
  </tr>
  <tr>
    <td>PicC</td>
    <td>一行一个</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicR124&_testTempID=PicC">info.asp?ModID=PicR124&amp;_testTempID=PicC</a></td>
    <td>一行一个图片+内容</td>
  </tr>
  <tr>
    <td>News</td>
    <td>新闻</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfN124">info.asp?ModID=InfN124</a></td>
    <td>新闻列表</td>
  </tr>
  <tr>
    <td>Jobs</td>
    <td>职位</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=InfA124&TypID=A110030">info.asp?ModID=InfA124&amp;TypID=A110030</a></td>
    <td>职位列表-可点应聘</td>
  </tr>
  <tr>
    <td>Pics</td>
    <td>产品</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicS124">info.asp?ModID=PicS124</a></td>
    <td>图片列表</td>
  </tr>
  <tr>
    <td>Vdos</td>
    <td>视频</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicV124">info.asp?ModID=PicV124</a></td>
    <td>视频列表</td>
  </tr>
  <tr>
    <td>Down</td>
    <td>Down</td>
    <td><a target="_blank" href="../../page/info.asp?ModID=PicV125">info.asp?ModID=PicV125</a></td>
    <td>下载列表</td>
  </tr>
  <tr>
    <td>---</td>
    <td>---</td>
    <td>---</td>
    <td>---</td>
  </tr>
  <tr>
    <td rowspan="10" align="center">详情</td>
    <td>1News</td>
    <td>经典新闻</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtinf-2010-78-B6WV.5EWMM">iview.asp?KeyID=dtinf-2010-78-B6WV.5EWMM</a></td>
    <td>标题+内容</td>
  </tr>
  <tr>
    <td>1Pics</td>
    <td>经典产品</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-D3AV.01DXN">iview.asp?KeyID=dtpic-2010-6W-D3AV.01DXN</a></td>
    <td>标题+内容+参数+图片</td>
  </tr>
  <tr>
    <td>1Pic3</td>
    <td>&nbsp;</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-FJRR.01DWT">iview.asp?KeyID=dtpic-2010-6W-FJRR.01DWT</a></td>
    <td>标题+内容+3图片+相关图片</td>
  </tr>
  <tr>
    <td>1Vdos</td>
    <td>&nbsp;</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-G8TU.01DHD">iview.asp?KeyID=dtpic-2010-6W-G8TU.01DHD</a></td>
    <td>视频+内容 </td>
  </tr>
  <tr>
    <td>6UD</td>
    <td>&nbsp;</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-D3AV.01DXN&_TestTempID=6UD">iview.asp?KeyID=(KeyID)&amp;_TestTempID=6UD</a></td>
    <td>图片+内容(上下)</td>
  </tr>
  <tr>
    <td>6Left</td>
    <td>&nbsp;</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-D3AV.01DXN&_TestTempID=6Left">iview.asp?KeyID=(KeyID)&amp;_TestTempID=6Left</a></td>
    <td>图片(左)+内容</td>
  </tr>
  <tr>
    <td>6Right</td>
    <td>&nbsp;</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-D3AV.01DXN&_TestTempID=6Right">iview.asp?KeyID=(KeyID)&amp;_TestTempID=6Right</a></td>
    <td>图片(右)+内容</td>
  </tr>
  <tr>
    <td>2Pics</td>
    <td>产品无订购</td>
    <td><a target="_blank" href="../../page/iview.asp?KeyID=dtpic-2010-6W-D3AV.01DXN&_TestTempID=2Pics">iview.asp?KeyID=(KeyID)&amp;_TestTempID=2Pics</a></td>
    <td>标题+内容+参数+图片</td>
  </tr>
  <tr>
    <td>ImgRelat</td>
    <td>相关图片</td>
    <td><a href="../../page/iview.asp?KeyID=dtpic-2010-6W-FKCA.01DSV" target="_blank">iview.asp?KeyID=dtpic-2010-6W-FKCA.01DSV</a></td>
    <td>显示相关图片</td>
  </tr>
  <tr>
    <td>---</td>
    <td>---</td>
    <td>---</td>
    <td>---</td>
  </tr>
  <tr>
    <td colspan="5" align="left" class="mCLeft fSiz12">注意：&nbsp;<A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></td>
  </tr>
</table>
<P class=dTitle><A id=FlagEditor name=FlagSetup></A>安装与配置</P>
<P><span class="pDot2">初始安装</span>(setup) 见 /setup/readem.txt文件；<br>
  <span class="pDot2">重新安装</span>(setup) 见 /setup/readem.txt文件 或 /sadm/setup/readem.txt文件；<br>
</P>
<P class=dTitle><A id=FlagUTF3 name=FlagTips></A> 特色应用与扩展开发</P>
<P><span class="pDot2">A.</span>添加基本模块：<br>
  <span class="pDot3">a.1.</span> 系统与设置 &gt;&gt; 模块 &gt;&gt; 图片与视频 &gt;&gt; 增加一个模块 “PicT168-测试信息”... <br>
  建议，以图片为主的，用Pic开头，防在[图片与视频]组内；以新闻文字为主的，用Inf开头，防在[新闻与介绍]组内；<br>
  <span class="pDot3">a.2.</span> 系统与设置 &gt;&gt; 配置 &gt;&gt; 初始化菜单 刷新 ... <br>
  <span class="pDot3">a.3.</span> 刷新浏览器 &gt;&gt; 图片与视频 &gt;&gt; 可看到多了个 “测试信息” , 此时，可以设置类别并刷新；<br>
  <span class="pDot3">a.4.</span> 添加信息；前台查看：(path)page/info.asp?ModID=PicT168；</P>
<p>好了，新增加的模块，默认是一个新闻列表，详情显示新闻详情！<br>
  现在我们来改造这个模块，以下设置，入口都在：系统与设置 &gt;&gt; 参数 &gt;&gt; 资料模版 &gt;&gt; </p>
<p><span class="pDot2">B.</span>模块高级设置：<br>
  <span class="pDot3">b.1</span> 设置 图片列表显示方式 --- tmsPara.asp模版参数 (假如要显示为2x3的图片列表)<br>
  找到行：tmsListTab([N]) = &quot;PicA:[...]&quot; '宽4:3 --- 2x3列行 <br>
  在最后的双引号前加“PicT168;”<br>
  <span class="pDot3">b.2</span> 设置 详情显示方式 --- tmsPara.asp模版参数 (假如要显示3个图片)<br>
  找到行：tmsShowTab([N]) = &quot;1Pic3:[...]&quot; '标题+内容+3图片+相关图片<br>
  在最后的双引号前加“PicT168;”<br>
  <span class="pDot3">b.3</span> 设置 设置图片个数 --- tmsPara.asp模版参数 (假如要上传3个图片)<br>
  找到行：' tmsPics.asp 图片个数/相关图片 <br>
  加两行：<br>
  Dim PicT168Count <br>
  PicT168ImgCount = 3 <br>
  <span class="pDot3">b.4</span> 设置 可添加相关图片 --- tmsPara.asp模版参数<br>
  找到行：' tmsPics.asp 图片个数/相关图片 <br>
  加两行：<br>
  Dim PicT168Relat <br>
PicT168ImgRelat = true <br>
<span class="pDot3">b.5</span> 设置 可选择次类别 --- tmsTyp2.asp次类别参数 <br>
  加4行：<br>
  Dim PicT168Typ2Lable,PicT168Typ2Code,PicT168Typ2Name<br>
  PicT168Typ2Lable = &quot;拍摄模式&quot;<br>
  PicT168Typ2Code = &quot;ModTyp212;ModTyp216;ModTyp220&quot;<br>
PicT168Typ2Name = &quot;人物模式;室外模式;自动&quot;<br>
<span class="FntF00">注意：前台查看地址请加参数：&amp;DepID=[次类别ID]如：<br>
(path)page/info.asp?ModID=PicT168&amp;DepID=ModTyp212 </span><br>
<span class="pDot3">b.6</span> 设置默认资料 模版：--- 增加参数 tmpPicT168<br>
  填写好内容，此内容在你添加资料时，就自动填充<br>
  <br>
<span class="pDot2">C.</span>刷新，添加资料，在前后台看效果吧！！！<A href="#"><SPAN class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagUTF5 name=FlagBad></A> 一起来抱怨</P>
<P><span class="pDot2">1. </span>基于Access开发：如果总资料达到500 K(50万记录)，单个表资料达到50 K(5万记录)，恐怕要用Sql数据库了！<br>
  <span class="pDot2">2. </span>语言包：如果仅一个简体中文版，语言包真的多余！<br>
  <span class="pDot2">3. </span>utf-8：如果仅一个简体中文版，无Ajax,Flash+XML技术，用utf-8编码，没有任何优势！<A href="#"><SPAN class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF5 name=FlagNotes></A> 责权申明 与鸣谢</P>
<P><span class="pDot2">1. </span>责权申明：本系统版权归作者所在的公司---<span class="FntF00">东莞网</span>(<a href="http://www.dg.gd.cn/">dg.gd.cn</a>)所有，作者<span class="FntF00">Peace</span>(XieYS)保留修改完善再发布权利；<br>
  <span class="pDot2">2. </span>发布申明：任何从正当途径获得本系统者，可自由传播，供免费使用或二次开发本系统；上述责权人 不作技术支持和服务依据，欢迎技术交流与讨论：Peace(XieYS), QQ:80893510, xpigeion@163.com。<span class="FntF0F">请用于正当用途</span>！本系统开发者，支持正版，尊重劳动成果 ---- 请您也 充分利用网络上的免费的优秀的代码，购买需要的正版程序(或软件)，尊重并感谢每一位无私奉贤的人！<br>
  <span class="pDot2">3. </span>外部程序/模块：<br>
  本系统使用的一些
涉及的一些模块和代码，参考了网上许多优秀文章和程序，列举主要的如下：<br>
  <span class="pDot3">维基百科：</span>简繁体转换字库从这里提取：http://zh.wikipedia.org/w/index.php?title=Wikipedia:Unihan%E7%B9%81%E7%AE%80%E4%BD%93%E5%AF%B9%E7%85%A7%E8%A1%A8&amp;variant=zh-cn<br>
  <span class="pDot3">eWebEditor编辑器：</span>本系统所用编辑器核心代码来源于eWebEditor编辑器，仅用于学习和交流；如用于商业用途，请到官方下载正版：http://www.ewebeditor.net/download.asp，或用它的免费版，或找其他编辑器如FCK编辑器等，（从1.1版开始，开始逐步废除eWebEditor，改用FCKEditor）。<br>
  <span class="pDot3">FCKEditor编辑器：</span>(较新版本使用此编辑器) 本系统所用编辑器核心代码来源于FCKEditor编辑器，此编辑器免费，开源，兼容性很好！[2009-07-07]。<br>
  <span class="pDot3">其它代码与程序：</span>这里不一一说明。<br>
  <span class="pDot2">4. </span>鸣谢：<br>
  <span class="pDot3">感谢</span>互连网以及在互连网上的许许多多奉献者；<br>
  <span class="pDot3">感谢</span>维基百科,eWebEditor编辑器，FCKEditor编辑器等自愿者和相关团队；<br>
  <span class="pDot3">感谢</span>我所在的公司：东莞网络（www.dg.gd.cn）以及很多同事；<br>
  <span class="pDot3">感谢</span>你关注并合法使用本程序！<A href="#"><SPAN class=dCode>[TOP]</SPAN></A></P>
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD align=left vAlign=top bgColor=#ffffff class="fSiz12"><a href="vhelp.asp"  class="pILink">基本说明</a> <a href="vfolder.asp" class="pILink">目录说明</a> <a href="vfile.asp" class="pILink">文件说明</a> <a href="vdb.asp" class="pILink">数据库说明</a> <a href="vfunc.asp" class="pILink">重要函数说明</a></TD>
    </TR>
    <TR>
      <TD vAlign=top align=right bgColor=#ffffff class="dCode">更新 Peace[XieYS] 
        2010-08-02 &nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
