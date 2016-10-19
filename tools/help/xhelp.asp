<!--#include file="../himg/tconfig.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>帮助中心 -<%=Config_Name%></title>
<LINK href="../himg/hstyle.css" type=text/css rel=stylesheet>
</head>
<body>
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD vAlign=top align=center bgColor=#ffffff class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 后台管理帮助文件</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff>
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
          <tr>
            <td valign="top" bgcolor="#FFFFFF">　　
              能点击来到这页，<span class="FntF0F">非常感谢您操作的细心与严谨</span>！请仔细阅读本说明，如果需要，<span class="FntF0F">请对照网站前后台进行测试阅读</span>。<br>
              -=&gt; <a href="../../smod/file/edt_demo.asp" target="_blank">编辑器基本属性使用说明 (To网站资料管理员)</a><br>
              -=&gt; <a href="xserver.asp" target="_self">服务器常用安全配置 和 站点常用安全配置 (To服务器管理员[网管])</a></td>
            <td width="30%" valign="top" nowrap bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"> &middot; <a href="#FlagFirst">目的和意义</a><br>
              &middot; <a href="#FlagEditor">内容管理与编辑器</a><span class="FntF00">***</span><br>
              &middot; <a href="#FlagComm">常用资料管理与设置</a><span class="FntF00">***</span><br>
              &middot; <a href="#FlagVDO">视频(和大文件管理)</a><br>
              &middot; <a href="#FlagBase">（客户端）基本环境</a><br>
              &middot; <a href="#FlagSafe">安全与超时</a><br>
              &middot; <a href="#FlagPicUP">图片文件上传</a><br>
              &middot; <a href="#FlagTopsN">技巧与提示</a><br>
              &middot; <a href="#FlagVers">英文或其它语言版</a><br>
            </td>
          </tr>
        </table>
      </TD>
    </TR>
  </TBODY>
</TABLE>
<P class=dTitle><A id=FlagUTF4 name=FlagFirst></A> 目的和意义：</P>
<P><span class="pDot2">1. </span>网站做后台模块的意义就是：离开网站制作人员，<span class="FntF00">让用户(您)自主维护您的网站资料</span>！所以，请认真编辑处理您网站的文字图片资料；文字编辑，简单图片处理等，请自己动手 ------ 如果不会，请上网查资料或去培训部学习；网站制作人员不可能也无这个义务培训您这些！这里<span class="FntF00">整理一些网站维护中常见的小技巧和小问题，敬请留意</span> ------ 如果你不看这些，请慢慢摸索也可！<br>
  <span class="pDot2">2. </span>本系统开发者和所在的公司，都没有开网页设计相关的培训课程；任何类似如下问题(包含但不限于)，都谢绝回答（如果是技术探讨，另当别论）：<br>
  <span class="pDot3">1. </span>“浏览器是什么”，“Outlook是什么？”，“域名是什么？”，“网址栏在哪里？”，“邮件签名怎样设置”等问题；--- 请到电脑培训部去学习；<br>
  <span class="pDot3">2. </span>“我不能上网”，“我电脑中毒了”等问题；--- 不能上网可找网络提供商，中毒了请专门的杀毒公司；<br>
  <span class="pDot3">3. </span>“怎样用Photoshop”，“怎样用Dreamweaver改网页”等问题；--- 请到电脑培训部学习去培训；<br>
  <span class="pDot3">4. </span>“怎样用FTP”，等问题；--- 可在网上找找资料，或不用了；提示你如果用了FTP请自行负责后果；<br>
  以上，并不是怕你知道了抢我们的饭碗，我倒是提醒你另外请人管理你网站了；相反，我倒希望你(网站管理者)多知道一些上述常识或知识，用好网络，用好你的网站，管理好你网站的资料；<br>
  <span class="pDot2">3. </span>如果我说话态度不好，深表歉意！如果您不接受自己动手进行文字编辑图片处理等，那请做静态网站！<br>
  <span class="pDot2">4. </span>敬请您认真地<span class="FntF00">用心地</span>去维护您的网站资料！谢谢！当您能自如管理你网站图片文字内容的时候，相信您也有无比的自豪感和成就感？！ <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>

<P class=dTitle><A id=FlagEditor name=FlagEditor></A>内容管理与编辑器</P>
<P><span class="pDot2">1. </span>本段对 后台编辑框 的一些关键操作 做一些简单提示

  <a href="../../smod/file/edt_demo.asp" target="_blank"><img src="../himg/editor.jpg" width="580" height="60" border="0" align="right"></a>

注意：1.按钮实际位置 以实物为准；2.不可能每个按钮都做说明，请理解；3.更多的操作，请自己摸索！<br>
本段 包含：[<span class="Fnt00F">两种换行</span>] [<span class="Fnt00F">从Word和网页中复制资料</span>] [<span class="Fnt00F">建议用默认字体</span>] [<span class="Fnt00F">插入图片和图片属性</span>] [<span class="Fnt00F">各种链接(文字链接/QQ聊天/图片链接/下载等的链接</span>] <span class="Fnt00F">[添加视频</span>] [<span class="Fnt00F">添加表格</span>]。详细的 <a href="../../smod/file/edt_demo.asp" target="_blank">编辑器基本属性使用说明</a>(可能要登陆)。<A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A>
</P>
<P class=dTitle><A id=FlagComm name=FlagComm></A>常用资料管理与设置：</P>
<P>这是通用设置说明，我们保留不通知客户的情况下，作优化性的小修改或小调整的权力！以下仅作提示：<span class="FntF00"><br>
  A. 具体栏目对应关系，以前台栏目显示为准； <br>
  B. 具体的菜单项名称和一些文字描述等，以看到的网站实际内容为准！</span><br>
  <br>
  <span class="pDot2">1. </span>设置 栏目(设置类别)<span class="dTitle"><A id=FlagComm2 name=FlagComm_1></A></span><br>
  <span class="pDot3">[设置说明]</span> 一般情况下，各栏目和类别已经设置好；如第一次使用，请在后台先检查各栏目和类别,有必要的话重新设置一次，并刷新；注意，有些栏目设置，由于前台显示限制，后台设置不起作用，此功能一般用于设置子类别；<br>
  <span class="pDot3">[入口/步骤]</span> (登陆)后台 &gt;&gt; 在左侧相关菜单，点“类别”或类别管理，进入类别管理页；设置好相关类别；设置好后点“刷新”使设置生效。<a href="../himg/type01.jpg" target="_blank">点此见：参考图</a>。<br>
  <span class="pDot2">2. </span>设置 版权信息：<span class="dTitle"><A id=FlagComm3 name=FlagComm_2></A></span><br>
  <span class="pDot3">[设置说明]</span> 一般的，下方的 Copyright 版权信息，如果是文字形式，则提供后台修改。<br>
  <span class="pDot3">[入口/步骤]</span> (登陆)后台 &gt;&gt; 系统与设置 &gt;&gt; 参数 &gt;&gt; 备注 &gt;&gt; rCopy.htm(版权信息) &gt;&gt; 保存。<br>
  <span class="pDot3">[扩展]</span> 如有英文版，类似的 修改 rCEng.htm参数,设置好保存后 刷新生效。<a href="../himg/copy01.jpg" target="_blank">点此见：参考图</a>。<br>
  <span class="pDot2">3. </span>设置 友情连接：<span class="dTitle"><A id=FlagComm4 name=FlagComm_3></A></span><br>
  <span class="pDot3">[设置说明]</span> 如果是文字形式的 友情连接，没有特殊要求，则提供后台修改。<br>
  <span class="pDot3">[入口/步骤]</span> (登陆)后台 &gt;&gt; 刷新与杂项（或相关菜单） &gt;&gt; 友情连接 增加 分类 &gt;&gt; <br>
  设置好保存后，请点一下 友情连接 列表页 右上角的 “刷新”。<a href="../himg/link01.jpg" target="_blank">点此见：参考图</a>。<br>
  <span class="pDot2">4. </span>说明 刷新与杂项/系统与设置(有的根据需要，合并了这两项)：<span class="dTitle"><A id=FlagComm5 name=FlagComm_4></A></span><br>
  <span class="pDot3">1. 系统记录</span> 建议经常删除很久以前的记录数据，但建议你保留最近几个月的记录；<br>
  <span class="pDot3">2. 数据库管理</span> 建议一周或一月备份/压缩一次,保留最近3~5个备份文件，如恢复数据库，请谨慎！；<br>
  <span class="pDot3">3. 个人密码,管理员列表</span> 请根据需要，设置适当的管理员；<br>
  <span class="pDot3">4. 刷新与杂项</span> 系统与设的其它设置项目，一般都设置好；如果有些项目用不上的，请不用管它 <br>
  --- 他们几乎不占用空间资源！如果确认   用不上，可以删除，如重新设置或删除项目，请谨慎操作 或 联系开发人员！<br>
  <span class="pDot3">5. 刷新原理与说明</span> <span class="FntF00">刷新是为了优化网站</span>，把一些很少更新的东西，但又在很多地方都要用到的资料，刷新成静态或部分静态，以达到优化效果。---[技巧]你可以修改好一组项目后，一起执行刷新；如你可以把所有类别(栏目)都设置好后再刷新类别；如你可以把所有参数都设置好后再刷新参数；如你可以把所有友情连接都设置好后再刷新友情连接……<br>
  <span class="pDot2">5. </span>通用信息属性设置：<span class="dTitle"><A id=FlagComm6 name=FlagComm_5></A></span> <br>
  一般的，新闻，图片信息，有如下通用属性，设置和说明如下：<br>
  <span class="pDot3">删除.所选：</span>删除列表中选中的项目；（以下都为针对选中项操作）<br>
  <span class="pDot3">设置_显示：</span>一般不起作用，作为功能扩展备用；<br>
  <span class="pDot3">设置_推荐：</span>一般不起作用；特殊情况下，如首页需要从很多信息（图片）中选取部分在首页显示，可能要用此属性，具体逻辑设置情况，一般另有文件特别说明；<br>
  <span class="pDot3">设置_顺序：</span>一般不起作用；<br>
  一般信息默认顺序全部为888，前台一般按时间[后-先]顺序，即最新（最近）发布的信息最前显示，如果要别的显示逻辑，请事先说明；<br>
  特殊情况下，如需要把少数信息提前（靠前）显示，可能要用到此“设置顺序”，注意，为减少以后维护负担，此项设置不宜使用过多，特别设置的信息，过期后，请把它属性设置还原，否则，以后发布的最新资料，也许不会靠前显示；<br>
  <span class="pDot3">常用杂项快捷设置：</span> 刷新与杂项 &gt;&gt; 常用杂项 --- 快捷设置  &gt;&gt; <br>
  参数 中一般都可找到，这里只是把常用的参数提取出来，方便管理；<br>
  <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF name=FlagVDO></A>视频(和大文件管理)：</P>
<P><span class="FntF00">如果有视频管理模块（栏目）</span>，请注意本项目；注意：后台一般不上传视频，只作视频相关信息管理。<br>
  <br>
  <span class="pDot2">1. </span>视频文件一般占用空间非常大，建议用FTP上传，建议视频大小最好在15M以下，播放片长10分钟以内，FLV格式或WMV格式；如只是测试可以用更小的视频文件；<br>
  <span class="pDot2">2. </span>视频播放要求的服务器环境和网络环境都很高，如播放很卡等不在本程序系统范围内讨论，视频文件压缩处理等操作请网站资料管理相关人员自行解决。<br>
  <span class="pDot2">3. </span>视频文件一般占用空间非常大,一般企业租用空间都放不了几个大视频，这里只作提示，相关问题请联系空间供应商<br>
  <span class="pDot2">4. </span>IIS设置，这是给网站配置人员看的；为了让网站支持flv媒体，请在IIS中作 如下设置：IIS 》》站点》》属性》》http头》》mime类型》》增加》》<br>
  .flv ; application/x-shockwave-flash，其它根据需要设置(<a href="xserver.asp" target="_blank">常见文件MIME类型.txt</a>)。(租用虚拟主机这，此部分有空间供应商设置) <br>
  <span class="pDot2">5. </span> FTP上传相关资料: 本节给不会使用FTP的网站内容管理员参考；本节仅作提示，不实际培训！<br>
  <span class="pDot3">1. </span>假如你有管理网站的FTP权限，且你懂得用FTP管理网站，请看看FTP管理大文件的建议：<a href="../../upfile/readme.txt" target="_blank">目录文件规划建议</a>。（警告：如果您不懂得用FTP管理网站，建议不要管它了，否则后果自负！如果空间提供商不给你FTP权限，不在本程序讨论范围内！） <br>
  <span class="pDot3">2. </span>FTP下载相关连接: 用 <a 
href="http://www.baidu.com/" target="_blank">baidu.com</a>，<a 
href="http://www.google.cn/" target="_blank">google.cn</a>，<a 
href="http://cn.bing.com/" target="_blank">cn.bing.com</a> 等搜索 CuteFTP,  LeapFTP, FlashFXP, TurboFTP 下载等关键字，下载并安装FTP客户端软件，并按以下2方法配置使用FTP上传大文件；<br>
  <span class="pDot3">3. </span>FTP使用相关连接: <a 
href="http://www.baidu.com/s?wd=FTP%CA%B9%D3%C3" target="_blank">FTP使用@baidu.com</a>, <a 
href="http://www.google.cn/search?hl=zh-CN&source=hp&q=ftp%E4%BD%BF%E7%94%A8%E6%96%B9%E6%B3%95&aq=0&oq=FTP%E4%BD%BF%E7%94%A8" target="_blank">FTP使用@google.cn</a>, <a 
href="http://cn.bing.com/search?q=FTP%E4%BD%BF%E7%94%A8&go=&form=QBLH&filt=all" target="_blank">FTP使用@cn.bing.com</a>；<A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF2 name=FlagBase></A> （客户端）基本环境：</P>
<P><span class="pDot2">1. </span>后台管理推荐使用IE7,IE8浏览器，如用其它浏览器,相关设置问题,请自行解决！<br>
  <span class="pDot2">2. </span>正常浏览本系统，请满足以下基本环境：开通Javascript；启用iFrame；在后台，使用1024x768以上效果最佳；<br>
  本系统不需要安装任何第三方插件；<br>
  <span class="pDot2">3. </span>以下环境测试代码可以参考： <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<table width="480" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF">测试Javascript结果：<span id="PeaceJSTest"><font color="#FF00FF">您没有启用Javascript，或设置不正确</font></span>；<br>
      测试iFrame：<a 
									  href="http://www.dgchr.com/about/iframe.asp" target="_blank"><font color="#FF00FF">点击测试</font></a>。</td>
  </tr>
</table>
<P class=dTitle><A id=FlagUTF3 name=FlagSafe></A> 安全与超时：</P>
<P><span class="pDot2">1. </span>为了您的安全，请仔细阅读本条款：<br>
  1) 如果您是本系统管理员或会员，请保管好您的帐号密码，因用户自身原因，泄密密码而造成的损失由用户负责；<br>
  2) 使用后一定要登出系统；<br>
  <span class="pDot2">2. </span>如果某网页20分钟（左右）没有任何操作，可能出现“超时”这种情况，这也是一种安全机制。<br>
  如果登陆了系统，请先退出系统，重新登陆；如果是注册时出现，请刷新后在进行。<br>
  <span class="pDot2">3. </span>强烈建议：<span class="FntF0F">任何管理员帐号密码等，不要设置为常用的，别人容易猜测到的帐号密码</span>；比如，下的帐号密码串：administrator, admin, admin888, adm888, master, manage, manager,  123456, asdf, guest等，很容易被猜测到，建议不要用。 <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF4 name=FlagPicUP></A> 图片文件与上传：</P>
<P><span class="pDot2">1. </span>文件大小(占用空间,单位是KB或B)：尽管程序可以设置上传较大的图片文件，但这还要看服务器的相关设置；一般情况，上传200K以下的图片文件是比较保险的，只要正确处理，一般200K的图片文件能基本满足网站需要；<br>
  文件尺寸(宽和高,单位是象素[px])：一般以前台显示空间为基准，进行调整；<br>
  <span class="pDot2">2. </span>文件名目录名：放在本地的待上传文件，其文件名（含目录名）建议全用英文半角的字母，数字或下划线；<span class="FntF00">不能用</span><span class="FntF0F">空格</span><span class="Fnt00F">引号</span><span class="FntF00">等特殊字符</span>；也建议不要用中文；<br>
  <span class="pDot2">3. </span>文件格式：图片建议使用jpg,gif格式；如不能正常显示(如显示为一个红叉)，可能是图片格式有问题；如可以用Photoshop“存储为Web格式”,这不是唯一方法，这里只作提示，不具体培训；<br>
  <span class="pDot2">4. </span>所有上传文件类型，建议用常见的图片(jpg,gif)，Flash(swf,flv)，Office(.doc.xls.ppt.pps.wps.et.dpt等)，视频(flv,wmv等)，如上传其它类型，可先压缩后再上传；任何包含有可疑恶意代码的文件，禁止程序上传，如需要，可先压缩后再上传；<br>
  <span class="pDot2">5. </span>图片处理小工具：本开发人员推荐一个工具，可批量处理缩放图片；绿色(不用安装)，仅一个文件，仅40K……本人Win2003下正常运行，当然，这也不是唯一方法，仅做提示；如果你不能使用，完全可以使用别的软件如PS,FW等。有必要请下载：<a href="../../upfile/myftp/down/ZoomImg.rar">(下载批量图片处理工具)</a>。 <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF5 name=FlagTopsN></A> 技巧与提示：</P>
<P><span class="pDot2">1. </span>用 &quot;站务笔记&quot; 协助管理：<br>
  提示，<span class="FntF00">一般</span>系统都附带“站务笔记”功能，你可以把网站管理中的使用心得或注意事项写进去。如果不使用，或没有此项，完全不影响网站其它功能。此项目一般放在：后台 &gt;&gt; 刷新与杂项 &gt;&gt; 留言与笔记 &gt;&gt; 站务笔记-增加笔记 或 后台 &gt;&gt; 刷新与杂项 &gt;&gt; 站务笔记-增加笔记 中。<br>
  <span class="FntF00">注意</span>：=-- 敏感资料，保密资料（如帐号密码等），请不要写进站务笔记；如果换管理员，可吧站务笔记的记录一并交接给下一个管理员。 --= <br>
  <span class="pDot2">2. </span>刷新与杂项，系统与设置菜单：<br>
  <span class="pDot3">1. </span>系统记录 : 建议经常删除很久以前的记录数据，但建议你保留最近一个月左右的记录；<br>
  <span class="pDot3">2. </span>数据库管理 : 建议一周或一月备份/压缩一次,保留最近3~5个备份文件，如恢复数据库，请谨慎！；<br>
  <span class="pDot3">3. </span>个人密码,管理员列表 : 请根据需要，设置适当的管理员；<br>
  <span class="pDot3">4. </span> 刷新与杂项，系统与设的其它设置项目，一般都设置好；如果有些项目用不上的，请不用管它 --- 他们几乎不占用空间可资源！如果确认用不上，可以删除，如重新设置或删除项目，请谨慎操作 或 联系开发人员！<br>
  <span class="pDot3">5. </span> 在参数和系统设置中，一些设置项目，可能用不上，可能是预留一些扩展，<span class="FntF00">不必刨根问底</span>；可以随意添加设置项目，但要启用或实现这个设置，可能就是个大工程或一个不能实现的工作；比如：您可添加如下开关参数(SwhAutOpen---天黑自动开灯---Y.是;  SwhAutSend---自动发短信---Y.是)，但这样的功能已经超出本系统讨论范围。<br>
  <span class="pDot2"><span class="pDot">3</span>. </span>每个系统，每个网站根据需要或版本不同，术语与操作可能有稍微不同，以上说明仅作参考。<A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF6 name=FlagVers></A> 英文或其它多语言版：</P>
<P>本程序扩展多个语言版本是很容易的，如果你有多语言版，且象一般产品一样，多语言版是同步的资料，那本程序尽最大努力节约您的管理时间成本，提高你的效率和降低你的运营费用，提示如下：<br>
  <span class="pDot3">0. </span>信息编号：多语言版需要资料同步 的情况下，信息编号要求不能为空，不能重复，如果无资料，请用默认]<br>
  <span class="pDot3">1. </span>类别管理：如果是 多语言版，设置好中文类别后，再同步设置英文版本的类别；<br>
  <span class="pDot3">2. </span>资料管理：多语言版，产品等维护资料时，可以这样，先维护（增加，修改图片）好中文版，再进入其它语言版本，资料不用重新增加；果某一个版本，如英文，或日文等没有资料，则<span class="FntF00">执行列表中的同步命令</span>，仅修改文字内容即可，这样节约您的时间管理成本。  尽管如此，但这里事先说明两点：<br>
  <span class="pDot2">一是：</span>如果确实需要多个语言版本，请先策划好； 事先就策划好做两个或多个版本 比 做好一个版本后，再添加其它语言版本 要 简单些；花费也许也要少些！比如（但不等于实际）：以做一个中文版的费用为 [1个单位]，事先策划做两个版本的费用则为 [1.5个单位]；而先做一个版本后再加一个版本的费用可能是1+1个单位或更多！<br>
  <span class="pDot2">二是：</span>是否有必要做多个语言版本，提醒以下个人意见！<br>
  <span class="pDot3">1. </span>做英文或其它语言版本的必要性：<br>
  如果没有能力维护英文版或其它语言版的能力；请不要浪费资经，不要浪费资源，不要浪费感情！请不要做英文或其它语言版本！<br>
  <span class="pDot3">2. </span> 对中文版以外的版本，请事先<span class="FntF0F">你们自己提供</span>公司名称公司简介或专业术语等翻译资料；对公司名称和专业术语等，就是请专门的翻译公司(人员)也难以办到，对英文版本 的 网络专业术语，可以不必提供，本程序作者给你搞定，你可以节约这笔开支；中文，英文以外的其它语言版，网络专业术语也要提供，或另外找专业人员！<br>
  <span class="pDot3">3. </span> <span class="FntF0F">提供必要的缩写或简写翻译资料</span>：主要是根据前台排版的需要，需提供对应的必要的缩写或简写翻译资料，比如一些菜单等处，排版时不能显示完整的文字，则此时需要提供对应的缩写或简写翻译资料，请根据需要提供。<br>
  <span class="pDot3">4. </span>对公司名称，专业术语等，如果企图用在线翻译等工具翻译你们的公司名称或专业术语，而你们自己不能提供的，那只能表明你们整个公司的<span class="FntF0F">悲哀</span>，做这样的语言版本是没有必要的。<br>
  <span class="pDot3">5. </span>空壳的版本 --- 对公司形象影响很坏：<br>
  如果作好一个英文或其它语言版，里面几乎没有资料，表明你们没有维护这些语言版本的能力，这样的版本多一个不如少一个 ——— 这样的版本真的对公司形象影响很坏！<br>
  <span class="pDot2">三是：</span>多语言版本显示问题：<br>
  <span class="pDot3">1. </span>如果浏览多语言版，显示乱码或方框，一般与本系统无关，<span class="FntF0F">本系统采用国际化编码</span>，只要你电脑有相关字库或字体，本系统几乎可以显示全世界所有文字！<br>
  <span class="pDot3">2. </span>如果出现上述问题，可以考虑用win2000以上系统，安装多语言包和相关字库；这里只做提示，不提供具体培训，如需要，请找你们自己相关专业人员，或上网查资料，或去电脑部培训。<br>
  <span class="pDot3">3. </span>如果用第三方软件查看多语言版本，请考虑第三方软件本身是否有问题，或找相关人员解决。<A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD align=left vAlign=top bgColor=#ffffff class="fSiz12">请您用心的发布网站资料，常常编辑更新资料，用心的管理网站！这是对我最大的支持；祝：学习工作生活愉快！ </TD>
    </TR>
    <TR>
      <TD vAlign=top align=right bgColor=#ffffff class="dCode">更新 Peace[XieYS] 
        2010-08-02 </TD>
    </TR>
  </TBODY>
</TABLE>
<script type="text/javascript">PeaceJSTest.innerHTML='<font color="#FF00FF">您启用了Javascript</font>'</script>
</body>
</html>
