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
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 border=0>
  <TBODY>
    <TR>
      <TD vAlign=top align=center bgColor=#ffffff class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 版本信息-V1.X</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
          <tr>
            <td valign="top" bgcolor="#FFFFFF">　　这是给 网站开发者查看使用的，<span class="FntF00">内容管理者请离开</span>！！！ <br>              
              <a href="#1.9" class="pILink">&middot;Ver1.9</a><a href="#1.8" class="pILink">&middot;Ver1.8</a><a href="#1.7" class="pILink">&middot;Ver1.7</a><a href="#1.6" class="pILink">&middot;Ver1.6</a><a href="#1.5" class="pILink">&middot;Ver1.5</a><a href="#1.4" class="pILink">&middot;Ver1.4</a><a href="#1.3" class="pILink">&middot;Ver1.3</a><a href="#1.2" class="pILink">&middot;Ver1.2</a><a href="#1.1" class="pILink">&middot;Ver1.1</a><a href="#1.0" class="pILink">&middot;Ver1.0</a><a href="#FlagMod" class="pILink">&middot;默认模块说明</a></td>
            <td width="25%" valign="top" bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"><!--#include file="vmenu.asp"--></td>
          </tr>
        </table></TD>
    </TR>
  </TBODY>
</TABLE>
<P class=dTitle><A id=FlagEditor8 name=1.9></A> Ver1.9</P>
<p><span class="pDot3">R.</span>重写 数据移动工具；完善权限测试等；预计2.0所要的功能部分基本完成，剩下一些帮助文档资料整理等，预计在2.1版前完成。(2010-07-31)<br>
  <span class="pDot3">Q.</span>修改 inc/home/jsdate.js -=&gt; jsdate.asp 适合4（或更多）个语言版本；(2010-07-29) <br>
  <span class="pDot3">P.</span>修改 smod/info/info_list.asp -=&gt; 文件，图片/视频下，无法设置推荐属性；(2010-07-28)  <br>
  <span class="pDot3">O.</span>完成 供求的 留言，评论，应聘，订购，准备发布2.0版啦；[<span class="FntF0F">2010-07-26近几天有点小感冒？！</span>]  (2010-07-27)<br>
  <span class="pDot3">N.</span>增加 几种(图片列表)显示模式；  (2010-07-24)<br>
  <span class="pDot3">M.</span>修正 数据存数据库下，企业介绍,在线答疑,多页介绍无法显示等重大问题；  (2010-07-23)<br>
  <span class="pDot3">L.</span>优化 数据转化/清理程序（/smod/adupd/upd_data.asp）；(2010-07-20下午又重装系统，浪费一个下午多时间，新加共享等问题.....)；  (2010-07-22)<br>
  <span class="pDot3">K.</span>供求前台显示：已完成；评论,留言,应聘,订购暂未完成； (2010-07-17)<br>
  <span class="pDot3">J. </span>完成 供求前台显示 基本完成；[<span class="FntF0F">15下午 我公司电脑装Ylmf OS，家里装空调!</span>] (2010-07-16)<br>
  <span class="pDot3">I. </span>修正 内容分页jsPager.js程序，使其点击后，跳转到页面标题处；(2010-07-14)<br>
  <span class="pDot3">H.</span>完成供求后台 会员发布信息；[<span class="FntF0F">12下午~13上午换新电脑,尽快发布本程序的2.0版本以回报!</span>] (2010-07-12)<br>
  <span class="pDot3">G.</span>完成 公文内容存文件；(2010-07-10)<br>
  <span class="pDot3">F. </span>upfile目录，除sys目录，其他目录分开，更利于大型项目管理；(2010-07-10)<br>
  <span class="pDot3">E. </span>完成 aJax判定FCKEditor 的上传权限；(2010-07-09)<br>
  <span class="pDot3">D.</span>完成 内容分页jsPager.js程序；(2010-07-08)<br>
  <span class="pDot3">C.</span>完成 论坛内容存文件 工作；增强权限管理；(2010-07-07)<br>
  <span class="pDot3">B.</span>内部会员，会员可发布后台信息；(2010-07-01 ~ 2010-07-06)<br>
  <span class="pDot3">A.</span>增加可选设置：内容存文件-内容存数据库；<br>
  已完成：介绍(招聘) 新闻 课程 部门 产品 人物 其它(3图片) 视频 下载 留言；<br>
  未完成：论坛(+PK)；公文；供求（2010-06-XX ~ 2010-06-29）  <br>
--- 2010-07-01 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagEditor7 name=1.8></A> Ver1.8</P>
<p><span class="pDot3">J. </span>增加：版权杂项 --- 信息设置 (2010-06-XX)<br>
  <span class="pDot3">I. </span>增加：[ParTemp]模版参数 (2010-06-XX)<br>
  <span class="pDot3">H. </span>增加：[ParTemp]次类别参数 (2010-06-XX)<br>
  <span class="pDot3">G. </span>会员申请，防注改进；06-12<br>
  <span class="pDot3">F. </span>配置目录替换 改成任意表任意列替换；06-11<br>
  <span class="pDot3">E. </span>缓存重新整理；06-09<br>
  <span class="pDot3">D. </span>效率测试；06-08<br>
  <span class="pDot3">C. </span>缓存重新整理；06-07<br>
  <span class="pDot3">B. </span>FCKEditor2.6，特殊符号，表情符号，上传视频，上传FLV(SWF)，服务器浏览等，完美解决；（2010-05-26）<br>
<span class="pDot3">A. </span>js 图片播放器+filters滤镜效果(for IE)；（2010-05-25）<br>
--- 2010-05-26 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagEditor6 name=1.7></A> Ver1.7</P>
<p><span class="pDot3">R. </span>bbs,binn 合并；（2010-05-24）<br>
  <span class="pDot3">Q. </span>新增参数：信息模块设置(rModTyp2.asp),在新闻中起用；（2010-05-21）<br>
  <span class="pDot3">P. </span>起用FCKEditor2.6；FCK上传/浏览与现有程序结合；改进FCK特殊符号，笑脸符号...（2010-05-21~05-24）<br>
  <span class="pDot3">O. </span>前台起用缓存（如首页）。（2010-05-20）<br>
  <span class="pDot3">N. </span>修改参数新增：大窗口模式。（2010-05-19）<br>
  <span class="pDot3">M. </span>新增：ASP缓存(/tools/out/cchClass.asp)类；修改函数(func_vbs.asp:Get_vPath(xLen)加端口)；增加代码块(ext/xunion.asp)；修改参数(sadm/system/para_rem.asp,para_set1.asp,config.asp,sadm/edfck/fckconfig.js可以Html编辑)（2010-05-18）<br>
  <span class="pDot3">L. </span>新增：内部公文+短信[企信通]二次接入代码（ext_files/msg/*.*）,使用时请稍微修改（2010-04-21~05-03）<br>
  <span class="pDot3">K. </span>新增：找会密码通用代码（ext/sadm/adm12.*）,使用时请稍微修改（2010-05-03）<br>
  <span class="pDot3">J. </span>修改：vdo/img_up.asp 加水印(2010-05-03)<br>
  <span class="pDot3">I. </span>新增：MSSQL数据库管理(压缩，备份，还原：系统与设置 &gt;&gt; 数据库管理 )；使用时请做稍微修改；(2010-04-22)<br>
  <span class="pDot3">H. </span>新增：单表自定义管理（系统与设置 &gt;&gt; 数据库管理 &gt;&gt; sys_dbinfo &gt;&gt; 强行删除表资料）,可用于管理外部资料，使用时请做稍微修改；（如XX医院-删除MSSQL资料）(2010-04-22)<br>
  <span class="pDot3">G. </span>修改：&quot;UsrPerm&quot; -=&gt;UsrPStr, UsrPStr=&quot;UsrP&quot;&amp;Left(Config_Code,5)(2010-04-20)<br>
  <span class="pDot3">F. </span>新增：rShow.asp显示风格 参数，适合多种显示风格，配套函数为:func3\GetSFlags(xMod,xType,xFlag), 标记如下（LinkDetail;LinkPic;LinkNull;Show3Def;Show2UD;Show2PC;Show2CP;Show1Cont;Show1Pic;）(2010-04-17)<br>
  <span class="pDot3">E. </span>升级：“图片广告”，使用纯js播放广告图片(swf),支持swf文件,支持IE6，IE7，FireFox  (2010-04-13)<br>
  <span class="pDot3">D. </span>修改：内外论坛 --- 我的帖子等混淆  (2010-04-09)<br>
  <span class="pDot3">C. </span>新增：系统与设置 &gt;&gt; 配置 &gt;&gt; 1Key配置  (2010-04-09)<br>
  <span class="pDot3">B. </span>新增：留言与笔记 --- 信息块管理 (2010-04-08)<br>
  <span class="pDot3">A. </span>修改：会员与查询 --- 通用查询 + (新增)首页提示(member\hmmsg,XX医院-生日祝福)。(2010-04-08)<br>
--- 2010-04-08 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagEditor5 name=1.6></A> Ver1.6</P>
<p><span class="pDot3">R. </span>增加QQ浮动模版: /img/rnd_nid/rbox_nid.htm:Rnd_n14（2010-04-03）<br>
  <span class="pDot3">Q. </span>\bbs, \binn\ 帖子数，回复数都为0的错误修正（2010-04-02）<br>
  <span class="pDot3">P. </span>\ext\ext_files\fart_order\ 收集整理方盛歐藝定单（2010-04-01）<br>
  <span class="pDot3">O. </span>屏蔽定单（默认产品不要定单，改为评论---简化！！！）( 2010-03-31)<br>
  <span class="pDot3">N. </span>ext\mail.asp 增加调试参数（每次提交到邮件都花很多时间 ...愿此次以后，调试更方便一些）( 2010-03-31) <br>
  <span class="pDot3">M. </span>sadm\user\|member\info\error.asp 简洁美化；(2010-02 ~ 2010-03)<br>
  <span class="pDot3">L. </span>smod\link\info_add.asp,info_edit.asp 小修正。(2010-02 ~ 2010-03)<br>
  <span class="pDot3">K. </span>smod\vdo\file_up.asp,sadm\edfck\...\asp\io.asp 上传小修正,使附件文件名比较统一；(2010-02 ~ 2010-03)<br>
  <span class="pDot3">J. </span>sadm\system\para_rem.asp 因与IIS过滤软件的冲突，进行小修正；(2010-02 ~ 2010-03)<br>
  <span class="pDot3">I. </span>member\admin\mconfig.asp 增加防注参数；(2010-02 ~ 2010-03)<br>
  <span class="pDot3">H. </span>member\mu_app.asp 增加Ajax参数防注，注意文件名是否有改变；(2010-02 ~ 2010-03)<br>
  <span class="pDot3">G. </span>inc\*_inc\adm_main.asp,mem_main.asp 增加安全设置；(2010-02 ~ 2010-03)<br>
<span class="pDot3">F. </span>增加常用代码:ext\code.asp,ext\code.js.asp(子域名跳转等)。(2010-02 ~ 2010-03)<br>
  <span class="pDot3">E. </span>配置向导参数，常用资料管理与设置（设置类别,版权信息,友情连接等说明）(2010-02-04)<br>
  <span class="pDot3">D. </span>新闻中心,导入。(2010-02-01) <br>
  <span class="pDot3">C. </span><span class="FntF00">代码减肥</span>……ysWeb_PubData.Peace!DB，eweb28，eweb55……。(2010-01-27)<br>
  <span class="pDot3">B. </span>gbook前台显示调整；清理 tools:char,files,js,temp目录；自定义参数(如空间大小)设置。(2010-01-26)<br>
  <span class="pDot3">A. </span>func_vbs.asp增加Get_State()函数,用于定单-状态等情况下。(2010-01-25)<br>
--- 2010-01-25 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagEditor4 name=1.5></A> Ver1.5</P>
<p><span class="pDot3">R. </span><span class="FntF00">重新整理规划/upfile/附件目录</span>。(2010-01-22)<br>
  <span class="pDot3">Q. </span>移动setup目录 。(2010-01-22)<br>
  <span class="pDot3">P. </span>增加 外部接口(ext/out/upd_js.asp) 外部js 刷新有利于 扩展。(2010-01-09)<br>
  <span class="pDot3">O. </span>加密规则调整(Config+PassWD+UserID)，有利于 Passport 扩展。(2010-01-08)<br>
  <span class="pDot3">N. </span>为管理入口等页添加 robots = noindex, nofollow 设置  。(2010-01-08)<br>
  <span class="pDot3">M. </span>修改文件：/sadm/index.asp 。(2010-01-08) <br>
  <span class="pDot3">L. </span>改目录：install  ---&gt; setup，少两个字母舒服些。(2010-01-08)<br>
  <span class="pDot3">K. </span>修正 上一篇，下一篇Bug, 用LogATime 做参考，原来是KeyID,当修改时间后，顺序错误。(2010-01-07)<br>
  <span class="pDot3">J. </span>修正 FCK编辑器 fck_file.js   配置在子目录下，上传文件的图标不显示。(2010-01-07)<br>
  <span class="pDot3">I. </span><span class="FntF00">过滤up_class.asp上传的文件</span>，如含有恶意html标记的伪装图片文件。(2010-01-05)<br>
  <span class="pDot3">H. </span>清理垃圾 (CharTable，ext_files) ，备份还原（CharTable）。(2010-01-04) <br>
  <span class="pDot3">G. </span><span class="FntF00">更新几个重要目录</span>：home/-=&gt;smod/； script/-=&gt;upfile/；idoc/-=&gt;doc/；(2009-12-30)<br>
  <span class="pDot3">F. </span>一系列小调整和改善；添加一组简易调查，添加几组图片后台管理，见ext/ext_files/pic_*；(2009-12-25)<br>
  <span class="pDot3">E. </span>启用可选用模块，用于扩展；目录:ext/ext_files/；说明:ext/ext_read.txt。(2009-12-25)<br>
  <span class="pDot3">D. </span>缩略图优化改进：新闻，图片，视频，投票的[编辑&gt;缩略图]优化改进。(2009-11-25)<br>
  <span class="pDot3">C. </span>留言与笔记改进：留言与笔记可以使用编辑器。(2009-11-24)<br>
  <span class="pDot3">B. </span>增加：“<span class="FntF0F">系统编辑器</span>”配置参数：默认为:FCK2.5; eweb28:ewebeditor2.8; eweb55:ewebeditor5.5；适用于新闻，图片，视频等栏目。(2009-11-21)<br>
  <span class="pDot3">A. </span>增加：编辑器字体设置选项，如黑体，宋体；注意，如果客户端无相应字体，则显示默认字体，如“幼圆”。<br>
  --- 2009-11-16 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagEditor3 name=1.4></A> Ver1.4</P>
<p><span class="pDot3">R. </span>屏蔽：图片，视频的“设置_UBB”功能。(2009-11-10)<br>
  <span class="pDot3">Q. </span>增加：评论“设置_显示”功能。(2009-11-03)<br>
  <span class="pDot3">P. </span>视频到评论，友情连接增加，修正两个小bug。<br>
  <span class="pDot3">O. </span>增加：信息发布统计。(2009-10-17)<br>
  <span class="pDot3">N. </span>浮动QQ，全面支持：QQ，MSN，Skype，阿里巴巴，Taobao旺旺等及时通讯。(2009-10-16) <br>
  <span class="pDot3">M. </span>增加：头条设置后台程序，前台未起用。(2009-10-14)<br>
  <span class="pDot3">L. </span>简洁前台/peng/目录，前台起用图片广告设置。(2009-10-14)<br>
  <span class="pDot3">K. </span>增加：统计计数12种图片数字风格（以前为纯文字）。(2009-10-14)<br>
  <span class="pDot3">J. </span>增加：系统菜单管理，用于动态设置菜单（特定情况下，进行较复杂设置）。(2009-10-09)<br>
  <span class="pDot3">I. </span>增加：内部公文系统[/doc/]。<br>
  <span class="pDot3">H. </span>增加：多语言版，增加《清理》程序，与《同步》相对。(2009-09-29)<br>
  <span class="pDot3">G. </span>增加：简洁前台/peng/目录。(2009-09-28)<br>
  <span class="pDot3">F. </span>增加：RSS订阅功能。(2009-09-28)<br>
  <span class="pDot3">E. </span>增加：编辑器FTP结合：编辑器与大文件管理(FTP)结合[<a href="../../smod/vdo/file_view.asp">/smod/vdo/file_view.asp</a>]；<br>
  <span class="pDot3">D. </span>增加：Editor使用说明：Editor常用功能演示；<br>
  <span class="pDot3">C. </span>增加：图片资料管理说明：多语言版本图片资料管理说明；<br>
  <span class="pDot3">B. </span>增加：FTP使用提示：增加FTP使用提示；<br>
  <span class="pDot3">A. </span>增加小认证码：增加小认证码，用于特殊场合，与现有程序兼容[/sadm/login.asp]；<br>
--- 2009-09-27 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagEditor2 name=1.3></A> Ver1.3</P>
<p> <span class="pDot3">Q. </span>升级peace_player播放器：完全兼容前一版本；处理空连接；可用毫秒控制；sysPath用完全路径,sysfNO可取消；<br>
  <span class="pDot3">P. </span>增加：刷新与杂项，系统与设置 的菜单说明；<br>
  非开发模式(≠isExpert)下，模块 参数 设置项做更合理的屏蔽和提示(一客户
  <!--qxisc-->
  为难我的结果)。（2009-09-21)<br>
  <span class="pDot3">O. </span>会员注册页：增加用户名,认证码项目 的 Ajax提示。（2009-09-21)<br>
  <span class="pDot3">N. </span>home,setup,sadm,script,tools等目录增加(Safe)安全指向文件,指向首页。（2009-09-21)<br>
  <span class="pDot3">M. </span>从Ver1.3开始，版本信息倒序排列。（2009-09-18)<br>
  <span class="pDot3">L. </span>修正多语言版中，新闻详细页的“上(下)一篇”提示总为中文。（2009-09-18)<br>
  <span class="pDot3">K. </span>修改发邮件代码，邮件主体用unicode编码。（2009-09-18）<br>
  <span class="pDot3">J. </span>修改前台config文件。（2009-09-18）<br>
  <span class="pDot3">I. </span>修正清理member后，增加管理员出错。（2009-09-18）<br>
  <span class="pDot3">H. </span>增加-导入缓存中的商品到购物车-按钮。（2009-09-17）<br>
  <span class="pDot3">G. </span>改进-清理程序，已经清理过的程序，不再显示。（2009-09-17）<br>
  <span class="pDot3">F. </span>增加网站空间报警程序(在登陆首页)（为租用空间者提供方便）。（2009-09-16）<br>
  <span class="pDot3">E. </span>订购产品，提示登陆/注册（会员+订购）。（2009-09-15）<br>
  <span class="pDot3">D. </span>视频下载 模块中，选择文件时，增加上传功能（为上传小文件提供方便）。（2009-09-15）<br>
  <span class="pDot3">C. </span>增加首页SEO参数设置，可扩充内页SEO参数。（2009-09-14）<br>
  <span class="pDot3">B. </span>后台图片视频增加时模版不能显示问题。<br>
  <span class="pDot3">A. </span>前台定单提交无定单项问题。<br>
--- 2009-09-10 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A> </p>
<P class=dTitle><A id=FlagEditor name=1.2></A> Ver1.2</P>
<P><span class="pDot3">A. </span>增加-供求后台发布(管理员可发 布行业新闻 产品供求，之前仅会员才可发布)<br>
  <span class="pDot3">B. </span>增加-供求前台 (会员黄页 会员供求 行业新闻 企业招聘 )<br>
  <span class="pDot3">C. </span>改进-认证码 (9种认证码可随机组合成认证码群)<br>
  <span class="pDot3">D. </span>增加-产品购物车 (可设置 付款方式 - 配送方式 - 时间要求等参数)<br>
  <span class="pDot3">E. </span>会员注册-企业,个人,团体分类<br>
  <span class="pDot3">F. </span>会员注册-反注册机程序: 增加系统参数(防注册机随机码系列,防注册机会员ID,..认证码配置,管理入口随机码等)<br>
  <span class="pDot3">G. </span>增加-图片广告,文字广告管理，启用peace_pic_player.swf图片播放器<br>
  <span class="pDot3">H. </span>新闻与介绍,图片与视频,留言启用分部门权限管理(2009-09-09)<br>
  <span class="pDot3">I. </span>增加-安装程序 ( /setup/setup.asp ) (2009-09-09) <br>
  <span class="pDot3">J. </span>增加-清理程序 (清理setup,清理BBS,清理trade,清理vote,清理member) (2009-09-10) <br>
  --- 2009-09-08 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagEditor2 name=1.1></A> Ver1.1</P>
<P><span class="pDot3">A. </span>废除eWebeditor编辑器, 改用FCK编辑器<br>
  --- 2009-07-07 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF4 name=1.0></A> Ver1.0</P>
<P><span class="pDot2">1.1.</span> 特色功能 A:<br>
  <span class="pDot3">无限级树型，开合菜单：</span>公用函数，css控制前台，兼容FF,IE,有无DOCTYPE兼容；<br>
  <span class="pDot3">utf-8编码：</span>日，韩多语言版本，Ajax,Flash+XML中有长远优势；<br>
  <span class="pDot3">水印开关：</span>文字,图片水印 开关；<br>
  <span class="pDot3">编辑器：</span>5.5商业版本，图片删除同步，权限认证；<br>
  <span class="pDot3">多语言版本：</span>同步处理，图片共用；<br>
  <span class="pDot3">多语言版前台：</span>语言包控制，节省共用代码；<br>
  <span class="pDot3">繁体，js转化：</span>维基百科 简繁对应表；<br>
  <span class="pDot3">图片等比缩放：</span>js处理；<br>
  <span class="pDot3">认证码：</span>所有入口，图片认证码；<br>
  <span class="pDot3">动态栏目，参数：</span>前后台栏目，动态设置，参数控制；<br>
  <span class="pDot3">动态设置：</span>广告动态设置，兼容FF,IE,有无DOCTYPE兼容<br>
  <span class="pDot3">新闻，图片维护小细节：</span>增加资料时返回处理，增加资料可设置默认模版（）<br>
  <span class="pDot3">新闻，图片标题：</span>可设置颜色；<br>
  <span class="pDot3">前台兼容：</span>前台FF,IE6，IE7兼容测试；<br>
  <span class="pDot3">数据库：</span>数据库压缩备份功能；<br>
  <span class="pDot3">信息交互：</span>评论，应聘，订购处理；<br>
  <span class="pDot3">会员互动：</span>会员与管理员可互动留言；<br>
  <span class="pDot3">上下页：</span>前台浏览详细的新闻，图片时，有上下页（篇）；<br>
  <span class="pDot3">友情连接：</span>可添加文字，图片友情连接；<br>
  <span class="pDot3">首页综合js效果：</span>首页各种js效果综合显示，FF,IE6，IE7兼容测试：<br>
  <span class="pDot3">计数器：</span>自带有计数器；<br>
  <span class="pDot3">工具箱：</span>万年历，算24，素数，任意进制转换，阿江ASP 探针，木马扫把（文件，数据库扫描），安全中心（ IP跟踪,IP限制，防注入警告，防注入反攻击）... <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A><br>
  <span class="pDot2">1.2.</span>特色功能  B:<br>
  <span class="pDot3">SQL防注入：</span>系统本身防注入；另加 安全中心的 IP跟踪,IP限制，防注入警告，防注入反攻击程序，可自由启用；<br>
  <span class="pDot3">站点配置：</span>自己慢慢阅读相关资料；<br>
  <span class="pDot3">系统LOG记录：</span>自带有系统LOG记录功能；<br>
  <span class="pDot3">扩展类别设置：</span>提供丰富参数接口；<br>
  <span class="pDot3">留言笔记：</span>自带有站务笔记，私人笔记（秘密）；<br>
  <span class="pDot3">权限分配：</span>按模块分配权限。<br>
  --- 2009-04-28 <A href="#"><SPAN class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF2 name=FlagMod></A> 默认 模块说明（Ver1.0,实际在不断更新中...）</P>
<table width="750" border="0" align="center">
  <tr>
    <td valign="top"><span class="pDot2">2.1. </span>新闻与介绍 <br>
      企业介绍 增加 评 类别 <br>
      招聘信息 增加 申 类别 <br>
      新闻中心 增加 评 类别</td>
    <td valign="top"><span class="pDot2">2.2.</span> 图片与视频 <br>
      产品图片 增加 单 类别 <br>
      人物图片 增加 评 类别 <br>
      其它图片 增加 评 类别 <br>
      视频下载 增加 评 类别 <br>
      备用图片 增加 分类</td>
    <td valign="top"><span class="pDot2">2.3.</span> 留言与笔记 <br>
      (中文)访客留言 分类 <br>
      (英文)访客留言 分类 <br>
      站务笔记 增加 分类 <br>
      私人秘密 - 增加秘密</td>
  </tr>
  <tr>
    <td valign="top"><span class="pDot2">2.4. </span>会员与查询 <br>
      会员管理 参数 注册 <br>
      会员反馈 通知 笔记 <br>
      通用查询 分类 导入</td>
    <td valign="top"><span class="pDot2">2.5. </span>论坛与投票 <br>
      公开论坛帖子 - 分类 <br>
      观点辩论 观点 类别 <br>
      内部论坛帖子 - 分类 <br>
      投票项目 类别 记录 <br>
      问卷调查 类别 记录</td>
    <td valign="top"><span class="pDot2">2.6.</span> 商务与供求 <br>
      企业资料 参数 类别 <br>
      企业介绍 ------- 类别 <br>
      新闻中心 ------- 类别 <br>
      供求 -- 交易 -- 类别 <br>
      招聘 -- 应聘 -- 类别</td>
  </tr>
  <tr>
    <td valign="top"><span class="pDot2">2.7.</span> 刷新与杂项 <br>
      信息刷新 -- 管理帮助 <br>
      浮动广告 -- 统计计数 <br>
      友情连接 增加 分类</td>
    <td valign="top"><span class="pDot2">2.8.</span> 系统与设置 <br>
      系统记录 模块 参数 <br>
      个人密码 管理员列表 <br>
      扩展类别 数据库管理 <br>
      信息类别 配置 工具</td>
    <td valign="top">&nbsp;</td>
  </tr>
</table>
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD vAlign=top align=right bgColor=#ffffff class="dCode">更新 Peace[XieYS] 
        2009-04-24 ~ 2010-07-31 &nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
