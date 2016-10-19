<!--#include file="../himg/tconfig.asp"-->
<%

Call Chk_Perm1(xPara,"")

%>
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
      <TD vAlign=top align=center bgColor=#ffffff class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 版本信息-V2.X</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
          <tr>
            <td valign="top" bgcolor="#FFFFFF">　　这是给 网站开发者查看使用的，<span class="FntF00">内容管理者请离开</span>！！！ <br>
              <a href="#2.3" class="pILink">&middot;Ver2.4</a><a href="#2.3" class="pILink">&middot;Ver2.3</a><a href="#2.2" class="pILink">&middot;Ver2.2</a><a href="#2.1" class="pILink">&middot;Ver2.1</a><a href="#2.0" class="pILink">&middot;Ver2.0</a></td>
            <td width="25%" valign="top" bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"><!--#include file="vmenu.asp"--></td>
          </tr>
        </table></TD>
    </TR>
  </TBODY>
</TABLE>
<P class=dTitle><A id=FlagEditor3 name=2.4></A> Ver2.4</P>
<P><span class="pDot3">04010. </span>修正 浏览器检测(Firefox,Chrome)，浏览器兼容(Chrome)，PHP未同步；(2011-10-28)<br>
  <span class="pDot3">04009. </span>整理精简kernel目录(1.1M)，精简[个人笔记](1.8M)，共精简2.9M左右。(2011-10-27)<br>
  <span class="pDot3">04008. </span>更新 bbs目录，使论坛可用新编辑器；(2011-10-26) <br>
  <span class="pDot3">04007. </span>修正 升级编辑器后，查看空间使用量的错误；(2011-10-19) <br>
  <span class="pDot3">04006. </span>恢复 10-06格式化磁盘了, 恢复自09-08~10-30之间的部分资料；编辑器升级 KindEditor(4.0~4.0.1)；(2011-10-11) <br>
  <span class="pDot3">04005. </span>增加 内容模版/特殊字符 插入API；(2011-09-26) <br>
  <span class="pDot3">04004. </span>增强 附件管理 功能，可以在此文件中 直接插入文件/图片/视频到Editor；(2011-09-23) <br>
  <span class="pDot3">04003. </span>一个参数 切换6个编辑器(CK3 - Kind - xh1 - eWeb - U1 - TinyMCE - FCK 等)；asp下，CK3 - Kind - xh1 - FCK 有文件上传；asp下，CK3 - Kind - FCK 有文件上传；其它文件上传，暂未配制；(2011-09-10) <br>
  <span class="pDot3">04002. </span>增加 CKEditor,可设置参数切换CK/Kind/FCK；(2011-09-01)<br>
<span class="pDot3">04001. </span>V2.4发布, 增加 KindEditor, 可设置参数切换Kind/FCK；2011-08-31 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagEditor name=2.3></A> Ver2.3</P>
<P><span class="pDot3">03045. </span>整理 个人笔记中，几条私人记录整理分离(个人笔记~保密项目)；(2011-08-22)<br>
  <span class="pDot3">03044. </span>增加 IE6倒计时提醒(后台登陆入口)；(2011-08-18)<br>
  <span class="pDot3">03043. </span>修改 把原来exp/apis/改为exp/api/；(2011-07-30)<br>
  <span class="pDot3">03042. </span>重新配置 asp,php,net三个项目，将发布在oa.dg.gd.cn虚拟目录下；(2011-07-27)<br>
  <span class="pDot3">03041. </span>修正 服务跟踪 中 第二次点[返回并刷新]出现的错误；(2011-07-20)<br>
  <span class="pDot3">03040. </span>增加 dbm.asp数据库管理工具(2011-07-20)<br>
  <span class="pDot3">03039. </span>修改 js:setImgSiz2-&gt;setImgSize函数并优化(之前3个参数,现在优化为只传一个参数）(2011-07-11)<br>
  <span class="pDot3">03038. </span>增加 定制管理菜单（入口：管理中心 &gt;&gt; 修改管理密码/定制管理菜单）(2011-07-04)<br>
  <span class="pDot3">03037. </span>增加 类别ID调换/更改（入口：各类别列表）(2011-07-04)<br>
  <span class="pDot3">03036. </span>增加 定单统计报表（产品与图片 &gt;&gt; 定单 &gt;&gt; (左上的)报表 &gt;&gt;）(2011-06-24)<br>
  <span class="pDot3">03035. </span>改进 类别分组管理 type_center.php；(2011-06-04面试C# 没有带简历，再次迷茫ing；)(2011-06-08)<br>
  <span class="pDot3">03034. </span>完善 服务跟踪(/tools/help/ctrace.asp)项，服务器端地址拟定为 (http://about.dg.gd.cn/tools/help/strace.asp)；且动态更新服务器端；(2011-05-13~14)<br>
  <span class="pDot3">03033. </span>迷茫 了解模板Smarty，比较框架ThinkPHP/CanPHP/Fleaphp，了解Perl/Python/Ruby等，迷茫中...还是尽快完善<span class="FntF0F">我自己的PHP2.3</span>；(2011-05-11)<br>
  <span class="pDot3">03032. </span>增加 服务跟踪(/smod/adupd/trace.asp)项 又马上屏蔽，又一次确认[<span class="FntF00">暂时停止 重大更新，转向PHP</span>]；(2011-05-09)<br>
  <span class="pDot3">03031. </span>增加 参数(SwhShowSpace)显示<span class="FntF0F">隐藏占用空间信息</span>；(2011-05-09)<br>
  <span class="pDot3">03030. </span>附件管理 增加显示附件大小；(2011-04-28)<br>
  <span class="pDot3 Fnt00F">03014. </span>再次确认[<span class="FntF00">暂时停止 重大更新，转向PHP</span>]；修改 上传检测，想提高效率，<span class="Fnt00F">但几乎没有什么收获</span>！(2011-04-28)<br>
  <span class="pDot3">03029. </span>短信系统 增加php,c#下的编码函数；(2011-04-25)<br>
  <span class="pDot3">03028. </span>整理 目录\ext\apis（cset,eplay,iis,map,menu,mzoo,rss,scan,scroll4,seo）；(2011-04-22)<br>
  <span class="pDot3">03027. </span>修改 类别更新程序（smod/type/code_tupd.asp）；(2011-04-21)<br>
  <span class="pDot3">03026. </span><span class="FntF0F">增加 jquery库</span>（\inc\home\jquery-1.5.2.min.js），/tools/mzoom/pic_show.asp首次引用；(2011-04-21)<br>
  <span class="pDot3">03025. </span>1Key配置 重新改写[1Key配置]；修改[清理垃圾]，提高效率；(2011-04-20)<br>
  <span class="pDot3">03025. </span>会员注册 增加防Session过期；评论(remark)增加删除当前评论操作(2011-04-19)<br>
  <span class="pDot3">03024. </span>短信系统 增加新注册会员/新增加定单通知；(2011-04-19)<br>
  <span class="pDot3">03023. </span>短信系统 增加短信群发；(2011-04-18)<br>
  <span class="pDot3">03022. </span>短信系统 增加:电话薄-批量增加,群发组设置功能,用于群发；突破最大号码限制；(2011-04-18)<br>
  <span class="pDot3">03021. </span>调整 几个关键函数(Get_MDHNSX/Get_YMDX ---&gt; Get_FmtID, Rnd_ID2                      ---&gt; Rnd_Base, Get_OrdID                    ---&gt; rs_OrdID, rs_AutID)，是其功能更灵活强大，且更合理；(2011-04-18)<br>
  <span class="pDot3">03020. </span>增加 3个参数(<span class="FntF0F">用于无限扩展信息模块功能</span>的参数见Fix03013)，改为手动更新；(2011-03-28)<br>
  <span class="pDot3">03019. </span>修正 会员导出Excel，使导出更加人性化，更加完美；(2011-03-28)<br>
  <span class="pDot3">03018. </span>修正 水印，检查类型；修正 会员类别IP搜索；(2011-03-25)<br>
  <span class="pDot3">03017. </span>增加整理 5款图片播放效果(eplay\11~15)；(2011-03-24)<br>
  <span class="pDot3">03016. </span>删除 表单参数；发现表单参数成为鸡肋；删除(并备份\ext\xtemp\form.asp)；(2011-03-22)<br>
  <span class="pDot3">03015. </span>更新 代码生成器；更新 会员申请MemID处理；(2011-03-15~16)<br>
  <span class="pDot3">03014. </span><span class="FntF00">暂时停止 重大更新，转向PHP,C++！</span>其实，更新是无止境的，但目前，恐怕会弄的越来越笨重，丧失效率...；下一步重大更新：可能是Editor的FCK到CK更新，或dll+主机管理系统，或Member的防注机制更新(半成品文件和参数：表单参数,xtemp/app4x.asp)；今天：3.14 07:22 圆周率日圆周率时分；(2011-03-14)<br>
  <span class="pDot3">03013. </span>增加 3个参数(Config_dbTab-资料存放表格; Config_upDir-附件存放目录; Config_mdKey-模块标识前缀)，<span class="FntF0F">用于无限扩展信息模块功能</span>；如增加[商务信息],[专题信息],[分类信息](T04Info-dt004-M04124)等，只增加参数，增加相关数据表即可，其它发布管理页可共用；也可用于把现有模块如视频从图片模块中分离。(2011-03-12)<br>
  <span class="pDot3">03012. </span>增加 可选发邮件组件(JMail=JMAIL.Message,CDONTS=CDONTS.NewMail,CDO=CDO.Message)；(2011-03-04)<br>
  <span class="pDot3">03011. </span>增加 文字水印描边功能(<span class="FntF0F">Persits.Jpeg本身无描边功能, Peace程序实现</span>)；(2011-03-04)<br>
  <span class="pDot3">03010. </span>增加 图片认证码自定义Session；(2011-03-01)<br>
  <span class="pDot3">03009. </span>整理 使文件，目录结构，表结构更合理：1. 配置信息从WebInfo表移动到AdmPara参数表；2.水印参数专门设置；3. WebInfo表改为WebAdvart，只存放广告，系统菜单(MrXItem9)，外部连接(OutLink)；4. Config.asp文件，合并配置,系统,数字,开关四类参数；(2011-02-25)<br>
  <span class="pDot3">03008. </span>增加 为每个短信会员，自动生成调外部用代码，呵呵，可以“量产”了；(2011-02-17)<br>
  <span class="pDot3">03007. </span>增加 通用短信接口 的 多语言支持；(2011-02-15)<br>
  <span class="pDot3">03006. </span>增加 函数Get_A30SN，Get_A30Chk，用于通用的安全认证；增加函数echo(xStr),主要用语调试；(2011-02-14) <br>
  <span class="pDot3">03005. </span>增加 参数SwhModGroup（管理菜单分组参数）,适合大型站点需要；(2011-02-12) <br>
  <span class="pDot3">03004. </span>增加 <span class="FntF0F">通用短信接口</span>，目前支持：bucp.net博星SDK，noName(张长江提供)和一个模拟测试接口(test.Peace)；扩展接口，只增加一个api_xxx.asp文件，配置相应参数，重写相关发信息等必要函数即可；(2011-02-11) <br>
  <span class="pDot3">03003. </span>修改 生成静态页: .stm改为.htm后缀(为了安全性，没办法用个麻烦点的)；(2011-01-19) <br>
  <span class="pDot3">03002. </span>增加 guest发布；<br>
<span class="pDot3">03001. </span>增加 函数Check_RTest，正则表达式检查字符串；修改 函数fil_exist,fil_del,简化参数，chk_file；使用正则表达式up_class.asp:ChekFType上传检查，更准确；2010-01-17 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagEditor name=2.2></A> Ver2.2</P>
<P><span class="pDot3">02064. </span>修改 FCKEditor，插入附件,插入/编辑超链接，插入的附件，自动用文件名作为Title，且提示信息不是路径，而是原文件名；(2011-01-13) <br>
  <span class="pDot3">02063. </span>增加 功能：脚本加密解密（入口：工具 &gt;&gt; 脚本加密解密） ；(2011-01-10) <br>
  <span class="pDot3">02062. </span>增加 功能：<span class="FntF0F">自动取内容第一个图为附图</span>，自动缩略（模版参数：tmsPara.asp设置） ；(2011-01-10)<br>
  <span class="pDot3">02061. </span>修改 问卷调查，使其支持单选，多选，可填，必填（之前无单选) ；(2011-01-08)<br>
  <span class="pDot3">02060. </span>增加整理 Asp木马扫描器：(入口：工具 &gt;&gt; Asp木马扫描器 &gt;&gt; ) ；(2011-01-06)<br>
  <span class="pDot3">02059. </span>整理 说明文件：<span class="FntF0F">服务器设置参考文件 </span>(入口：管理帮助 &gt;&gt; ) ；(2011-01-05)<br>
  <span class="pDot3">02058. </span>更新 <span class="FntF0F">上传文件安全更新</span>(\sadm\func1\up_class.asp|func_file.asp:chk_file) ；(2011-01-04)<br>
  <span class="pDot3">02057. </span>增加 未审信息(类别管理&gt;&gt;入口) ；(2010-12-30)<br>
  <span class="pDot3">02056. </span>启用 发信息审核 的权限设置；(2010-12-30)<br>
  <span class="pDot3">02055. </span>增加 后台登陆入口检查浏览器 ；(2010-12-30)<br>
  <span class="pDot3">02054. </span>增加 未审评论(类别管理&gt;&gt;入口) ；(2010-12-29)<br>
  <span class="pDot3">02053. </span>增加 论坛关闭参数(SwhBBSStop) 和 论坛关闭通知参数(BBSStop.htm)， 后台控制论坛是否关闭；(2010-12-29)<br>
  <span class="pDot3">02052. </span>修改 参数tmsTyp2.asp使InfTyp2可引用InfType资料；修改 参数tmsPara.asp使其可设置默认参数(Rem,Vote,Next)；(2010-12-22)；<br>
  <span class="pDot3">02051. </span>修改 /tools/menu/peaceMenu01.htm 兼容IE8；(2010-12-22)；<br>
  <span class="pDot3">02050. </span>更新 Editor-ie6兼容文件(/sadm/edfck/editor/skins/default/fck_dialog_ie6.js) 之前修改时有误操作；(2010-12-18)；<br>
  <span class="pDot3">02048. </span>整理 子域名跳转文件(/ext/go.asp) 用于子域名跳转，Blog等；(2010-12-18)；<br>
  <span class="pDot3">02047. </span>增加函数(func_const.asp:SubDirect) 用于域名跳转，SEO优化等；(2010-12-18)；<br>
  <span class="pDot3">02046. </span>优化 内部公文(/doc/*.*)；(2010-12-16)；<br>
  <span class="pDot3">02045. </span>增强 站务笔记导出功能(/out/admlogs.asp)；(2010-12-15)；<br>
  <span class="pDot3">02044. </span>增加 搜索引擎,截断程序(函数) （func_perm.asp:Function ChkSpider），用于管理页等屏蔽搜索引擎；(2010-12-15)；<br>
  <span class="pDot3">02043. </span>增加 密码强度检测（/inc/home/jsPass.js），用于申请会员；(2010-12-15)；<br>
  <span class="pDot3">02042. </span>内部公文 细分为公开公文,定向接收菜单； (2010-12-15)；<br>
  <span class="pDot3">02041. </span>增加 代码生成器，自动生成导入,增加,修改,显示数据库字段的常用代码（/member/ecard/admimp_code.asp，入口：系统与设置 &gt;&gt; 数据库管理 &gt;&gt; dbtabs &gt;&gt; 生成 &gt;&gt;）； (2010-12-15)；<br>
  <span class="pDot3">02040. </span>增加 两个限时安全函数（/sadm/func1/func_time.asp:Get_AHChk,Get_AHour）；(2010-12-10)；<br>
  <span class="pDot3">02039. </span>更新/整理（/inc/home/jsInfo.js,jsPlugs.js）；增强FF兼容性；(2010-12-07)；<br>
  <span class="pDot3">02038. </span>更新 使用空间统计（/sadm/user/space.asp）；可自定义，添加虚拟目录；(2010-11-29)；<br>
  <span class="pDot3">02037. </span>更新/整理 简繁转化，使（inc/home/conv002.js,convert.js）共用转化表；(2010-11-25)；<br>
  <span class="pDot3">02036. </span>更新/整理 前台播放器（/inc/home/playpic.asp）；(2010-11-25)；<br>
  <span class="pDot3">02035. </span>更新 使用空间统计（/sadm/user/space.asp）更精准；可适应大站模式的虚拟目录；(2010-11-25)；<br>
  <span class="pDot3">02034. </span>增加 func3.asp:get_1Pic函数；修改func3.asp:ListTemp,ListLink等函数；(2010-11-23)；<br>
  <span class="pDot3">02033. </span>增加 inc/home/conv002.js简繁转化 插件,兼容FF；增加 得到默认信息参数的函数func_sfile.asp:get_TmrPara(xSet)；(2010-11-18)；<br>
  <span class="pDot3">02032. </span>修改 /smod/file/img_set.asp图片路径，smod/ibat/rpic_up.asp上传路径，/tools/mzoom/lbox-script.js鼠标效果代码；(2010-11-18)；<br>
  <span class="pDot3">02031. </span>增加 js<span class="FntF0F">自动获取关键字</span> 插件(/inc/home/jskeys.js)，用于信息发布/修改；(2010-11-17)；<br>
  <span class="pDot3">02030. </span>整理 完全兼容MSSQL数据库(/sadm/func2/func_const.asp:MSSQL配置/Access配置)；(2010-11-12)；<br>
  <span class="pDot3">02029. </span>整理 js工具(插件)：(tools/iis/tabconv/out_html.js; home/inc/isBrows.js,jsEvent.js)等全部放在 inc/home/jsPlugs.js文件中；(2010-11-08)；<br>
  <span class="pDot3">02028. </span>增加 工具-表格行列转换(/tools/iis/tabconv/tabconv.htm)；(2010-11-04)；<br>
  <span class="pDot3">02027. </span>增加 js插件(/tools/iis/tabconv/out_html.js)，在FF下支持outerHTML，用于表格行列转换；(2010-11-04)；<br>
  <span class="pDot3">02026. </span>设置 后台登陆文件以adm开头；(2010-11-02)；  <br>
  <span class="pDot3">02025. </span><span class="FntF0F">再见PHP！</span>工作要忙死了；早两天宝宝有点不乖...(暂时离开PHP一段时间...)。(2010-11-01)<br>
  <span class="pDot3">02024. </span>修改 背景图，干掉几个png背景图，改用gif背景图，使他们在IE6可下显示。(2010-11-01)<br>
  <span class="pDot3">02023. </span>修改 tools/_admin/tperm.asp文件,当目录不存在（主要是去掉一些不要的模块）时的错误；(2010-10-26)；<br>
  <span class="pDot3">02022. </span>合并文件 func_obj.asp, func_mail.asp, img_wmark.asp合并到func_obj.asp文件；(2010-10-25)；<br>
  <span class="pDot3">02021. </span>批量分析 IIS站点（tools/iis/qiis.asp，使用02020的函数）；(2010-10-25)； <br>
  <span class="pDot3">02020. </span>增加函数 执行命令（Do_Cmd(xStr)如Do_Cmd(&quot;ping www.dg.gd.cn -w 1 -n 1&quot;)）；(2010-10-25)；<br>
  <span class="pDot3">02019. </span>优化 3函数(Get_rsOpt,Get_rsOpt2,Get_rsCBox)；(2010-10-19)；<br>
  <span class="pDot3">02018. </span>收集 播放器（/inc/home/playpid.swf）；(2010-10-19)；<br>
  <span class="pDot3">02017. </span>增加 Editor参数组设置，限制Editor中，图片的最大宽度，高度；(2010-10-14)；<br>
  <span class="pDot3">02016. </span>增加 数据库结构/资料导出功能，供Mysql使用(导入) --- <span class="FntF0F">本系统不能局限于Windows+IIS了</span>；(2010-10-12)；<br>
  <span class="pDot3">02015. </span>增加 (集成)IIS查询分析器，表格排序；(2010-10-11)；<br>
  <span class="pDot3">02014. </span>修改 批量下载，<span class="FntF0F">效率还是不如java等程序的多线程...</span>；(2010-10-09)；<br>
  <span class="pDot3">02013. </span>修改 代码导出程序，使改变目录时，自动找子目录；(2010-10-08)；<br>
  <span class="pDot3">02012. </span>修改 对时程序；(2010-10-07)；<br>
  <span class="pDot3">02011. </span>增加 <span class="FntF0F">批量下载,代码导出</span>(/tools/out/code.asp,down.asp)；(2010-10-06)；<br>
  <span class="pDot3">02010. </span>增加 &lt;%NoDown%&gt;表，进一步加强安全防范(From:良精志诚科技liangjing.net)；(2010-09-30)；<br>
  <span class="pDot3">02009. </span>增加 一种认证码(/sadm/pcode/img_codk.asp)；(2010-09-30)；<br>
  <span class="pDot3">02008. </span>增加 会员导出Excel(一客户要求，导出的Excel仿真度还非常高，特别是表格线...)；(2010-09-28)；<br>
  <span class="pDot3">02007. </span>修正 信息复制时同时复制图片；复制和同步时同时增加相关图片；(2010-09-26)；<br>
  <span class="pDot3">02006. </span>增加 参数专用设置页(/smod/info/set_para.asp,具体应用时请重新修改此文件)；(2010-09-26)；<br>
  <span class="pDot3">02005. </span>增加 参数容量(改成备注字段，1~96个参数，单个参数最大600字符)；(2010-09-25)；  <br>
  <span class="pDot3">02004. </span>增加 单页定向采集(/smod/ibak/xx01-imp.asp,具体应用时请重新修改此文件)；(2010-09-25)；<br>
  <span class="pDot3">02003. </span>增加 公用采集规则 和 单页采集设置(但未使用)；(2010-09-24)；<br>
  <span class="pDot3">02002. </span>增加 无缝上下左右滚动加定高定宽停顿效果[兼容ie/ff](位置:/tools/scroll4/)；(2010-09-21)；<br>
  <span class="pDot3">02001. </span>增加 <span class="FntF0F">相关图片批量上传</span>(位置/smod/rpic_*3文件)；(2010-09-20)；<br>
--- 2010-09-20 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagEditor2 name=2.1></A> Ver2.1</P>
<P><span class="pDot3">Y. </span>增加 <span class="FntF0F">信息文章读后感受投票功能</span>(以插件形式，放在/tools/nmood/目录)；(2010-09-19)<br>
  <span class="pDot3">X. </span>增加 <span class="FntF0F">信息复制功能</span>(信息列表-高级设置-复制,用于添加相似的产品等)；(2010-09-18)<br>
  <span class="pDot3">W. </span>增加 <span class="FntF0F">信息采集功能</span>(可采集含缩略图的信息，文本中的图片，暂时未处理)；(2010-09-17)<br>
  <span class="pDot3">V. </span>增加 对普通信息内容，<span class="FntF0F">保存远程图片到本地</span>功能(开关参数:SwhRemSave[保存远程图片]控制)；(2010-09-17)<br>
  <span class="pDot3">U. </span>增加 保存远程图片到本地函数(upremote.asp:RemoteReplaceUrl)，将用于信息采集等功能；(2010-09-16)<br>
  <span class="pDot3">T. </span>增加 列表和正文内容的隐藏标记(.Peacd.)，便于不同站点之间的信息采集(批量导入)；(2010-09-15)<br>
  <span class="pDot3">S. </span>增加 <span class="FntF0F">图片批量上传</span>，开始筹划信息采集(批量导入)的功能；(2010-09-14)<br>
  <span class="pDot3">R. </span>增加 前台留言，是否开启编辑框开关（开关参数:SwhGbkEditor）；(2010-09-11)<br>
  <span class="pDot3">Q. </span>修正 检查附图非法输入，但可复制，粘贴等操作；(2010-09-11)<br>
  <span class="pDot3">P. </span>修正 同步操作时，KeyID重复等出错；(2010-09-11)<br>
  <span class="pDot3">O. </span>修改 函数GetNLay(xMD,xID,xLink,xStr1,xStr2)，使类别可以有连接；(2010-09-10)<br>
  <span class="pDot3">N. </span>增加 一个图片播放器（/inc/home/playp02.swf）；(2010-09-10)<br>
  <span class="pDot3">M. </span>增加 应用两个图片放大效果插件(tools/mzoom/)；(2010-09-10)<br>
  <span class="pDot3">L. </span>增加 js格式化工具(/tools/peace/js-fmt.htm)；(2010-09-09)<br>
  <span class="pDot3">K. </span>增加 pfile/lang/目录，统一管理多语言版的语言包；(2010-09-09)<br>
  <span class="pDot3">J. </span>修改 <span class="FntF0F">会员改成多语言版(简，繁，英)</span>；(2010-09-06)<br>
  <span class="pDot3">I. </span>增加 会员增加繁体转化版；(2010-09-04)<br>
  <span class="pDot3">H. </span>修改 初始化ID使用GUID；简化 取回密码页(3页改成一页)；(2010-09-03)<br>
  <span class="pDot3">G. </span><span class="FntF0F">筹备 php版本同步程序</span>；(2010-09-02)get_pw.asp<br>
  <span class="pDot3">F. </span>增加函数rs_Row(xcon,xSql,xLen),rs_Tab(xcon,xSql),修改函数rs_Val(xcon,xSql)；(2010-09-02)<br>
  <span class="pDot3">E. </span>修改 后台登陆 增强安全性（增加禁止软件自动填写信息,增加App26Chk超时检测）；(2010-09-01)<br>
  <span class="pDot3">D. </span>简化 前台模版，便于维护；(2010-08-31)<br>
  <span class="pDot3">C. </span>增加 <span class="FntF0F">前后台表单提交随机参数 </span>( /member/login.asp 三页面， /member/mu_app_123456.asp， /page/gbook.asp， /page/remark.asp， /bbs/badd.asp， /bbs/pkjoin.asp， /trade/gbook.asp, /smod/info/info_add.asp )， 阻止工具提交表单(如aoyou浏览器的智能添表)；(2010-08-30)<br>
  <span class="pDot3">B. </span>增加 <span class="FntF0F">前台模版</span>，对应目录为/pfile/temp，[会员，供求无此功能]；(2010-08-24)<br>
  <span class="pDot3">A. </span>增加 前台单号查询；(2010-08-24)<br>
--- 2010-08-24 Peace(XieYS) <A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF name=2.0></A> Ver2.0</P>
<p><span class="pDot3">R. </span>增加 函数rs_OrdID()，生成定单号等如：108RB-JK82K-984(YYMDH-NSXSS-999)(2010-08-20)<br>
  <span class="pDot3">Q. </span>增加 会员定单查看；修改定单参数，定单提交；(2010-08-20)<br>
  <span class="pDot3">P. </span>增加 特定参数选项[参数不是直接填写，而是用系统类别中的参数，用Option选；如产地,级别]；(2010-08-17)<br>
  <span class="pDot3">O. </span>增加 提醒功能，对应参数remind.js[提醒参数]；(2010-08-17)<br>
  <span class="pDot3">N. </span>修改 对时工具 &gt;&gt; 对时(本地时间/服务器时间)；(2010-08-16)<br>
  <span class="pDot3">M. </span>修正 产品订购 &gt;&gt; 修改订购数量时出错；(2010-08-13)<br>
  <span class="pDot3">L. </span>增加 <span class="FntF0F">生成静态页</span> 逻辑模式基本完成；(2010-08-13)<br>
  <span class="pDot3">K. </span>增加 js图片播放器(/inc/home/jsPlay02.js)，比之前jsPager.js效果更好；(2010-08-12)<br>
  <span class="pDot3">J. </span>增加 标签替换函数(func_sfile.asp:rep_TmpTags)，用于生成静态页作准备；(2010-08-11)<br>
  <span class="pDot3">I. </span>增加 工具 &gt;&gt; 对时(本地时间/服务器时间)；(2010-08-10)<br>
  <span class="pDot3">H. </span>增加 工具 &gt;&gt; 站务笔记导出；(2010-08-10)<br>
  <span class="pDot3">G. </span>增加 参数tmp_File.asp文件模板，和对应目录pfile/temf/，作为资料添加选用；(2010-08-09~2010-08-10)<br>
  <span class="pDot3">F. </span>增加 tools/out/check.asp功能，?act=Index，显示列表，信息，版本，广告，更新等；(2010-08-07)<br>
  <span class="pDot3">E. </span>完成 目录说明,文件说明,数据库说明，重要函数说明；(2010-08-07)<br>
  <span class="FntF00">发现一个问题，Access中还有查询，窗体，报表等，是否都可用在Asp上？现在确认它的查询可用！！！</span><br>
  <span class="pDot3">D. </span>增强 模板选择 功能，可增加，删除，修改模板；可选系统模板；(2010-08-06)<br>
  <span class="pDot3">C. </span>完善 开发者说明文件 - 参数显示-示例,特色应用与扩展 ；(2010-08-05)<br>
  <span class="pDot3">B. </span>简化 优化 防注配置 参数设置；(2010-08-04)<br>
  <span class="pDot3">A. </span>修正 简化函数Show_sfImgs(func_sfile.asp文件)；(2010-08-04)<br>
--- 2010-08-04 Peace(XieYS) <A href="#"><SPAN class=dCode>[TOP]</SPAN></A>
</p>
<TABLE cellSpacing=1 cellPadding=8 width=750 align=center bgColor=#999999 
border=0>
  <TBODY>
    <TR>
      <TD vAlign=top align=right bgColor=#ffffff class="dCode">更新 Peace[XieYS] 
        2010-08-02 ~ xx&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>
