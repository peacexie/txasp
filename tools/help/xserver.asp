<!--#include file="../himg/tconfig.asp"-->
<%Call Chk_Perm1(xPara,"")%>
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
      <TD vAlign=top align=center bgColor=#ffffff class="dCode"><STRONG class="fSiz14"><%=Config_Name%> - 服务器设置参考文件</STRONG></TD>
    </TR>
    <TR>
      <TD vAlign=top bgColor=#ffffff><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          <td valign="top" bgcolor="#FFFFFF">　　
            如果你仅管理内容，此部分可以不用管；租用虚拟主机的，此部分由空间供应商设置；当服务器重新配置等，可把此文件提供给配置者(网管)参考。<br>
            -=&gt; <a href="../_admin/demo-fck.asp" target="_blank">编辑器基本属性使用说明 (To网站资料管理员)</a><br>
            -=&gt; <a href="xhelp.asp" target="_self">后台管理帮助文件 (返回.....)</a></td>
          <td width="30%" valign="top" nowrap bgcolor="#FFFFFF" style="line-height:130%; border:1px solid">&middot; <a href="#FlagFirst">概述和说明</a><br>
            &middot; <a href="#FlagServer">服务器常用安全配置</a><br>
            &middot; <a href="#FlagSite">站点常用安全配置</a><br>
            &middot; <a href="#FlagMIME">常见文件MIME类型</a><br>
            &middot; <a href="#FlagRel">网站服务器漏洞可相互影响</a> <br></td>
        </tr>
      </table></TD>
    </TR>
  </TBODY>
</TABLE>
<P class=dTitle><A id=FlagUTF4 name=FlagFirst></A> 概述和说明：</P>
<P><span class="pDot2">1. </span>客户自己管理的服务器，请自行处理，<span class="FntF00">这里只作提示，不做具体培训</span>！<br>
<span class="pDot2">2. </span>这里<span class="FntF00">仅一些基本配置，更多高深做法，请查看更专业的资料</span>！<br>
<span class="pDot2">3. </span>本文<span class="FntF00">很多设置，其实跟本不需要，系统也可正常运行，但是比较容易受攻击而已</span>！！！<br>
<span class="pDot2">4. </span>如果你是服务器相关管理员，而你又不知道以上是怎么回事，那赶快 查资料，或请教懂的人，或另请人来管理服务器！！！！！！！！！！<br>
<span class="pDot2">5. </span>本系统运行在Win2003+IIS下；起用FSO,Adodb.Stream自带组件；根据需要，安装JMail,ASPJpeg组件；<br>
注意，JMail,ASPJpeg组件所在的目录，需要Users读取权限，如可以直接设置.dll所在的目录或文件的Users读取权限：<br>
D:\Program Files\Persits Software\AspJpeg\Bin<br>
D:\Program Files\Dimac\w3JMail4\jmail.dll ！ <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF name=FlagServer></A> 服务器常用安全配置：</P>
<P><span class="pDot2">1. </span>NTFS权限<br>
系统盘和站点放置盘必须设置为NTFS格式,方便设置权限；<br>
系统盘和站点放置盘除administrators 和system的用户权限全部去除；<br>
  <span class="pDot2">2. </span>端口 和 防火墙<br>
  启用windows自带防火墙,只保留有用的端口,比如远程和Web,Ftp(3389,80,21)等等,有邮件服务器的还要打开25和130端口改变远程登陆，FTP的默认端口；<br>
<span class="pDot2">3. </span>帐号 和 密码：<br>
  MS SQL Server，更改sa密码为你都不知道的超长密码,在任何情况下都不要用sa这个帐户.<br>
  禁用 系统administrator 帐号，或设置一个你自己也不知道的超长密码，作为馅井；<br>
  设置ServU密码，<br>
ServU密码和系统登陆密码 都使用强密码；<br>
<span class="pDot2">4. </span>卸载 W.Script.S.hell, 去除HappyTime(欢乐时光)威胁：把以下代码存成.bat文件；运行 <br>
<span class="Fnt999">reg delete &quot;HKEY_CLASSES_ROOT\CLSID\{06290BD5-48AA-11D2-8432-006008C3FBFC}&quot; /f<br>
reg delete &quot;HKEY_CLASSES_ROOT\Scriptlet.TypeLib&quot; /f<br>
reg delete &quot;HKCR\CLSID\{06290BD5-48AA-11D2-8432-006008C3FBFC}&quot; /f<br>
reg delete &quot;HKCR\Scriptlet.TypeLib&quot; /f<br>
regsvr32 /u C:\WINDOWS\System32\wshom.ocx<br>
  rem --del C:\WINDOWS\System32\wshom.ocx<br>
  regsvr32 /u C:\WINDOWS\system32\sh<span></span>ell32.dll<br>
rem --del C:\WINDOWS\system32\sh<span></span>ell32.dll</span><br>
<span class="pDot2">5. </span>让 IIS 上传&gt;200K <br>
  1. 停止IIS --- net stop iisadmin /y<br>
  2. 打开 (Driver)\Windows\system32\inesrv\metabase.xml<br>
  3. 修改 ASPMaxRequestEntityAllowed 的值为自己需要的, 默认为 204800 如改为24123000<br>
4. 启动IIS --- net start iisadmin<br>
<span class="pDot2">6. </span>让 IIS可以访问 .stm 静态文件：<br>
IIS &gt;&gt; Web服务扩展 &gt;&gt; 在服务器端的包含文件 &gt;&gt; (设置)允许<br>
<span class="pDot2">7. </span>让 IIS可以 下载/打开 flv等媒体文件：<br>
IIS &gt;&gt; 站点 &gt;&gt; 属性 &gt;&gt; http头 &gt;&gt; mime类型 &gt;&gt; 增加 &gt;&gt; (.flv等类型文件) <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF2 name=FlagSite></A> 站点常用安全配置：</P>
<P><span class="pDot2">1. </span>增加一个新用户组(如iis_guests)；为每个站点建立一个独立的系统用户,设置为属于 iis_guests组，去掉远程控制的勾；每个站点 都用以上设置的 用户运行；<br>
  <span class="pDot2">2. </span>设置站点目录的写权限<br>
  script/， upfile/  或 web/ (如果有的话)<br>
<span class="pDot2">3. </span>去掉以下目录的“执行”权限：<br>
  script/ 或 upfile/<br>
IIS站点 &gt;&gt; 目录 &gt;&gt; 右键 &gt;&gt; 属性 &gt;&gt; 执行权限 &gt;&gt; 设置无<br>
<span class="pDot2">4. </span>blog 写权限： (如果有的话)<br>
  \blog\DataSource(data)<br>
  \blog\GG(upadv)<br>
  \blog\u(u*)<br>
  \blog\UploadFiles(upfile) (去掉此目录的“执行”权限)<br>
\blog\XmlData()<br>
<span class="pDot2">5. </span>blog子域名 跳转：<br>
  1. 解析泛域名；<br>
2. 设置一个站点，空主机头；指向 (Root)\web\(sys)dir 目录！ <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
<P class=dTitle><A id=FlagUTF3 name=FlagMIME></A> 常见文件MIME类型：</P>
<p>/* ============================================================<br>
  常见文件MIME类型(http中content-type头值) ，先查看：系统默认设置：IIS管理器 &gt;&gt; 右键 &gt;&gt; 属性 &gt;&gt; MIME类型 &gt;&gt; <br>
  系统默认设置默认设置的，就可不用设置了，没有设置的，根据需要设置，否则可能打不开相关文件。如：<br>
  .flv - application/x-shockwave-flash (根据需要设置)<br>
  .Peace!Bak - application/zip (为安全起见，一般不要设置)<br>
.Peace!DB - application/zip (为安全起见，一般不要设置)<br>
更多，请看系统默认设置。 <A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagUTF5 name=FlagRel></A> 网站服务器漏洞可相互影响：</P>
<p>网上一找，68元/168元建网站一大把；域名1元/年一大把！<br>
  关于域名1元/年：一年之后还能找得到续费的人或单位吗？一年之后续费还是1元一年吗？<br>
关于 168元建网站：请看下面“本系统还值什么 --- 发现我们的价值！”另，他们建的网站能同时达到以下多少个亮点，他们建的网站还有那些其它亮点(谢谢推荐)；另建成的网站,程序安全稳定就好，有这方面有保障吗？关于网站安全的安全性和稳定性请看以下说明：<br>
<span class="pDot3">网站，服务器 漏洞可相互影响：</span>最好是 好代码（网站） + 好服务器(含服务管理)<br>
这就是 为什么 有的做同一个网站 可以报价 68,168,1680 .... 同一个域名(空间) 可以是1元/年，200元/年...<br>
结论是：（除了无理要价）有时做网站，管服务器 成本确实高，而且确实值很高的价格<br>
<span class="pDot3">网站 与 服务器 </span><br>
服务器 管理好了， 即使差的网站代码 不会影响 别的网站；<br>
服务器 没有管理好， 一个差的网站代码 可能影响 别的网站，甚至影响 整个服务器；<br>
服务器安全威胁，网站只是一方面，还有 ServU,FTP (这里不多说)等；<br>
<span class="pDot3">我做的网站 </span>7~8年以来，一直在尽最大的努力，不断的做安全处理；<br>
从网站单方面看，目前安全性我有很高的自信！即使这样，还是没有终点！<br>
但有几点除外：<br>
a. 泄露后台管理密码，或可以猜测的简单管理密码，或操作不当<br>
b. 泄露 相关FTP 密码等<br>
c. 服务器 跨站攻击，从别的网站入侵到我做的网站:我无能为力！<A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></p>
<P class=dTitle><A id=FlagUTF8 name=FlagValue></A> 本系统还值什么 --- 发现我们的价值！</P>
<P><span class="pDot2">1. </span>新闻产品信息添加/管理时 的一些亮点：<br>
  <span class="pDot3">标题颜色：</span><br>
  <span class="pDot3">模板选择：</span><br>
  <span class="pDot3">编辑器：</span>(兼容IE6,IE7,IE8,FF,结合FCK/eWeb两大编辑器优势)<br>
  <span class="pDot3">关键字：</span>做优化设置(自动获取)<br>
  <span class="pDot3">属性设置：</span>需要时使用<br>
  <span class="pDot3">添加资料后返回选项：</span>添加资料后可选 返回列表/继续添加 选项，提供效率<br>
  <span class="pDot3">高级设置：</span>(信息复制等快捷设置)<br>
  <span class="pDot3">批量上传：</span>思想来源于宝宝树的照片批量上传，可提供效率<br>
  <span class="pDot3">相关图片：</span>需要时使用，可批量上传<br>
  <span class="pDot3">信息采集：</span>(专业人士使用)<br>
  <span class="pDot3">删除垃圾附件：</span>[删除资料时，自动删除垃圾附件，很多系统不会的哦]<br>
  <span class="pDot3">定制信息属性：</span>(如根据需要设置产品规格,价格,产地等最多90多个参数)<br>
  <span class="pDot2">2. </span>新闻产品信息前台显示 的一些亮点：<br>
  <span class="pDot3">分页测试：</span>http://peace.96327.com/page/iview.asp?KeyID=dtinf-2010-78-B6WV.5EWMM<br>
  <span class="pDot3">读后感受投票：</span><br>
  <span class="pDot3">评论：</span><br>
  <span class="pDot3">综合介绍：</span>http://peace.96327.com/page/info.asp?ModID=InfA124，点右边连接，看各种显示方式；<br>
  <span class="pDot2">3. </span>网站系统 亮点<br>
  <span class="pDot3">空间占用情况：</span>提示系统空间占用情况<br>
  <span class="pDot3">站务笔记：</span><br>
  <span class="pDot3">管理工具管理帮助：</span><br>
  <span class="pDot3">权限分组：</span><br>
  <span class="pDot3">图片同步（导入）：</span>用于多语言版，提供效率<br>
  <span class="pDot3">简体/繁體转换：</span><br>
  <span class="pDot3">安全稳定/防注入/防注册机：</span>个人觉得[<span class="FntF0F">行业领先</span>]<br>
  <span class="pDot3">参数设置：</span>根据需要，可设置“配置 | 随机 | 数字 | 开关 | Jmail | Editor | SMS | 水印 | 日期 | 联系”等参数<br>
  <span class="pDot2">4. </span>可选模块和功能。<br>
  <span class="pDot3">内部公文：</span>塘厦教育/长安教育<br>
  <span class="pDot3">生成静态：</span> <br>
  <span class="pDot3">短信系统：</span>塘厦教育/sms.dg.gd.cn<br>
  <span class="pDot3">会员与查询：</span>见 <a href="works.htm" target="_blank">我们做过的一些典型网站系统</a><br>
  <span class="pDot3">论坛与投票：</span>http://peace.96327.com/vote/vote.asp?ID=0BA6A13A237FBM1D7025WDP1<br>
  <span class="pDot3">图片水印：</span>图片和文字水印，文字水印可描边[<span class="FntF0F">行业领先</span>]<br>
  <span class="pDot3">广告设置：</span>(如浮动客服QQ,MSN,Skype)<br>
  注意，以上提及的一些功能，有些按需配置，不需要的可能看不到，有些功能请额外下单并收费。<A 
href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></P>
</body>
</html>


