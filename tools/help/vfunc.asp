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
              <span class="pILink">重要函数说明</span>
              <a href="#FuncTime" class="pILink">func_time.asp</a>
              <a href="#FuncRS" class="pILink">func_rs.asp</a>
            <a href="#FuncVBS" class="pILink">func_vbs.asp</a><a href="#FuncOpt" class="pILink">func_opt.asp</a><a href="#FuncSFile" class="pILink">func_sfile.asp</a><a href="#FuncFunc3" class="pILink">func_func3.asp</a></td>
            <td width="25%" valign="top" bgcolor="#FFFFFF" style="line-height:130%; border:1px solid"><!--#include file="vmenu.asp"--></td>
          </tr>
      </table></TD>
    </TR>
  </TBODY>
</TABLE>

<P class=dTitle><A id=FlagUTF5 name=FlagTopsN></A> 重要函数说明</P>
<table width="750" border="0" align="center" cellpadding="8" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td align="left" bgcolor="#FFFFFF"><ul>
      <li class="Fnt00F">func_time.asp<span class="dTitle"><A id=FlagUTF4 name=FuncTime></A></span></li>
      <li>Get_GUID(sVer,sAddr) 得到36位GUID如：0BA8973D-4567-FA63-0FA8-7F000001FD9B</li>
      <li>Get_AppDay() 会员申请时间限制函数</li>
      <li>Get_yyyymmdd(xDate) 得到8位YYYYMMDD格式日期</li>
      <li>Get_hhmmss() 得到6位HHMMSS格式时间</li>
      <li>Get_mSec() 得到3位毫秒数</li>
      <li>Fmt_Time(xStr,xType) 格式化时间</li>
      <li>Get_9999ID(xStr,xLen) 用0补齐xLen位ID数字</li>
      <li>Get_AutoID(xN) 得到xN位不重复ID，xN&gt;=12</li>
      <li>Rnd_ID(xType,xLen2) 得到xLen位随机ID；xType:N:数字，KEY:数字+大写字母（不含易混字母：Ii  Ll  Oo  Zz）</li>
      <li>Next_ID(xOld,xMin,xStep) 得到下一个ID，自动根据规则....</li>
      <li>Base_32(xNum,xBase,xLen,xDir) 得到xLen位32进制数</li>
      <li><A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></li>
      <li class="Fnt00F">func_rs.asp<span class="dTitle"><A id=FlagUTF name=FuncRS></A></span>（以下，xConn表示数据库连接字串；xSql表示sql语句）</li>
      <li>rs_SPPage(xConn,xTabCols,xTabWhere,xPagNow,xPagSize,xKeyID,xKeySort,xRSPage) 分页存储过程（仅MSSQL）</li>
      <li>Add_Log(xconn,xUsrID,xAct,xSys,xNote) 增加一条系统记录</li>
      <li>rs_Exist(xconn,xSql) 是否存在</li>
      <li>rs_Count(xconn,xTab) 记录条数</li>
      <li>rs_Val(xconn,xSql) 查询一个值</li>
      <li>rs_AutID(xconn,xTab,xCol,xModPath,xType,xTime)  得到自动ID (dtdef-2010-87-C849.3W8XS)</li>
      <li>rs_TypID(xconn,xTab,xCol,xStr,xLen) 得到自动ID (类别表.WebTyps等)</li>
      <li>rs_DoSql(xconn,xSql) 执行sql语句</li>
      <li><A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></li>
      <li class="Fnt00F">func_vbs.asp<span class="dTitle"><A id=FlagUTF2 name=FuncVBS></A></span></li>
      <li>Get_BSize(xByte) 得到大小(KB/MB/GB)</li>
      <li>Get_CIP() 得到用户IP</li>
      <li>Get_State(xState,xKey,xMsg) 得到带颜色的状态</li>
      <li>Get_Option(xmid,xfirst,xend,xstep) 得到数字option选项</li>
      <li>Get_SOpt(xCode,xName,xVal,xFlag) 得到特定字串的option选项</li>
      <li>Show_RExp(xStr,xPatrn,xObj) 通用正则表达式Replace</li>
      <li>Show_sTitle(xText,xColor) 显示带颜色标题</li>
      <li>Show_SLen(xStr,xLen) 显示指定长度字串,英文算半个长度</li>
      <li>Show_Html(xHtml) 过滤html中的不安全标记</li>
      <li>Show_HView(xText) 显示UBB标记</li>
      <li>Show_HText(xStr,xLen) 提取HTML中的纯文本</li>
      <li>Show_Text(xText) 显示文本字串，处理&lt;,&gt;,',&quot;等[Tab]</li>
      <li>Show_Form(xText) 显示表单项内容，处理'&quot;等</li>
      <li>Show_jsStr(xText) 显示js字串</li>
      <li>Get_vPath(xLen) 得到文件路径</li>
      <li>Get_fName() 得到文件名</li>
      <li>RequestF(xPName,xPType,xLen) 获得表单数据，并检查如：<br>
        <span class="Fnt999">UserID=RequestF(&quot;uid&quot;,&quot;C&quot;,12): 获得表单id数据，最多12个字符; </span></li>
      <li>RequestQ(xPName,xPType,xDefault) 获得地址栏数据，并检查如：<br>
        <span class="Fnt999">kID=RequestQ(&quot;kid&quot;,&quot;N&quot;,0): 获得地址栏id，如果不是数字，则默认为0; </span></li>
      <li>RequestS(xPName,xPType,xDefault) 获得表单或地址栏数据，并检查如：<br>
        <span class="Fnt999">uBirth=RequestS(&quot;Birth&quot;,&quot;D&quot;,&quot;1900-12-31&quot;): 获得表单或地址栏的Birth数据，如不是日期，则默认为1900-12-31; </span></li>
      <li>RequestSafe(xPName,xPType,xDefault) 检查数据，xPType：N:数字，D:日期/时间，其它字符则处理撇号；<br>
        <span class="Fnt00F">以上</span><span class="FntF0F">4个Request*函数</span><span class="Fnt00F">，解决处理了所有sql注入相关问题！！！</span></li>
      <li>js_Alert(xMsg,xAct,xAddr) 显示js提示框</li>
      <li>Chr_Filter(xObjStr) 过滤关键字，返回关键字 或 空</li>
      <li>Chr_Fil2(xObjStr) 过滤关键字，*代替关键字      
      </li>
      <li><A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></li>
      <li><span class="Fnt00F">func_opt.asp</span><span class="dTitle"><A id=FlagUTF3 name=FuncOpt></A></span></li>
      <li>Get_TypeLable(xModID) 得到opt标签</li>
      <li>Get_TypeName(xModID,xType) 得到opt对应的TypName</li>
      <li>Get_TypeOpt(xModID,xType) 得到opt</li>
      <li>Get_Typ2Name(xModID,xTyp2) 得到次类别名称</li>
      <li>Get_Typ2Opt(xModID,xTyp2) 得到次类别opt</li>
      <li>Get_rsOpt(xconn,xSql,xDef) 得到opt[=]</li>
      <li>Get_rsOpt2(xconn,xSql,xDef) 得到opt[inStr()]</li>
      <li>Get_rsCBox(xconn,xSql,xDef) 得到checkbox</li>
      <li>Get_rsTree(xconn,xSql,xDef,xStyle) 得到树型目录</li>
      <li><A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></li>
      <li><span class="Fnt00F">func_sfile.asp</span><span class="dTitle"><A id=FlagUTF6 name=FuncSFile></A></span></li>
      <li>get_1Img(xID,xImg) 得到第一个图的url</li>
      <li>Show_sfImgs(xImgName,xID) 得到img数组</li>
      <li>Show_sfRead(xID,xFile) 读取文件</li>
      <li>Show_sfGbook(xID,xFile) 显示gbook...文件</li>
      <li>Show_sfData(xID,xFile) 显示news,pics...文件</li>
      <li>Check_sfData(xPath,xFTab,xMsg) 检查文件是否存在</li>
      <li>rel_ModTab(xMod) 从模块号得到表名</li>
      <li>rel_TabPath(xTab) 从表名得到路径</li>
      <li>rel_IDTab(xID) 从ID得到表名</li>
      <li>add_sfFile() <span class="Fnt00F">Sub</span> 内容写文件</li>
      <li>del_sfDir(xTab,sID) 删除  News,Pics,Vdos,Jobs 资料</li>
      <li>del_sfCont(xTab,sID) 删除  Gbook,Reply,GboSend 资料</li>
      <li>get_TmpLink() <span class="Fnt00F">Sub</span> 得到连接URL</li>
      <li>get_TmpID(xTab,xTyps) 得到模板（模式）ID</li>
      <li>get_TmpCont(xCode) 得到内容模板</li>
      <li><A href="#"><SPAN 
class=dCode>[TOP]</SPAN></A></li>
      <li>\inc\home\func3.asp<span class="dTitle"><A id=FlagUTF7 name=FuncFunc3></A></span><br>
      </li>
      <li>GetItemPart(xMD,xTM) 得到类别列表的一部分 含 xTM;</li>
      <li>GetItemTree(xMD,xUrl,xTM) 得到 折叠菜单</li>
      <li>GetItemLays(xMD,xUrl,xTM)得到 分级 类别列表（多级）</li>
      <li>GetItemList(xMD,xUrl,xTM)得到 类别列表（含Flag标记:Sin,Mul,Page,List,Pics,Link,LInn）</li>
      <li>GetMName(xMD) '得到 模块名称 ModID:名称</li>
      <li>GetTName(xMD,xID,xFlag) '得到 类别名称 或 Lay字串 TypID:名称,S110048;S120112;</li>
      <li>GetNLay(xMD,xID,xLink,xStr1,xStr2) 得到 类别名称 含级别：衣(服饰)系列 &gt;&gt; 童装 &gt;&gt; </li>
      <li>ListLink(xType,xTemp) 得到 (图片)连接/////////////////////</li>
      <li>ListPNext(xTab,xMod,xTyp,xTim,xWhr) 得到 上下页</li>
      <li>ListGTemp(xType,xID,xPath2)得到 摸版内容：xType:Code,File,Para; </li>
      <li>ListTemp(xTemp,xLine) 得到 摸版列表，用于测试</li>
      <li>ListPub(xTemp,xLen,xSql)得到 资料列表，新闻，图片 </li>
    </ul></td>
  </tr>
</table>

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
