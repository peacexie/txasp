<!--#include file="config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.idir {
	width:100px;
	height:18px;
	word-break:break-all;
	float:left;
	overflow:hidden;
	padding:0px 5px 0px 3px;
	margin:1px 1px;
	border-right:0px solid #F0F0F0;
}
.clear {
	clear:both;
	visibility: hidden;
}
</style>
</head>
<body>
<%

yAct = Request("yAct") 
Tim1 = Timer()
'sTab = "CExt:CDocs:CMem:CBBS;CTrade:CGboU324:SMove:SDep:MHome:MSys:MHead:MOrder"
'aTab = Split(sTab,":")

'' ///// ///// ///// ///// ///// 
If yAct="set2" Then
'' ///// ///// ///// ///// ///// 

'设定参数
pStr = "Config_Name,wmk_Mark,edrImg_Max,smaState,Config_Mode"
pTab = Split(pStr,",")
For i=0 To uBound(pTab)
  iPar = pTab(i)
  iVal = RequestS(iPar,"C",128)
  sql = "UPDATE AdmPara SET ParText='"&iVal&"' WHERE ParCode='"&iPar&"'"
  Call rs_DoSql(conn,sql)
Next
echo sql
'Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&Config_Path&"' WHERE ParCode='Config_Path'")

'清理目录
pStr = RequestF("dir","C",1200)
pTab = Split(Replace(pStr," ",""),",")
For i=0 To uBound(pTab)
  iVal = pTab(i)
  iDir = Config_Path&""&iVal&"/"
  iDir = Replace(iDir,"//","/")
  Call fold_del(iDir)
Next
echo iDir

'清理文件
pStr = RequestF("dfp","C",1200)
pTab = Split(Replace(pStr," ",""),",")
For i=0 To uBound(pTab)
  iVal = pTab(i)
  'iVal = RequestS(iPar,"C",128)
  iVal = Config_Path&""&iVal&""
  iVal = Replace(iVal,"//","/")
  Call fil_del(iVal)
Next
echo iVal

'删除资料
pStr = RequestF("del","C",1200)
pTab = Split(Replace(pStr," ",""),",")
For i=0 To uBound(pTab)
  iVal = pTab(i)
  sql = "DELETE FROM AdmSyst WHERE SysID='"&iVal&"'"
  Call rs_DoSql(conn,sql)
Next
echo sql

'删除[表](DB)
pStr = RequestF("dtb","C",1200)
pTab = Split(Replace(pStr," ",""),",")
For i=0 To uBound(pTab)
  iVal = pTab(i)
  sql = "DROP TABLE "&iVal&""
  Call rs_DoSql(conn,sql)
Next
echo sql

'开关参数:(关闭)
pStr = RequestF("dyn","C",1200)
pTab = Split(Replace(pStr," ",""),",")
For i=0 To uBound(pTab)
  iVal = pTab(i)
  sql = "UPDATE AdmPara SET ParText='N' WHERE ParCode='"&iVal&"'"
  Call rs_DoSql(conn,sql)
Next
echo sql

'保留模块

'设置参数:(2)
  
If Request("SMove")="Y" Then ''站务笔记
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysType='Home' WHERE SysID='GboU124'")
Else
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysType='Gbook' WHERE SysID='GboU124'")  
End If
If Request("MHome")="C" Then ''刷新与杂项
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID IN('HomAdv1','HomAdv2','HomAdv3','HomCount','HomLnk1','BBSVD24')")
ElseIf Request("MHome")="N" Then
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='HomAdv1'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='HomAdv2'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='HomAdv3'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='HomCount'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='HomLnk1'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='BBSVD24'") 
Else ''Y
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='3' WHERE SysID='HomAdv1'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='3' WHERE SysID='HomAdv2'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='3' WHERE SysID='HomAdv3'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='5' WHERE SysID='HomCount'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='6' WHERE SysID='HomLnk1'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='9' WHERE SysID='BBSVD24'") 
End If
If Request("MSys")="C" Then ''系统与设置
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID IN('SysTyps','SysMenu','SysConfig','SysTools','SysType')")
ElseIf Request("MHome")="N" Then
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='SysTyps'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='SysMenu'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='SysConfig'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='SysTools'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='X' WHERE SysID='SysType'") 
Else ''Y
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='6' WHERE SysID='SysTyps'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='7' WHERE SysID='SysMenu'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='9' WHERE SysID='SysConfig'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='A' WHERE SysID='SysTools'") 
  Call rs_DoSql(conn,"UPDATE AdmSyst SET SysTop='A' WHERE SysID='SysType'") 
End If
If Request("MHead")="N" Then ''头条,统计
  Call k1Set1("tmInfoEnd","li","") 
End If
If Request("MOrder")="N" Then ''定单
  Call k1Set1("tmPicsEnd","a","&nbsp;栏目设置") 
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParCode='tmPic_S999' WHERE ParCode='tmPicS999'") 
End If
If Request("SClear")="Y" Then ''数据清理
  Response.Write "<IFRAME width='98%' height='24' name=LeftMenu src='../../smod/adupd/upd_data.asp?yAct=MOnes' scrolling='no'></IFRAME>"
  Response.Write "<IFRAME width='98%' height='24' name=LeftMenu src='../../smod/adupd/upd_data.asp?yAct=DData' scrolling='no'></IFRAME>"
End If 

'' ///// ///// ///// ///// ///// 
End If
'' ///// ///// ///// ///// ///// 
'<li hdFlag><a ... hdFlag>
%>
<br>
<table width="680" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="20%" align="center" bgcolor="#FFFFFF"><strong>1Key配置</strong></td>
          <td align="center" nowrap bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <%If Config_Mode = "isExpert" Then%>
          <!--#include file="inc_menu.asp"-->
          <%End If%>
        </tr>
      </table></td>
  </tr>
  <tr align="center">
    <td align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tdFFF">
        <form name="fm1Key" id="fm1Key" method="post" action="?">
          <tr>
            <td align="left" valign="top" nowrap>设定参数:(输入)</td>
            <td align="left" valign="top" nowrap><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="20%">站点名称</td>
                <td><input id="Config_Name" maxlength="120" size="36" value="<%=Config_Name%>" name="Config_Name" />
                (Config_Name)</td>
              </tr>
              <tr>
                <td>水印控制</td>
                <td><input id="wmk_Mark" maxlength="120" size="36" value="X" name="wmk_Mark" />
                (wmk_Mark,P/T/X)</td>
              </tr>
              <tr>
                <td>编辑器-图片显示</td>
                <td><input id="edrImg_Max" maxlength="120" size="36" value="Null" name="edrImg_Max" />
                  (edrImg_Max,Width/Height/Null)</td>
              </tr>
              <tr>
                <td>短信接口</td>
                <td><input id="smaState" maxlength="120" size="36" value="Close" name="smaState" />
                  (smaState,isOpen/Close)</td>
              </tr>
              <tr>
                <td>系统管理模式</td>
                <td><input id="Config_Mode" maxlength="120" size="36" value="isExpert" name="Config_Mode" />
                  (Config_Mode,isExpert/Null)</td>
              </tr>
            </table></td>
          </tr>
          
          <tr>
            <td align="left" valign="top" nowrap>保留模块:(js组)</td>
            <td align="left" valign="top" nowrap>
              <div class='idir'>
                <input type='checkbox' name='dun' id='dun' value='doc' onClick="setNull(this)">
                内部公文</div>
              <div class='idir'>
                <input type='checkbox' name='dun' id='dun' value='member' onClick="setNull(this)">
                会员与查询</div>
              <div class='idir'>
                <input type='checkbox' name='dun' id='dun' value='bbs' onClick="setNull(this)">
                论坛与投票</div>
              <div class='idir'>
                <input type='checkbox' name='dun' id='dun' value='trade' onClick="setNull(this)">
                商务与供求</div> 
              <div class='idir'>
                <input type='checkbox' name='dun' id='dun' value='sms' onClick="setNull(this)">
                短信</div> 
            </td>
          </tr>
          
          <tr>
            <td align="left" valign="top" nowrap>清理目录:(组件)</td>
            <td align="left" valign="top" id="jdir"><%
t = ""
t = t&"/bbs,/doc,/member,/vote||,"
t = t&",/msg,/trade||,"
t = t&"/ext/code*,/ext/xtemp*,/ext/xtest,/ext/Pea测试||," '/ext/apis,
t = t&"/ext/api/cset,/ext/api/iis,/ext/api/scan*,/ext/api/rss,/ext/api/seo||,"
t = t&"/ext/api/eplay,/ext/api/scroll4,/ext/api/menu*,/ext/api/mzoom*,/ext/api/map||,"
t = t&"/upfile/dtbbs,/upfile/dtbus,/upfile/dtdoc,/upfile/tools||,"
t = t&"/setup,/sadm/setup,/img/Head,/img/tgif,"
a = Split(t,",")
For i = 0 To uBound(a)
  If a(i)<>"" Then
    v = a(i) :s = v
	If inStr(v,"*")>0 Then
	  v = Replace(v,"*","")
	  c = " xcheck"
	Else
	  c = " checked"
	End If
	v = Replace(v,",","")
	s = Replace(s,",","")
	v = Replace(Replace(v,",",""),"||","")
	d = "" :If NOT fold_exist(Config_Path,v) Then d = " disabled"
	Response.Write vbcrlf&"<div class='idir'><input type='checkbox' name='dir' id='dir' value='"&v&"' "&c&d&" />"&s&"</div>"
	If inStr(s,"||")>0 Then
	  Response.Write vbcrlf&"<div class='clear'></div>"
	End If
 End If
Next
		%></td>
          </tr>
          <tr>
            <td align="left" valign="top" nowrap>清理文件:(但个)</td>
            <td align="left" valign="top" nowrap id="jdfp"><%
t = ""
t = t&"/ext/ext_side.asp,/ext/go.asp,/ext/login.asp,/ext/mail.asp,/ext/order.asp,/ext/Pea测试.asp,"
a = Split(t,",")
For i = 0 To uBound(a)
  If a(i)<>"" Then
    v = a(i) :s = v
	If inStr(v,"*")>0 Then
	  v = Replace(v,"*","")
	  c = " xcheck"
	Else
	  c = " checked"
	End If
	v = Replace(v,",","")
	s = Replace(s,",","")
	d = "" :If NOT fil_exist(Config_Path&v) Then d = " disabled"
	Response.Write vbcrlf&"<div class='idir'><input type='checkbox' name='dfp' id='dfp' value='"&v&"' "&c&d&">"&s&"</div>"
  End If
Next
		%></td>
          </tr>
          <tr>
            <td align="left" valign="top" nowrap>删除资料:(数据)</td>
            <td align="left" valign="top" nowrap id="jdel"><%
t = ""
t = t&"ModMember,会员与查询;ModBBS,论坛与投票;ModTrade,商务与供求;ModDocs,内部公文;ModSms,短信系统;"
t = t&"ModHost,主机管理;Peace_Test,Peace测试||;"
t = t&"InfD124,部门介绍*;InfC124,课程介绍;InfA124,企业介绍;InfN224,新闻(英)||;"
t = t&"PicR124,人物图片;PicT124,其它图片*;PicV124,视频播放;PicV125,文件下载;PicTBak,备用图片||"
a = Split(t,";")
For i = 0 To uBound(a)
  If a(i)<>"" Then
    v = a(i) :p = inStr(v,",")
	If inStr(v,"*")>0 Then
	  c = " xcheck"
	Else
	  c = " checked"
	End If
	s = Mid(v,p) :v = Left(v,p)
	v = Replace(v,",","")
	s = Replace(s,",","")
	d = "" :If rs_Exist(conn,"SELECT * FROM AdmSyst WHERE SysID='"&v&"'")="EOF" Then d = " disabled"
	Response.Write vbcrlf&"<div class='idir'><input type='checkbox' name='del' id='del' value='"&v&"' "&c&d&">"&s&"</div>"
	If inStr(s,"||")>0 Then
	  Response.Write vbcrlf&"<div class='clear'></div>"
	End If
  End If
Next
%></td>
          </tr>
          <tr>
            <td align="left" valign="top" nowrap>删除[表]:(DB)</td>
            <td align="left" valign="top" nowrap id="jdtb"><%
t = "" :sTabs = t1List()
t = t&"DocsLogs,DocsNews,DocsRemark||,"
t = t&"MemCard,MemSyst,Peace测试||,"
t = t&"BBSInfo,BPKTitle,BPKView,BPKVote||,"
t = t&"TradeCorp,TradeGbook,TradeInfo,TradePara,TradeType||,"
t = t&"SmsCharge,SmsMember,SmsSend,SmsTelq,SmsTels,SmsTemp,SmsType||,"
a = Split(t,",")
For i = 0 To uBound(a)
  If a(i)<>"" Then
    v = a(i) :s = v
	If inStr(v,"*")>0 Then
	  c = " xcheck"
	Else
	  c = " checked"
	End If
	v = Replace(Replace(v,",",""),"||","")
	s = Replace(s,",","")
	'rs_ETab(conn,v) 很慢(=1s),以下快(=0.1s)
	d = "" :If inStr(sTabs,v)<=0 Then d = " disabled"
	Response.Write vbcrlf&"<div class='idir'><input type='checkbox' name='dtb' id='dtb' value='"&v&"' "&c&d&">"&s&"</div>"
	If inStr(s,"||")>0 Then
	  Response.Write vbcrlf&"<div class='clear'></div>"
	End If
  End If
Next
%></td>
          </tr>
          <tr>
            <td align="left" valign="top" nowrap>开关参数:(关闭)</td>
            <td align="left" valign="top" nowrap id="jdyn"><%
t = ""
t = t&"SwhModGroup,管理菜单分组;SwhRemSave,保存远程图片;SwhGuestPub,Guest发布;SwhMemApp,允许会员注册;SwhDepSubs,分公司(部门)"
a = Split(t,";")
For i = 0 To uBound(a)
  If a(i)<>"" Then
    v = a(i) :p = inStr(v,",")
	If inStr(v,"*")>0 Then
	  c = "xcheck"
	Else
	  c = "checked"
	End If
	s = Mid(v,p) :v = Left(v,p)
	v = Replace(v,",","")
	s = Replace(s,",","")
	Response.Write vbcrlf&"<div class='idir'><input type='checkbox' name='dyn' id='dyn' value='"&v&"' "&c&">"&s&"</div>"
	If inStr(s,"||")>0 Then
	  Response.Write vbcrlf&"<div class='clear'></div>"
	End If
  End If
Next
%></td>
          </tr>
          <tr>
            <td align="left" valign="top" nowrap>设置参数:(选择)</td>
            <td align="left" valign="top" nowrap><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="20%">站务笔记(移动)</td>
                <td><select name="SMove" size="1" id="SMove">
                  <option value="Y">Y.移动</option>
                  <option value="N" selected>N.否</option>
                </select>
                  留言与笔记 --- 刷新与杂项</td>
              </tr>
              <tr>
                <td>刷新与杂项</td>
                <td><select name="MHome" size="1" id="MHome">
                  <option value="Y">Y.显示</option>
                  <option value="C">C.清除</option>
                  <option value="N" selected>N.隐藏</option>
                </select>
                  浮动广告,图片广告,文字广告,统计计数,友情连接,网上调查</td>
              </tr>
              <tr>
                <td>系统与设置</td>
                <td><select name="MSys" size="1" id="MSys">
                  <option value="Y">Y.显示</option>
                  <option value="C">C.清除</option>
                  <option value="N" selected>N.隐藏</option>
                </select>
                  信息类别,系统菜单,系统配置,系统工具,系统类别</td>
              </tr>
              <tr>
                <td>头条,统计</td>
                <td><select name="MHead" size="1" id="MHead">
                  <option value="Y" selected>Y.要</option>
                  <option value="N">N.不要</option>
                </select>
                  头条设置 信息发布统计</td>
              </tr>
              <tr>
                <td>定单</td>
                <td><select name="MOrder" size="1" id="MOrder">
                  <option value="Y">Y.要</option>
                  <option value="N" selected>N.不要</option>
                </select>
                  定单参数 --- 定单管理</td>
              </tr>
              <tr>
                <td>数据清理</td>
                <td><select name="SClear" size="1" id="SClear">
                  <option value="Y" selected>Y.要</option>
                  <option value="N">N.不要</option>
                </select>
                  清理:个人笔记,DB数据</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align="left"><input name="yAct" type="hidden" id="yAct" value="set2">
            <a href="?">刷新</a><%=""&Timer()-Tim1&""%></td>
            <td align="left"><table width="100%" border="0" cellpadding="0" cellspacing="2">
              <tr>
                <td align="center"><input type="submit" name="button2" id="button2" value="一键1Key配置"></td>
                <td align="center">注意：清理目录,清理文件,删除资料,删除[表]操作不可逆！</td>
              </tr>
            </table></td>
          </tr>
        </form>
    </table></td>
  </tr>
</table>
<script type="text/javascript">
function setNull(e){
  var c = e.checked;
  var v = e.value;
  if(v=="doc"){
	sdir = "/doc";  sdfp = "/ext/login.asp"; 
	sdel = "ModDocs"; sdtb = "Docs";
  }
  if(v=="member"){
	sdir = "/member";  sdfp = ""; 
	sdel = "ModMember"; sdtb = "Mem"; 
  }
  if(v=="bbs"){
	sdir = "/bbs";  sdfp = ""; 
	sdel = "ModBBS"; sdtb = "BBS,BPK";  
  }
  if(v=="trade"){
	sdir = "/trade";  sdfp = ""; 
	sdel = "ModTrade"; sdtb = "Trade";  
  }
  if(v=="sms"){
	sdir = "/msg";  sdfp = ""; 
	sdel = "ModSMS"; sdtb = "Sms";    
  }
  v = true; if(e.checked==true) v=false;
  
  a = "dir,dfp,del".split(","); 
  for(k=0;k<a.length;k++){
	aVal = a[k]; //'dir'
	aStr = eval("s"+aVal) ; //sdir
	jID = "j"+aVal ; //'jdir'
	var e2 = document.getElementById(jID).getElementsByTagName("input");
	for(var j=0;j<e2.length;j++){ 
	  v2 = e2[j].value; 
	  if(aStr.indexOf(v2)>=0)
	  e2[j].checked = v; 
	}  
  }

  a = sdtb.split(",");
  for(i=0;i<a.length;i++){
	var e2 = document.getElementById('jdtb').getElementsByTagName("input");
	for(var j=0;j<e2.length;j++){ 
	  v2 = e2[j].value; 
	  if(v2.indexOf(a[i])>=0)
	  e2[j].checked = v; 
	}
  }

}



</script>
<%

Function k1Set1(xID,xFlag,xStr) '"tmInfoEnd","li",""
  SET rs=Server.CreateObject("Adodb.Recordset") 
  s = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='"&xID&"'")
  f1 = "<"&xFlag      : n1 = inStr(s,f1)
  f2 = "</"&xFlag&">" : n2 = inStr(s,f2)+Len(f2)
  SET rs=Nothing
  If n1>0 And n2>n1 Then
  t = Mid(s,n1,n2-n1)
  s = Replace(s,t,xStr) : s = Replace(s,"'","")
  'Response.Write s
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParRem='"&s&"' WHERE ParCode='"&xID&"'")
  End If
End Function

Function t1List() 
  Dim sql,iTab,listTab,rs : listTab=""
  If cfgDBType="Access" Then
	Set cnAcc = Server.CreateObject("ADODB.Connection")
	cnAcc.Open conn   
	Set rsAcc = cnAcc.OpenSchema(20) 'adSchemaTablesA
	Do while not rsAcc.EOF
	  iTab = rsAcc("TABLE_NAME")
	  listTab = listTab&iTab&";"
	 rsAcc.MoveNext
	Loop
	rsAcc.Close
	cnAcc.Close
  Else 'MsSql
    sql = " SELECT [name],[uid] FROM [sysobjects] WHERE (xtype = 'U') ORDER BY [xtype] DESC,[name]  " 
	Set rs = Server.Createobject("Adodb.Recordset")
	rs.Open Sql,conn,1,1
	Do While NOT rs.EOF 
	  iTab = rs("name")
	  listTab = listTab&iTab&";"
	rs.moveNext
	Loop
	rs.Close()
	Set rs = Nothing
  End If
  t1List = listTab
End Function
'echo t1List()
%>
</body>
</html>
