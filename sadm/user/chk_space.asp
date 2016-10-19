<!-- check space 1 -->
<%

runTim01 = Timer()
edNameID = rs_Val(conn,"SELECT ParText FROM AdmPara WHERE ParCode='edNameID'")

Function isSameDir(xStr)
  rDir = Server.MapPath("../../"&xStr)
  vDir = Server.MapPath(Config_Path&xStr)
  'Response.Write "<br>r:"&rDir
  'Response.Write "<br>v:"&vDir
  If rDir=vDir Then
    isSameDir = true
  Else
    isSameDir = false
  End If
End Function


'root
'vdir(upfile,img,edfck,blog~dbf,~error)
'sdir(子域名的子目录)
'log(~ftplog,~weblog)
'note


' Config 
Dim vDirStr,aDirTab,nDirSize : nDirSize=0
Dim aDirSize(24),qDirSize(24)
vDirStr = "" 'blog;bbs; --- img;sadm/edfck
aDirTab = Split(vDirStr,";")
For i=0 To uBound(aDirTab)
  aDirSize(i) = fold_Size(Config_Path&""&aDirTab(i)&"/")
  nDirSize = nDirSize+aDirSize(i)
Next




spCore  = fold_Size(Config_Path)         ':Response.Write "<br>spCore:"&spCore
spBase  = spCore
spImg   = fold_Size(Config_Path&"img/")  ':Response.Write "<br>spImg:"&spImg
spEdfck = fold_Size(Config_Path&"sadm/"&edNameID&"/")  ':Response.Write "<br>spEdfck:"&spEdfck
spUp    = fold_Size(Config_Path&"upfile/") ':Response.Write "<br>spUp:"&spUp
If(fold_exist(Config_Path,"~dbf")) Then
  spDB  = fold_Size(Config_Path&"~dbf/") 
  spDBf = "vdir"
Else
  spDB  = fold_Size(Config_Path&"upfile/#dbf#/")
  spDBf = "local"
End If
  'Response.Write "<br>spDB:"&spDB


If isSameDir("img/") Then
  spCore = spCore-spImg
Else
  spBase = spBase+spImg
End If
If isSameDir("sadm/edfck/") Then
  spCore = spCore-spEdfck
Else
  spBase = spBase+spEdfck
End If
If isSameDir("upfile/") Then
  spCore = spCore-spUp
  spBase = spBase-spUp
End If
If spDBf = "local" Then
  spUp = spUp-spDB
Else
  spBakDB = fold_Size(Config_Path&"upfile/#dbf#/")
  spDB = spDB+spBakDB
  spUP = spUP-spBakDB
End If
spByte = spBase+spUp+spDB+nDirSize

sqBase = FormatNumber(100*spBase/spByte,2)
sqUp   = FormatNumber(100*spUp/spByte,2)
sqDB   = FormatNumber(100*spDB/spByte,2)
For i=0 To uBound(aDirTab)
  qDirSize(i) = FormatNumber(100*aDirSize(i)/spByte,2)
Next

sqCore = FormatNumber(100*spCore/spBase,2) '%
spCore = Get_BSize(spCore) 'MB
'Response.Write "<br>"&sqCore


'Response.Write "<br>"
'Response.Write "<br>spBase:"&spBase&"(Core:"&spCore&")"
'Response.Write "<br>spUp:"&spUp
'Response.Write "<br>spDB:"&spDB
'Response.Write "<br>spByte:"&spByte

'Response.Write "<br>p:"&Config_Path
'Response.Write "<br>=:"&isSameDir("sadm/")
'Response.Write "<br>=:"&isSameDir("upfile/")


'//////////////////////////////////////////////


spDiv = 530 'Div 宽px
spAll = intSpace '总空间大小MB 
If Int(spAll) < 1 Then 
  spAll = 50
End If
'spByte = fold_Size("../../") '/
spUse = spByte/(1024*1024) 
spUseS = Get_BSize(spByte) 
p1 = FormatNumber(100*spUse/spAll,2)
p2 = 100-p1
w1 = Int(spDiv*p1/100)
w2 = spDiv-w1

If p1>=100 Then
  f = "Stop"
ElseIf p1>90 Then 
  f = "Alert"
Else 
  f = "None"
End If

spUse = FormatNumber(spUse,2)
spRem = spAll-spUse
If p1>=100 Then
  w1 = spDiv
  w2 = 0
  s1 = "已用"&p1&"% &nbsp; ("&spUse&"MB)"
  s2 = ""
ElseIf p1>=50 Then
  s1 = "已用"&p1&"% &nbsp; ("&spUse&"MB)"
  s2 = ""
Else
  s1 = ""
  s2 = "空闲"&p2&"% &nbsp; ("&spRem&"MB)"
End If


%>
<style type="text/css">
.spBox {
	height:24px;
	text-align:center;
	margin:0px;
	padding:0px;
}
.spBox0 {
 width:<%=spDiv%>px;
	border:3px solid #CCC;
	background-color:#9966FF;
	overflow:hidden;
}
.spBox1 {
 width:<%=w1%>px;
	color:#FFF;
	background-color:#6666FF;
	float:left;
	overflow:hidden;
}
.spBox2 {
 width:<%=w2%>px;
	color:#000;
	background-color:#FFFFF0;
	float:right;
	overflow:hidden;
}
div table.sp3Tab {
	margin:0px;
	padding:0px;
	border:0px;
}
div table.sp3Tab td{
	margin:0px;
	padding:0px;
	line-height:150%;
}
div table.sp3Tab td.sp3note{
	padding-left:3px;
}
div table.sp3Tab td.spBase {
	background-color:#006;
}
div table.sp3Tab td.spDB {
	background-color:#060;
}
div table.sp3Tab td.spUP {
	background-color:#600;
}
div.bgCore, span.bgCore{
    background-color:#00C; 	
	display:inline-block;
}
div.txCore, span.txCore{
    background-color:#00C; 	
	display:inline-block;
	padding:2px 2px;
	line-height:100%;
	margin:0px;
}
div table.sp3Tab td.spVDir {
	background-color:#366;
}
td.fColor, span.fColor{
	color:#CCC;
}
#sp3x td{
    padding:1px 8px;
	margin:0px;
	line-height:150%;
}
</style>
<a name="spView"></a>
<table  border="0" align="center" cellpadding="0" cellspacing="1">
  <tr>
    <td align="left"><strong><span class="songti">&#8226;</span>空间占用情况</strong>：总空间:<%=spAll%>(MB)，已用: <%=spUseS%>，百分比:<%=p1%>(%) 
      &nbsp; <a href="../../smod/adupd/para_set1.asp?ID=intSpace&Flag=Num&nLen=6&fRet=Home">[设置参数]</a> - [<a href="#spView" id="setFlag" onClick="setView()">使用详情</a>]</td>
  </tr>
  <tr id="Common">
    <td align="center"><div class="spBox spBox0" >
        <div class="spBox spBox2"><%=s2%></div>
        <div class="spBox spBox1"><%=s1%></div>
      </div></td>
  </tr>
  <tr id="Supper" style="visibility:hidden;display:none;">
    <td align="center">
     
       <div>
          
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="sp3Tab">
            <tr>
              <td class="spUP" width="<%=Int(sqUP)%>%">&nbsp;</td>
              <td class="spDB" width="<%=Int(sqDB)%>%">&nbsp;</td>
              <td width="<%=Int(sqBase)%>%" align="center" class="spBase">
              <div class="bgCore" style="width:<%=sqCore%>%; line-height:12px">&nbsp;</div>
              </td>
              <%For i=0 To uBound(aDirTab)%>
              <td class="spVDir" width="<%=Int(qDirSize(i))%>%">&nbsp;</td>
              <%Next%>

              
            </tr>
          </table>
      <div style="line-height:1px">&nbsp;</div>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" id="sp3x" class="sp3Tab">

            <tr>
              <td align="center" class="spUP fColor">上传附件+FTP上传</td>
              <td class="sp3note">此为使用空间详情: 附件占用: <%=Get_BSize(spUP)%> 占<%=sqUP%>% </td>
            </tr>
            <tr>
              <td align="center" class="spDB fColor">数据库+<span class="sp3note">备份</span></td>
              <td class="sp3note">数据库占用: <%=Get_BSize(spDB)%> 占<%=sqDB%>% [含备份文件]</td>
            </tr>
            <tr>
              <td width="30%" align="center" class="spBase fColor">基本代码+核心代码</td>
              <td class="sp3note">基本代码: <%=Get_BSize(spBase)%> 占<%=sqBase%>% <span class="txCore fColor">[其中核心代码:<%=spCore%>]</span></td>
            </tr>
            <%For i=0 To uBound(aDirTab)%>
            <tr>
              <td width="30%" align="center" class="spVDir fColor">虚拟目录</td>
              <td class="sp3note">虚拟目录<%=aDirTab(i)%>: <%=Get_BSize(aDirSize(i))%> 占<%=qDirSize(i)%>% </td>
            </tr>
            <%Next%>      
          </table>
          
        </div>   
          
          </td>
  </tr>
  <tr>
    <td align="left"> 注意:实际大小以空间提供商为准,这里只作管理提示, 建议你<a href="../../smod/adupd/para_set1.asp?ID=intSpace&Flag=Num&nLen=6&fRet=Home">[设置参数]</a>使它与你真实空间一致.</td>
  </tr>
</table>

<%
  Msg0 = ""&intSpace&"MB"
  Msg1 = "警告！您设置的空间大小是 "&Msg0&""
  Msg1 = Msg1& "\n使用空间已经达到 "&p1&"%"
  Msg1 = Msg1& "\n\n如果您的实际空间大小不是 "&Msg0&" \n请点[设置参数]自行设置；"
If f="Stop" Then
  Msg1 = Msg1& "\n\n如果确认空间大小是 "&Msg0&" \n请赶快联系空间提供商；"
  Msg1 = Msg1& "\n否则如出现问题，与本系统无关！"
End If
%>

<script type="text/javascript">

<%If inStr("Alert,Stop",f)>0 Then Response.Write("alert('"&Msg1&"');")%>

vCom = getElmID("Common");
vSup = getElmID("Supper");
sFlg = getElmID("setFlag");
function setView(){
  sMsg = sFlg.innerHTML;	
  if(sMsg=="使用详情"){
    vSup.style.display='';
    vSup.style.visibility='visible';
	sFlg.innerHTML ="关闭详情";
  }else{
    vSup.style.display='none';
    vSup.style.visibility='hidden';
	sFlg.innerHTML ="使用详情";
  }
}

</script>

<%
runTim01 = runTim01 - Timer()
Response.Write "<!--"&runTim01&"-->"
%>
<!-- check space 2 -->