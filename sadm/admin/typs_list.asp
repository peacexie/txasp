<!--#include file="config.asp"-->

<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
SET rs=Server.CreateObject("Adodb.Recordset") 

Function tmList(xID)
 Dim rs,s00,sql,sID,sItem,sName
 s00 = ""
 sql = "SELECT SysID,[SysName] FROM [AdmSyst] WHERE SysType='"&xID&"' AND SysTop<'U' ORDER BY SysTop,SysID " 
 SET rs=Server.CreateObject("Adodb.Recordset")
 rs.Open sql,conn,1,1 
 Do While Not rs.EOF 
   sID = rs("SysID")
   sName = rs("SysName")
   sItem = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tm"&sID&"'")
   If sItem="" Then
   sItem = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tm"&Left(sID,4)&"999'")
   End If
   If sItem="" Then
   sItem = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tm"&Left(sID,3)&"9999'")
   End If
   sItem = Replace(sItem,"($ModID)",sID) 'Left(sID,4)&"999"
   sItem = Replace(sItem,"($Name)",sName)
   s00=s00&sItem
 rs.MoveNext
 Loop
 rs.Close()
 SET rs=Nothing
 tmList = s00
End Function

yAct = Request("yAct")&";"
Page = RequestS("Page","N",1)

If inStr(yAct,"Clear;")>0 Then
  Call rs_DoSql(conn,"DELETE FROM WebTyps WHERE TypMod='"&RequestS("ID","C",24)&"'")
  msg = "清空资料成功！"
End If

If inStr(yAct,"setup;")>0 Then
  '清理setup 
  '/setup/ 
  Call files_move("../../setup/","../../sadm/setup/")
  Call fold_del("../../setup/")
  Response.Write "<BR><center>清理setup完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"BBS;")>0 Then
  '清理BBS 
  '/bbs/,/binn/,/smod/bbs/；
  'AdmPara.ParFlag=ParBBS,AdmSyst.SysType=ParBBS；
  'BBSInfo,BPKTitle,BPKView,BPKVote； 
  Call fold_del("../../bbs/")
  Call fold_del("../../binn/")
  Call fold_del("../../smod/bbs/")
  Call rs_DoSql(conn,"DELETE FROM AdmPara WHERE ParFlag='ParBBS'")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysType='BBS'")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ParBBS'")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ModBBS'")
  Call rs_DoSql(conn,"DROP TABLE BBSInfo")
  Call rs_DoSql(conn,"DROP TABLE BPKTitle")
  Call rs_DoSql(conn,"DROP TABLE BPKView")
  Call rs_DoSql(conn,"DROP TABLE BPKVote")
  Response.Write "<BR><center>清理BBS完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"Trade;")>0 Then
  '清理trade 
  '/trade/,/member/madm/,/member/mnews/,/member/mpics/,/member/mpub/；
  'AdmSyst.SysType=Trade；
  'TradeCorp,TradeNews,TradePara,TradePics,TradeType；
  Call fold_del("../../trade/")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysType='Trade'")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ModTrade'")
  Call rs_DoSql(conn,"DROP TABLE TradeCorp")
  Call rs_DoSql(conn,"DROP TABLE TradeNews")
  Call rs_DoSql(conn,"DROP TABLE TradePara")
  Call rs_DoSql(conn,"DROP TABLE TradePics")
  Call rs_DoSql(conn,"DROP TABLE TradeType")
  Response.Write "<BR><center>清理trade完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"msg;")>0 Then
  '清理msg
  '/msg/
  'AdmSyst.SysType=Sms；
  'SmsCharge,SmsMember,SmsSend,SmsTelq,SmsTels,SmsTemp,SmsType
  Call fold_del("../../msg/")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysType='Sms'")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ModSms'")
  Call rs_DoSql(conn,"DROP TABLE SmsCharge")
  Call rs_DoSql(conn,"DROP TABLE SmsMember")
  Call rs_DoSql(conn,"DROP TABLE SmsSend")
  Call rs_DoSql(conn,"DROP TABLE SmsTelq")
  Call rs_DoSql(conn,"DROP TABLE SmsTels")
  Call rs_DoSql(conn,"DROP TABLE SmsTemp")
  Call rs_DoSql(conn,"DROP TABLE SmsType")
  Response.Write "<BR><center>清理msg完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If
If inStr(yAct,"Vote;")>0 Then
  '清理vote 
  '/vote/,/smod/vote/；
  'VoteInfo,VoteItem,VoteLogs； 
  Call fold_del("../../vote/")
  Call fold_del("../../smod/vote/")
  'Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysType='Vote'")
  Call rs_DoSql(conn,"DROP TABLE VoteInfo")
  Call rs_DoSql(conn,"DROP TABLE VoteItem")
  Call rs_DoSql(conn,"DROP TABLE VoteLogs")
  Response.Write "<BR><center>'清理vote完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"Member;")>0 Then
  '清理member 
  '/member/,/inc/mem_inc/；
  'AdmSyst.SysType=Member；
  'Member_ABCDE,MemCard,MemSyst； 
  Call fold_del("../../member/")
  Call fold_del("../../inc/mem_inc/")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysType='Member'")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ModMember'") 
  '' Call rs_DoSql(conn,"DROP TABLE Member_ABCDE")
  Call rs_DoSql(conn,"DROP TABLE MemCard")
  Call rs_DoSql(conn,"DROP TABLE MemSyst")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ModMember'")
  Response.Write "<BR><center>'清理member 完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"IDocs;")>0 Then
  '清理IDocs 
  '/doc/；；
  'AdmUser_12345.UsrType=Inn%；
  'DocsLogs,DocsNews,DocsRemark； 
  Call fold_del("../../doc/")
  Call rs_DoSql(conn,"DELETE FROM AdmUser_12345 WHERE UsrType LIKE 'Inn%' AND LEN(UsrName)<4")
  Call rs_DoSql(conn,"DROP TABLE DocsLogs")
  Call rs_DoSql(conn,"DROP TABLE DocsNews")
  Call rs_DoSql(conn,"DROP TABLE DocsRemark")
  Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='ModInnDocs'")
  Response.Write "<BR><center>'清理IDocs 完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"PubData;")>0 Then
  '清理PubData 
  'Config_Path&"upfile/#dbf#/","ysWeb_PubData.Peace!DB"
  'CrdDemo.xls
  f2 = Request("f2")
  If f2<>"" And inStr(";CrdDemo.xls;",f2)>0 Then
    Call fil_del(Config_Path&"upfile/#dbf#/CrdDemo.xls")
  Else
    Call fil_del(Config_Path&"upfile/#dbf#/ysWeb_PubData.Peace!DB")
  End If
  Response.Write "<BR><center>'清理PubData 完成。<a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
End If

If inStr(yAct,"ClSys;")>0 Then  
  sql2 = " SELECT SysID FROM AdmSyst WHERE SysType='Groups' And SysID='x($SysType)' "
  sql3 = " SELECT SysID FROM AdmSyst WHERE SysType='Module' And SysID='Mod($SysType)' "
  
  sql = " SELECT * FROM [AdmSyst] ORDER BY SysType,SysID  "
  rs.Open sql,conn,1,1 '3
  Do While NOT rs.EOF 
   SysID = rs("SysID")
   SysType = rs("SysType") 'BBSP124:公开论坛:BBS 
   f2ID = rs_Exist(conn,Replace(sql2,"($SysType)",SysType))
   f3ID = rs_Exist(conn,Replace(sql3,"($SysType)",SysType))
   If f2ID="EOF" AND f3ID="EOF" Then
     Response.Write "<br>清理:"&SysID&":"&rs("SysName")&":"&SysType
	 Call rs_DoSql(conn,"DELETE FROM AdmSyst WHERE SysID='"&SysID&"'")
   End If 
   rs.Movenext
  Loop 
  rs.Close()
  
  sql4 = "DELETE FROM AdmPara WHERE ParFlag NOT IN(SELECT SysID FROM AdmSyst WHERE SysType='Para') " 
  Call rs_DoSql(conn,sql4)
  sql4 = "DELETE FROM WebTyps WHERE TypMod  NOT IN(SELECT SysID FROM AdmSyst) " 
  Call rs_DoSql(conn,sql4)
End If

If inStr(yAct,"ResType;")>0 Then 
  Response.Write "<BR><center><a href='?'><font color='#0000FF'>OK返回</font></a></center><HR>"
  s = ""
  sql = "SELECT SysID,[SysName] FROM [AdmSyst] WHERE SysType='Module' ORDER BY SysTop,SysID " 'And SysTop<'U' 
  
  rs.Open sql,conn,1,1 
  s1="" : s2="" : s3="" :mStr=""
  Do While Not rs.EOF 
    SysName = rs("SysName")
    SysID = rs("SysID")
    sAll = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tm"&Mid(SysID,4)&"All'")
    sLst = tmList(Mid(SysID,4))
    sEnd = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tm"&Mid(SysID,4)&"End'")
    sAll = sAll&sLst&sEnd
    sAll = Replace(sAll,"</li><li>","</li>"&vbcrlf&"<li>")
    Response.Write vbcrlf&vbcrlf&"<br>----------"&SysID&":"&SysName&vbcrlf
    Response.Write sAll
	' Check rM&Mid(SysID,4) 是否存在？　Insert Or Update
	If rs_Exist(conn,"SELECT ParCode FROM [AdmPara] WHERE ParCode='rM"&Mid(SysID,4)&"'") = "EOF" then
	  sql = " INSERT INTO [AdmPara] (ParName, ParCode, ParFlag, ParRem,LogAddIP,LogAUser,LogATime)VALUES(" 
	  sql = sql& "  '" &SysName&"', 'rM" &Mid(SysID,4)&"', 'ParMenu', '" &Replace(sAll,"'","''")&"'" 
	  sql = sql& " ,'"&Get_CIP()&"','"&Session("UsrID")&"','"&Now()&"')"
	  Call rs_Dosql(conn,sql)
	Else
	 If sAll&sLst&sEnd<>"" Then
	  Call rs_DoSql(conn,"UPDATE AdmPara SET ParName='"&SysName&"',ParRem='"&Replace(sAll,"'","''")&"' WHERE ParCode='rM"&Mid(SysID,4)&"'")
	 End If
	End If
	
  rs.MoveNext
  Loop
  rs.Close()
  
  Response.Write "<HR><IFRAME name=LeftMenu src='../system/upd_para.asp' scrolling='no'></IFRAME>"
End If

    sql = " SELECT * FROM [AdmSyst] "
	'sql =sql& " WHERE SysType NOT IN('Groups','Depart','Para','Module','Admin','Inner','System') "
	sql =sql& " WHERE SysID IN(SELECT TypMod FROM WebTyps) "
	sql =sql& " AND SysTop < 'u' "
	sql =sql& " ORDER BY SysType,SysTop,SysID " 
	
   rs.Open Sql,conn,1,1
   rs.PageSize = 15
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

%>
<br>
<table width="640" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="20%" align="center" bgcolor="#FFFFFF"><strong>信息类别</strong> </td>
          <td align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;<font color="#FF0000"><%=msg%></font></td>
          <%If Config_Mode = "isExpert" Then%>
		  <!--#include file="inc_menu.asp"-->
          <%End If%>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="8" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td width="5%" height="27" align="center" nowrap>NO</td>
    <td height="27" align="center">名称</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">编码</td>
    <td align="center" nowrap bgcolor="E0E0E0">Top</td>
    <td align="center" nowrap bgcolor="E0E0E0">N</td>
    <td align="center" nowrap bgcolor="E0E0E0">查看</td>
    <td align="center" nowrap bgcolor="E0E0E0">清空</td>
    <td height="27" align="center" nowrap bgcolor="E0E0E0">SysType</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="8" align="right" nowrap></td>
  </tr>
  <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
SysID = rs("SysID")
SysType = rs("SysType")
SysName = rs("SysName")
SysTop = rs("SysTop")
N = rs_Count(conn,"WebTyps WHERE TypMod='"&SysID&"'")
	  %>
  <form name="flist<%=i%>" method="post" action="?">
    <tr bgcolor="<%=col%>">
      <td align="right" nowrap><%=i%></td>
      <td nowrap><%=SysName%>
        <input name="yAct" type="hidden" id="yAct" value="edt">
      </td>
      <td align="left" nowrap><%=SysID%>
        <input name="KeyID" type="hidden" id="SysID" value="<%=SysID%>"></td>
      <td align="center" nowrap bgcolor="<%=col%>"><%=SysTop%></td>
      <td align="center" nowrap bgcolor="<%=col%>"><%=N%></td>
      <td align="center" nowrap>
      <%If Int(N) > "0" Then%>
      <a href="../../smod/type/type_list.asp?ModID=<%=SysID%>">查看</a>
      <%Else%>
      <span class="colCCC">查看</span>
	  <%End If%>
      </td>
      <td align="center" nowrap>
      <%If Config_Mode = "isExpert" Then%>
      <a onClick="Del_YN('?ID=<%=SysID%>&Page=<%=Page%>&yAct=Clear','确认删除?小心操作哦！')" href="#" >清空</a>
      <%Else%>
      <span class="colCCC">清空</span>
	  <%End If%>
      
      </td>
      <td align="center" nowrap><%=SysType%></td>
    </tr>
  </form>
  <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
  <%
  End If
	  
	  rs.Close()
	  
	  %>
  <form name="fmDir" method="post" action="" Xtarget="_blank">
  <tr bgcolor="#999999">
    <td colspan="8" align="left" bgcolor="#F0F0F0"><input name="sDir" type="text" id="sDir" value="../../smod/type/type_list.asp?ModID=InfX123" size="45" maxlength="255">
      <input type="button" name="button" id="button" value="类别管理" onClick="goType()">
      &nbsp;(用于主菜单中未出现的类别管理) </td>
  </tr>
  </form>
  <form name="fmRep" method="post" action="?" target="_self">
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="8" align="left">Tab:
        <input name="sTab" type="text" id="sTab" value="AdmPara" size="12" maxlength="255">
Col:
<input name="sCol" type="text" id="sCol" value="ParRem" size="12" maxlength="255">
Key:
        <input name="sKey" type="text" id="sKey" size="12" maxlength="255">
        (KeyID)<br>
<input name="sStr1" type="text" id="sStr1" value="/u/xtest21/upfile/" size="24" maxlength="255">
        -=&gt;
        <input name="sStr2" type="text" id="sStr2" value="/upfile/" size="24" maxlength="255">
        <input type="submit" name="button" id="button" value="替换">
        &nbsp;<br>
        (用于重新配置网站目录等后...)
        <input name="yAct" type="hidden" id="yAct" value="UpdData"> <span class="colF0F">正常运行的系统，请不要随便执行该操作！否则后果自负！！！</span></td>
    </tr>
  </form>  
      <%If Config_Mode = "isExpert" Then%>  
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center">清理项目</td>
      <td colspan="6" align="left">&nbsp;&nbsp;&nbsp;&nbsp;<span class="colF0F">用于初始配置网站，清理一些没有用的模块和数据；如果出错没有关系，可能是没有写权限或已经删除相关资料！正常运行的系统，请不要随便执行该操作！否则后果自负！！！</span><br>
      ---目录文件---；---数据资料---；---数据表---</td>
    </tr>
    <%If fold_exist("../../","setup") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><a href="#" onClick="Del_YN('typs_list.asp?yAct=setup','确认清理setup?小心操作哦！')">清理setup</a></td>
      <td colspan="6" align="left">/setup/：backup to /sadm/setup/</td>
    </tr>
    <%End If%>
    <%If fold_exist("../../bbs/","badm") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><a href="#" onClick="Del_YN('?yAct=BBS','确认清理BBS? 小心操作哦！')">清理BBS</a></td>
      <td colspan="6" align="left">/bbs/,/binn/,/smod/bbs/；<br>
      AdmPara.ParFlag=ParBBS,AdmSyst.SysType=ParBBS；<br>
      BBSInfo,BPKTitle,BPKView,BPKVote；</td>
    </tr>
    <%End If%>
    <%If fold_exist("../../trade/","madm") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><a href="#" onClick="Del_YN('?yAct=Trade','确认清理trade?小心操作哦！')">清理trade</a></td>
      <td colspan="6" align="left">/trade/,/member/madm/,/member/mnews/,/member/mpics/,/member/mpub/；<br>
        AdmSyst.SysType=Trade；<br>
        TradeCorp,TradeNews,TradePara,TradePics,TradeType；</td>
    </tr>
    <%End If%>
    <%If fold_exist("../../","msg") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><a href="#" onClick="Del_YN('?yAct=msg','确认清理trade?小心操作哦！')">清理sms</a></td>
      <td colspan="6" align="left">/sms/；<br>
        AdmSyst.SysType=Sms；<br>
      SmsCharge,SmsMember,SmsSend,SmsTelq,SmsTels,SmsTemp,SmsType；</td>
    </tr>
    <%End If%>
    <%If fold_exist("../../smod/","vote") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><a href="#" onClick="Del_YN('?yAct=Vote','确认清理vote?小心操作哦！')">清理vote</a></td>
      <td colspan="6" align="left">/vote/,/smod/vote/；<br>
        VoteInfo,VoteItem,VoteLogs；</td>
    </tr>
    <%End If%>
    <%If fold_exist("../../inc/","mem_inc") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><a href="#" onClick="Del_YN('?yAct=Member','确认清理member?小心操作哦！')">清理member</a></td>
      <td colspan="6" align="left">/member/,/inc/mem_inc/；<br>
      AdmSyst.SysType=Member；<br>
      <span class="colCCC">Member_ABCDE</span>,MemCard,MemSyst；</td>
    </tr>
    <%End If%>
    <%If fold_exist("../../","doc") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><p><a href="#" onClick="Del_YN('?yAct=IDocs','确认清理doc?小心操作哦！')">清理idoc<br>
        (内部公文)
      </a></p></td>
      <td colspan="6" align="left">/doc/；<br>
      AdmUser_12345.UsrType=Inn%；<br>      
      DocsLogs,DocsNews,DocsRemark；</td>
    </tr>
    <%End If%>
    <%If fil_exist(Config_Path&"upfile/#dbf#/ysWeb_PubData.Peace!DB") Then%>
    <tr align="center" bgcolor="#FFFFFF">
      <td height="22" colspan="2" align="center"><p><a href="#" onClick="Del_YN('?yAct=PubData','确认清理PubData?小心操作哦！')">清理PubData<br>
        (PubData)
      </a></p></td>
      <td colspan="6" align="left">ysWeb_PubData.Peace!DB；<br>
        <a href="?yAct=PubData&f2=CrdDemo.xls">CrdDemo.xls</a>；</td>
    </tr>
    <%End If%>
    
    
	  <%End If%>

</table>
<%

Response.Write yAct
If inStr(yAct,"UpdData")>0 Then 

 sTab = RequestS("sTab","C",48)
 sCol = RequestS("sCol","C",48)
 sKey = RequestS("sKey","C",48)
 sStr1 = RequestS("sStr1","C",120)
 sStr2 = RequestS("sStr2","C",120)
 
 If Len(sStr1)>2 And Len(sStr2)>2 AND sTab<>"" AND sCol<>"" AND sKey<>"" Then

  'sql = "SELECT "&sKey&","&sCol&" FROM ["&sTab&"] WHERE "&sCol&" LIKE '%"&sStr1&"%' "
  nRes = SysDBSwap(sKey,sCol,sTab," WHERE "&sCol&" LIKE '%"&sStr1&"%' ",sStr1,sStr2)
 
 Response.Write "<br>"&nRes&"条记录,处理完成:"&Now()
 End If

End If


SET rs=Nothing
%>

<!--


  Response.Write "<pre>"
  iTab = "InfoNews"
  iMod = "InfN124"
  SET rs=Server.CreateObject("Adodb.Recordset") 
  sql = "SELECT KeyID,LogATime FROM ["&iTab&"] WHERE KeyMod='"&iMod&"' "
  rs.Open sql,conn,1,1 
  if NOT rs.eof then 
   Do While NOT rs.EOF
    KeyID = rs("KeyID")
	AddTM = rs("LogATime")
	TimID = Get_TimID(AddTM)
	If Left(KeyID,10)<>Left(TimID,10) Then
	  NewID = TimID&Rnd_ID("KEY",12)
	  'Call rs_DoSql(conn,"UPDATE "&iTab&" SET KeyID='"&AutID&"' WHERE KeyID='"&KeyID&"'")
	  Response.Write "<br>"&KeyID&" : "&TimID&" : "&AddTM&" : "&NewID&""
	End If
    rs.MoveNext()
   Loop
  End If
  rs.Close()
  SET rs=Nothing 
  Response.Write "</pre>"


Function Get_TimID(xNow)
Dim YMD,HMS,str,tDate,ny,nm,nd,RDX,hh,mm,ss
   tDate = xNow
   ny = DatePart("yyyy",tDate)
   nm = DatePart("m",tDate)
   nd = DatePart("d",tDate)
  YMD = Hex(ny*380+nm*31+nd)
  YMD = Right("00000"&YMD,6)
   hh = DatePart("h",tDate)
   mm = DatePart("n",tDate)
   ss = DatePart("s",tDate)
  HMS = Hex((hh*60*60+mm*60+ss)*100)
  HMS = Right("00000"&HMS,6)
  str = YMD&HMS
Get_TimID = str
End Function


-->

<script type="text/javascript">
function goType()
{
  document.fmDir.action = document.fmDir.sDir.value;
  document.fmDir.submit();
}
</script>
</body>
</html>
