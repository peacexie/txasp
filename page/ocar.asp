<!--#include file="_config.asp"-->
<%

If verNow="2" Then

  vOrd_m0Title = "Order Product"
  vOrd_m0ErrCode = "Error Code!"
  vOrd_m0EdtCode = "System find the same Order ID; \nOrder ID Changed as"
  vOrd_m0OrderOK = "Order OK!"
  vOrd_m0EditOK = "Edit OK!"
  vOrd_m0ErrCheck = "Error Check Code!"
  vOrd_m0AddOK = "Add OK!"
  
  vOrd_m1OK = "Message: Order Send OK!"
  vOrd_m1ID = "Please remember the Order ID: "
  vOrd_m1View = "View Products"
  vOrd_m1Close = "[Close Window]"
  
  vOrd_m2Info = "Order Info."
  vOrd_m2Subj = "Order Title"
  vOrd_m2Code = "Order Code"
  vOrd_m2Auto = "(Auto)"
  vOrd_m2Money = "Money"
  vOrd_m2Ordt = "Order"
  vOrd_m2Note = "Notes<br />(&lt;120 words)"
  
  vOrd_m3Items = "Items"
  vOrd_m3Name = "Name"
  vOrd_m3Price = "Price"
  vOrd_m3Count = "Count"
  vOrd_m3Subtl = "Subtotal" 
  vOrd_m3Total = "Total"
  vOrd_m3Del = "Delete"
  vOrd_m3Edit = "Edit"
  
  vOrd_m4Send = "Delivery mode"
  vOrd_m4Time = "Time requirements"
  vOrd_m4Link = "Contact Info."
  vOrd_m4Name = "Name"
  vOrd_m4Title = "Title"
  vOrd_m4Man = "Man"
  vOrd_m4Female = "Female"
  vOrd_m4Addr = "Address"
  vOrd_m4Post = "Post Code"
  vOrd_m4Mobile = "Mobile"
  vOrd_m4Tel = "Tel"
  
  vOrd_m5Pays = "Pay Mathed"
  vOrd_m5Send = "Send Order"
  vOrd_m5SendOK = "Send Order"
  
  vOrd_jsName = "Input Name!"
  vOrd_jsAddr = "Input Address!"
  vOrd_jsTel = "Input Tel or Mobile!"
  vOrd_jsSubj = "Input Subject!"
  vOrd_jsNote = "Notes Must in 120 words!"
  vOrd_jsCode = "Input Check Code!"
  
  vOrd_m6Time = "Order Time"
  vOrd_m6Rem = "Please remember this NO."
  vOrd_m6Unit = "RMB"
  
Else

  vOrd_m0Title = "订购商品"
  vOrd_m0ErrCode = "认证码错误!"
  vOrd_m0EdtCode = "因为多人同时再线订购，发现定单号重复；\n定单号现自动改为"
  vOrd_m0OrderOK = "订购商品成功!"
  vOrd_m0EditOK = "修改成功!"
  vOrd_m0ErrCheck = "认证码错误!"
  vOrd_m0AddOK = "增加成功!"
  
  vOrd_m1OK = "提示：定单提交成功！"
  vOrd_m1ID = "请记住定单号："
  vOrd_m1View = "继续查看产品"
  vOrd_m1Close = "[关闭窗口]"
  
  vOrd_m2Info = "订单信息"
  vOrd_m2Subj = "定单标题"
  vOrd_m2Code = "定单编号"
  vOrd_m2Auto = "(自动)"
  vOrd_m2Money = "合计金额"
  vOrd_m2Ordt = "定单"
  vOrd_m2Note = "附加留言<br />(120字内)"
  
  vOrd_m3Items = "货款信息"
  vOrd_m3Name = "商品名称"
  vOrd_m3Price = "价格"
  vOrd_m3Count = "数量"
  vOrd_m3Subtl = "小计" 
  vOrd_m3Total = "合计"
  vOrd_m3Del = "删除"
  vOrd_m3Edit = "修改"
  
  vOrd_m4Send = "配送方式"
  vOrd_m4Time = "时间要求"
  vOrd_m4Link = "送货地址 和 联系方式"
  vOrd_m4Name = "收货人"
  vOrd_m4Title = "称 谓"
  vOrd_m4Man = "先生"
  vOrd_m4Female = "女士"
  vOrd_m4Addr = "地 址"
  vOrd_m4Post = "邮政编码"
  vOrd_m4Mobile = "手 机"
  vOrd_m4Tel = "电 话"
  
  vOrd_m5Pays = "支付方式"
  vOrd_m5Send = "确认提交"
  vOrd_m5SendOK = "确认提交定单"
  
  vOrd_jsName = "收货人 不能为空!"
  vOrd_jsAddr = "地 址 不能为空!"
  vOrd_jsTel = "手机,电话 不能同时为空!"
  vOrd_jsSubj = "定单标题 不能为空!"
  vOrd_jsNote = "附加留言 请<120字!"
  vOrd_jsCode = "认证码 不能为空!"
  
  vOrd_m6Time = "订购时间"
  vOrd_m6Rem = "请记录此定单号"
  vOrd_m6Unit = "元"
  
End If

MD = RequestS("MD","C",48) '"PicS124"
If Session("MemID")&""<>"" Then
  MemID = Session("MemID")
Else
  MemID = "("&Session.SessionID&")"
End If

Act = RequestS("Act","C",48)

If Act="OAdd" Then

 KeyCode = RequestS("KeyCode","C",48)
 InfSubj = RequestS("InfSubj","C",48)
 InfCont = RequestS("InfCont","C",255)

 LnkName   = RequestS("LnkName","C",48)	
 LnkSex    = RequestS("LnkSex","C",48)	
 LnkAddr   = RequestS("LnkAddr","C",255)	
 LnkPost   = RequestS("LnkPost","C",48)	
 LnkMobile = RequestS("LnkMobile","C",48)	
 LnkTel    = RequestS("LnkTel","C",48)	
 LnkEmail  = RequestS("LnkEmail","C",255)
 
 InfNum   = RequestS("InfNum","N",0)		
 InfPay   = RequestS("r3","C",48)		
 InfSend  = RequestS("r1","C",48)	
 InfTime  = RequestS("r2","C",48)	
 InfNow   = Now()
 
 SetCheck  = "-"
 SetPay    = "-"
 SetSend   = "-"	
 SetState  = "-"
 KeyID = Get_AutoID(24)
		
 sql = " INSERT INTO [OrdInfo] (" 
 sql = sql& "  KeyID,KeyCode,InfSubj,InfCont" 
 sql = sql& ", LnkName,LnkSex,LnkAddr,LnkPost,LnkMobile,LnkTel,LnkEmail" 
 sql = sql& ", InfNum,InfPay,InfSend,InfTime,SetCheck,SetPay,SetSend,SetState" 
 sql = sql& ",LogAddIP,LogAUser,LogATime" 
 sql = sql& ")VALUES(" 
 sql = sql& "  '" &KeyID&"','"&KeyCode&"','"&InfSubj&"','"&InfCont&"'" 
 sql = sql& ", '" &LnkName&"','"&LnkSex&"','"&LnkAddr&"','"&LnkPost&"','"&LnkMobile&"','"&LnkTel&"','"&LnkEmail&"'" 
 sql = sql& ",  " &InfNum&" ,'"&InfPay&"','"&InfSend&"','"&InfTime&"','"&SetCheck&"','"&SetPay&"','"&SetSend&"','"&SetState&"'" 
 sql = sql& " ,'"&Get_CIP()&"','"&MemID&"','"&InfNow&"'"
 sql = sql& ")" ': Response.Write sql
 ChkCode = Request("ChkCode") 
 If Session("ChkCode")<>uCase(ChkCode) Then
   Response.Write js_Alert(vOrd_m0ErrCode,"Redir","?")
 Else
   If rs_Exist(conn,"SELECT KeyCode FROM [OrdInfo] WHERE KeyCode='"&KeyCode&"'")<>"EOF" Then
      NewCode = rs_OrdID(conn,"yymmdd-r3",6,3)
	  sql = Replace(sql,",'"&KeyCode&"'",",'"&NewCode&"'")
	  KeyCode = NewCode
	  Msg2 = "\n"&vOrd_m0EdtCode&": "&NewCode&""
   Else
      Msg2 = ""
   End If
   Call rs_DoSql(conn,sql)
   Call rs_DoSql(conn,"UPDATE [OrdItem] SET KeyCode='"&KeyID&"' WHERE KeyCode='---' AND LogAUser='"&MemID&"'")
   Response.Write js_Alert(vOrd_m0OrderOK&Msg2,"Redir","?Act=OK&KeyCode="&KeyCode&"&InfNum="&InfNum&"&InfSubj="&InfSubj&"&InfNow="&InfNow&"&InfPay="&InfPay&"")
 End If

ElseIf Act="PDel" Then

   k = RequestS("kID","C",48)
   sql = "DELETE FROM OrdItem WHERE KeyID='"&k&"' AND KeyCode='---' AND LogAUser='"&MemID&"'"
   Call rs_DoSql(conn,sql)
   'Response.Write js_Alert("删除成功!","Redir","?")

ElseIf Act="PEdit" Then

   i = RequestS("i","N",0)
   n = RequestS("i"&i&"Num","N",1)
   k = RequestS("i"&i&"ID","C",48)
   sql = "UPDATE OrdItem SET InfCount="&n&" WHERE KeyID='"&k&"' AND KeyCode='---' AND LogAUser='"&MemID&"'"
   Call rs_DoSql(conn,sql)
   Response.Write js_Alert(vOrd_m0EditOK,"Redir","?")

ElseIf Act="PAdd" Then
 iPrice = RequestS("InfPrice","N",0) 
 iNum = RequestS("OrdNum","N",0) 
 iID = RequestS("ID","C",48) 
 sql = " INSERT INTO [OrdItem] (" 
 sql = sql& "  KeyID,KeyMod,KeyCode" 
 sql = sql& ", InfProID,InfPrice,InfCount" 
 sql = sql& ",LogAddIP,LogAUser,LogATime" 
 sql = sql& ")VALUES(" 
 sql = sql& "  '" & Get_AutoID(24) &"','PicS124','---'" 
 sql = sql& ", '" & iID &"', " & iPrice &", " & iNum &"" 
 sql = sql& " ,'"&Get_CIP()&"','"&MemID&"','"&Now()&"'"
 sql = sql& ")"
 ChkCode = Request("ChkCode") 
 If Session("ChkCode")<>uCase(ChkCode) Then
   Response.Write js_Alert(vOrd_m0ErrCheck,"Back",-1)
 Else
   Call rs_DoSql(conn,sql)
   Response.Write js_Alert(vOrd_m0AddOK,"Redir","?")
 End If

End If

DefID = rs_OrdID(conn,"yymmdd-r3",6,3)
'rs_OrdID(conn,"yymmdd-hnsx-s1r1",6,4)
'rs_OrdID(conn,"yymd-hnsx-r2",4,2)

InfNum1 = 0
InfNum2 = 0
InfNum3 = 0

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=vOrd_m0Title%>-<%=vPMsg_WName%></title>
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.OrdFLeg {
	padding:2px 5px 2px 5px;
	font-weight:bold;
}
.OrdFSet {
	padding:5px 5px 5px 5px;
	margin:2px 5px;
}
</style>
</head>
<body style='background-color:#FFF'>
<%
If Act="OK" Then
 Response.Write "<body style='background-color:#FFF'>"
 KeyCode = RequestS("KeyCode","C",48)
 InfSubj = RequestS("InfSubj","C",120)
 InfNum   = RequestS("InfNum","N",0)		
 InfPay   = RequestS("InfPay","C",48)		
 InfNow  = RequestS("InfNow","C",48)
 
 sql = "SELECT * FROM [OrdPara] WHERE KeyID='"&InfPay&"' "
 rs.Open sql,conn,1,1 
 If NOT rs.EOF Then
PaySubj = Show_Form(rs("InfSubj"))
PayCont = Show_Text(rs("InfCont"))
PayPic = rs("InfPic")
If PayPic="" Then
  PayPic = Config_Path&"img/tool/no_pic_160120.jpg"
Else
  PayPic = PayPic
End If
 End If
 rs.Close()
	
%>

<table width="640" border="0" align="center">
  <form action="" method="post" name="fm01" id="fm01">
    <tr>
      <td><fieldset class="OrdFSet">
        <legend class="OrdFLeg"><%=vOrd_m2Info%></legend>

          <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#f4f4f4">
            <%If verNow="1" Then%>
            <tr>
              <td colspan="2" align="center" valign="middle" bgcolor="#FFFFFF" class="fntF00 f14px">定单生成成功！感谢您的订购！<br />
                我们的工作人员会在24小时内与您联系，确认订单并沟通付款事宜。</td>
            </tr>
            <%Else%>
            <tr>
              <td colspan="2" align="center" valign="middle" bgcolor="#FFFFFF" class="fntF00 f14px">English ... 定单生成成功！感谢您的订购！<br />
                我们的工作人员会在24小时内与您联系，确认订单并沟通付款事宜。</td>
            </tr>
            <%End If%>
            <tr>
              <td align="right" valign="middle" bgcolor="#FFFFFF"><%=vOrd_m2Subj%>：</td>
              <td valign="middle" bgcolor="#FFFFFF"><%=InfSubj%></td>
            </tr>
            <tr>
              <td align="right" valign="middle" bgcolor="#FFFFFF"><%=vOrd_m6Time%>：</td>
              <td valign="middle" bgcolor="#FFFFFF"><%=InfNow%>
              <IFRAME name=SmsFrame src="../msg/out/smsapi.asp?Peace_Sms_RndID=<%=Timer%>&acu=SOrder&OrdID=<%=KeyCode%>" frameBorder=0 width="12" scrolling=no height="8"></IFRAME>
              </td>
            </tr>
            <tr>
              <td align="right" valign="middle" bgcolor="#FFFFFF"><%=vOrd_m2Code%>：</td>
              <td valign="middle" bgcolor="#FFFFFF"><%=KeyCode%> (<span class="fntF00"><%=vOrd_m6Rem%></span>)</td>
            </tr>
            <tr>
              <td width="20%" align="right" valign="middle" bgcolor="#FFFFFF"><%=vOrd_m2Money%>：</td>
              <td valign="middle" bgcolor="#FFFFFF"><%=InfNum%> (<%=vOrd_m6Unit%>)</td>
            </tr>
          </table>

      </fieldset>
        <div class="line10">&nbsp;</div>
        <fieldset class="OrdFSet">
        <legend class="OrdFLeg"><%=vOrd_m5Pays%></legend>


            <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
              
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center">
                <%=PaySubj%></td>
              <td valign="top"><%=PayCont%></td>
              <td width="30%" align="center"><img src="<%=PayPic%>" alt="<%=PaySubj%>" width="150" height="80" border="0" onload="javascript:setImgSize(this);" /></td>
            </tr>
              
            </table>

          <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" xbgcolor="#E0E0E0">
            <tr>
              <td align="center" valign="middle" bgcolor="#FFFFFF"><span class="OrdButton"> <a href="info.asp?ModID=PicS124"><%=vOrd_m1View%></a> </span><span class="OrdButton"> <a href="javascript:window.close()"><%=vOrd_m1Close%></a> </span></td>
            </tr>
          </table>
        </fieldset>
        <div class="line10">&nbsp;</div></td>
    </tr>
    <input name="Act" type="hidden" id="Act" value="Insert" />
  </form>
</table>
<%
Response.End()
End If


%>
<table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
  <form id="fmOrd01" name="fmOrd01" method="post" action="?">
    <tr>
      <td height="500" colspan="4" align="left" valign="top" style="BORDER:#E0E0E0 1px solid; padding:5px;"><!--Item Start-->
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg"><%=vOrd_m2Info%></legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center"><%=vOrd_m2Subj%></td>
              <td colspan="3"><input name="InfSubj" type="text" id="InfSubj" style="width:480px" value="<%=vOrd_m2Ordt%>-<%=DefID%>" size="60" maxlength="24" /></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center"><%=vOrd_m2Code%></td>
              <td width="30%"><input name="KeyCode" type="text" id="KeyCode" value="<%=DefID%>" size="20" maxlength="24" readonly="readonly" />
              <%=vOrd_m2Auto%></td>
              <td width="20%" align="center"><%=vOrd_m2Money%></td>
              <td width="30%"><input name="InfNum" type="text" id="InfNum" size="20" maxlength="24" readonly="readonly" />
              <%=vOrd_m2Auto%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center"><%=vOrd_m2Note%></td>
              <td colspan="3"><textarea name="InfCont" cols="60" rows="3" id="InfCont" style="width:480px"></textarea></td>
            </tr>
          </table>
        </fieldset>
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg"><%=vOrd_m3Items%> 　 <a href="info.asp?ModID=PicS124" class="cRed" style="font-weight:normal"><%=vOrd_m1View%>&gt;&gt;&gt;</a> 　 </legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td align="center" bgcolor="#FFFFF0">NO.</td>
              <td width="30%" bgcolor="#FFFFF0"><%=vOrd_m3Name%></td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><%=vOrd_m3Price%></td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><%=vOrd_m3Count%></td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><%=vOrd_m3Subtl%></td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;</td>
            </tr>
            <%
		 
sql = "SELECT * FROM [OrdItem] "
sql = sql & " WHERE KeyCode='---' AND LogAUser='"&MemID&"' "
sql = sql & " ORDER BY KeyID "
'Response.Write sql
aNum = 0
aSum = 0
aStr = ""
rs.Open sql,conn,1,1 
If rs.RecordCount <=1 Then
  dis1 = "disabled='disabled'"
End If
If rs.RecordCount <=0 Then
  dis2 = "disabled='disabled'"
End If
rs.PageSize = 240 
		 
		  if NOT rs.EOF then
rs.AbsolutePage = 1 
     for i=1 to rs.PageSize 
if i mod 2 = 1 then
   col = "F8F8F8"
else
   col = "FFFFF8"
end if

KeyID = rs("KeyID") 'Show_Form()
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfProID = rs("InfProID")
InfSubj = rs_Val("","SELECT InfSubj FROM InfoPics WHERE KeyID='"&InfProID&"'")
InfPrice = rs("InfPrice")
InfCount = rs("InfCount")
InfSum = rs("InfPrice")*rs("InfCount")
aSum = aSum+InfSum
aNum = aNum+InfCount
aStr = aStr&InfProID&";"
%>
            <tr bgcolor="<%=col%>">
              <td align="center" nowrap="nowrap"><%=i%></td>
              <td><%=InfSubj%><input name="i<%=i%>ID" type="hidden" id="i<%=i%>ID" value="<%=KeyID%>" /></td>
              <td nowrap="nowrap"><input name='i<%=i%>Price' type='text' id="i<%=i%>Price" value="<%=InfPrice%>" size='6' maxlength='6' readonly="readonly" /></td>
              <td nowrap="nowrap"><input name='i<%=i%>Num' type='text' id="i<%=i%>Num" value="<%=InfCount%>" onblur="caFee(<%=i%>)" size='6' maxlength='6' /></td>
              <td nowrap="nowrap"><input name='i<%=i%>Sum' type='text' id="i<%=i%>Sum" value="<%=InfSum%>" size='12' maxlength='8' readonly="readonly" /></td>
              <td align="center" nowrap="nowrap"><input type="submit" name="Submit" value="<%=vOrd_m3Edit%>" onclick="ESend(<%=i%>)" /></td>
              <td align="center" nowrap="nowrap"><input type="button" name="Button" value="<%=vOrd_m3Del%>" onclick="Del_YN('?Act=PDel&amp;kID=<%=KeyID%>','<%=vOrd_m3Del%>:\n<%=Show_jsStr(InfSubj)%>?')" <%=dis1%> /></td>
            </tr>
            <% 
rs.MoveNext 
if rs.eof then exit for 
     next 
   end if ' //////////////////////////////////////////
rs.Close()
%>
            <tr bgcolor="f0f0f0">
              <td align="center" nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;<%=vOrd_m3Total%>&nbsp;</td>
              <td bgcolor="#FFFFF0">&nbsp;</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;</td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><input name='aNum' type='text' id="aNum" value="<%=aNum%>" size='6' maxlength='6' readonly="readonly" /></td>
              <td nowrap="nowrap" bgcolor="#FFFFF0"><input name='aSum' type='text' class="ItmRight" id="aSum" value="<%=aSum%>" size='12' maxlength='8' readonly="readonly" /></td>
              <td align="center" nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;</td>
              <td align="center" nowrap="nowrap" bgcolor="#FFFFF0">&nbsp;</td>
            </tr>
          </table>
        </fieldset>
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg"><%=vOrd_m4Send%></legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center"><%=vOrd_m4Send%></td>
              <td><%

sql = "SELECT * FROM [OrdPara] "
sql = sql & " WHERE KeyMod='SendType' "
sql = sql & " ORDER BY InfTop "
rs.Open sql,conn,1,1 
Do While NOT rs.EOF 

KeyID = rs("KeyID")
InfSubj = Show_Form(rs("InfSubj"))
InfCont = Show_Form(rs("InfCont"))

InfNum = rs("InfNum")
InfTop = rs("InfTop")
InfState = rs("InfState")
InfPic = rs("InfPic")

InfNum = RequestSafe(InfNum,"N",0)
If InfState="Y" Then
  InfNum1 = InfNum
  ChkFlag = "checked"
Else
  ChkFlag = " "
End If

%>
                <input type="radio" name="r1" id="radio3" value="<%=KeyID%>" <%=ChkFlag%> onclick="fOrdSType('<%=InfNum%>','v1')"/>
                <%=InfCont%><br />
                <% 
rs.MoveNext 
Loop 

rs.Close()

%></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center"><%=vOrd_m4Time%></td>
              <td><%
sql = "SELECT * FROM [OrdPara] "
sql = sql & " WHERE KeyMod='SendTime' "
sql = sql & " ORDER BY InfTop "
rs.Open sql,conn,1,1 
Do While NOT rs.EOF 

KeyID = rs("KeyID")
InfSubj = Show_Form(rs("InfSubj"))
InfCont = Show_Form(rs("InfCont"))

InfNum = rs("InfNum")
InfTop = rs("InfTop")
InfState = rs("InfState")
InfPic = rs("InfPic")

InfNum = RequestSafe(InfNum,"N",0)
If InfState="Y" Then
  InfNum2 = InfNum
  ChkFlag = "checked"
Else
  ChkFlag = " "
End If

%>
                <input type="radio" name="r2" id="radio" value="<%=KeyID%>" <%=ChkFlag%> onclick="fOrdSType('<%=InfNum%>','v2')"/>
                <%=InfCont%><br />
                <% 
rs.MoveNext 
Loop 

rs.Close()

'Member Info....
If Session("MemID")&""<>"" Then
rs.Open "SELECT * FROM [Member"&Mem_aMemb&"] WHERE MemID='"&Session("MemID")&"'",conn,1,1 
if NOT rs.eof then  
MemName = rs("MemName")
MemNam2 = rs("MemNam2")
'MemSex = rs("MemSex")
MemFrom = rs("MemFrom")
MemMobile = rs("MemMobile")
MemTel = rs("MemTel")
MemEmail = rs("MemEmail")
MemExp = rs("MemExp")
MemFlag = rs("MemFlag")
end if 
rs.Close()
End If

%></td>
            </tr>
          </table>
        </fieldset>
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg"><%=vOrd_m4Link%></legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center"><%=vOrd_m4Name%></td>
              <td width="30%"><input name="LnkName" type="text" id="LnkName" value="<%=MemName%>" size="18" maxlength="24" /></td>
              <td width="20%" align="center"><%=vOrd_m4Title%></td>
              <td width="30%"><label>
                  <select name="LnkSex" id="LnkSex">
                    <option value="<%=vOrd_m4Man%>"><%=vOrd_m4Man%></option>
                    <option value="<%=vOrd_m4Female%>" selected="selected"><%=vOrd_m4Female%></option>
                  </select>
                </label></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center"><%=vOrd_m4Addr%></td>
              <td colspan="3"><input name="LnkAddr" type="text" id="LnkAddr" style="width:480px" value="<%=MemFrom%>" size="60" maxlength="60" /></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center"><%=vOrd_m4Post%></td>
              <td><input name="LnkPost" type="text" id="LnkPost" size="18" maxlength="6" /></td>
              <td align="center"><%=vOrd_m4Mobile%></td>
              <td><input name="LnkMobile" type="text" id="LnkMobile" value="<%=MemMobile%>" size="18" maxlength="24" /></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center"><%=vOrd_m4Tel%></td>
              <td><input name="LnkTel" type="text" id="LnkTel" value="<%=MemTel%>" size="18" maxlength="24" /></td>
              <td align="center">E-mail</td>
              <td><input name="LnkEmail" type="text" id="LnkEmail" value="<%=MemEmail%>" size="18" maxlength="60" /></td>
            </tr>
          </table>
        </fieldset>
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg"><%=vOrd_m5Pays%></legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <%
sql = "SELECT * FROM [OrdPara] "
sql = sql & " WHERE KeyMod='PayType' "
sql = sql & " ORDER BY InfTop "
rs.Open sql,conn,1,1 
Do While NOT rs.EOF 

KeyID = rs("KeyID")
InfSubj = Show_Form(rs("InfSubj"))
InfCont = Show_Text(rs("InfCont"))

InfNum = rs("InfNum")
InfTop = rs("InfTop")
InfState = rs("InfState")
InfPic = rs("InfPic")
If InfPic="" Then
  InfPic = Config_Path&"img/tool/no_pic_160120.jpg"
Else
  InfPic = InfPic
End If

InfNum = RequestSafe(InfNum,"N",0)
If InfState="Y" Then
  InfNum3 = InfNum
  ChkFlag = "checked"
Else
  ChkFlag = " "
End If
                      

%>
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center"><input type="radio" name="r3" id="radio2" value="<%=KeyID%>" <%=ChkFlag%> onclick="fOrdSType('<%=InfNum%>','v3')"/>
                <%=InfSubj%></td>
              <td><%=InfCont%></td>
              <td width="30%" align="center"><img src="<%=InfPic%>" alt="<%=InfSubj%>" width="150" height="80" border="0" onload="javascript:setImgSize(this);" /></td>
            </tr>
            <% 
rs.MoveNext 
Loop 

rs.Close()

%>
          </table>
        </fieldset>
        <fieldset class="OrdFSet">
          <legend class="OrdFLeg"><%=vOrd_m5Send%></legend>
          <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF">
              <td width="20%" align="center"><%=vPMsg_ChkCode%></td>
              <td><img src="../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')"/></td>
              <td align="center"><input name="ChkCode" type="text" id="ChkCode" size="8" maxlength="12" xonfocus="javascript:PicReLoad('../');"/></td>
              <td width="40%" align="center"><label>
                  <input name="Act" type="hidden" id="Act" value="OAdd" />
                  <input name="MD" type="hidden" id="MD" value="<%=MD%>" />
                  <input name="i" type="hidden" id="i" value="0" />
                  <input type="button" name="button" id="button" value="<%=vOrd_m5SendOK%>" onclick="ChkFData(document.fmOrd01)" <%=dis2%> />
                <span class="OrdButton"> 
                &nbsp;
                <input type="button" name="button2" id="button2" value="<%=vOrd_m1View%>" onclick="location.href='info.asp?ModID=PicS124';" />
                </span></label></td>
            </tr>
          </table>
        </fieldset>
        <!--Item End--></td>
    </tr>
  </form>
</table>
<script type="text/javascript">

var fmObj = document.fmOrd01;
var v1 = "<%=InfNum1%>";
var v2 = "<%=InfNum2%>";
var v3 = "<%=InfNum3%>";
var vDef = Number(fmObj.aSum.value) + Number(v1) + Number(v2) + Number(v3);

function ESend(n)
{
  fmObj.Act.value = "PEdit";
  fmObj.i.value = n;
}

function caFee(n)
{
  var oPrice = getElmID("i"+n+"Price");
  var oNum = getElmID("i"+n+"Num");
  var oSum = getElmID("i"+n+"Sum");
  
  if (isNaN(oNum.value)||(oNum.value<=0)||(oNum.value.indexOf(".")>=0)) {
    oNum.value = 1; 
	alert("<%=vPic_jsInt%>");
  }
  var v = oNum.value*oPrice.value;
  oSum.value = v.toFixed(2);
}


fmObj.InfNum.value = vDef; //alert(vDef);

function fOrdSType(vItem,vNow)
{
  var tNum = vItem; 
  tNum = vDef - Number(eval(vNow)) + Number(tNum);
  tNum = tNum.toFixed(2);
  fmObj.InfNum.value = tNum;
}

function ChkFData(fmName)
{
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
		 
 if (fmName.LnkName.value.length==0) 
   {   
     alert("<%=vOrd_jsName%>"); 
     fmName.LnkName.focus();
     eflag = 1; break;
   }
 if (fmName.LnkAddr.value.length==0) 
   {   
     alert("<%=vOrd_jsAddr%>"); 
     fmName.LnkAddr.focus();
     eflag = 1; break;
   }
 if ((fmName.LnkTel.value.length==0)&&(fmName.LnkMobile.value.length==0))
   {   
     alert("<%=vOrd_jsTel%>"); 
     fmName.LnkTel.focus();
     eflag = 1; break;
   }
		 
 if (fmName.InfSubj.value.length==0) 
   {   

     alert("<%=vOrd_jsSubj%>"); 
     fmName.InfSubj.focus();
     eflag = 1; break;
   }

 if (fmName.InfCont.value.length>120) 
   {   
     alert("<%=vOrd_jsNote%>"); 
     fmName.InfCont.focus();
     eflag = 1; break;
   }
   
 if (fmName.ChkCode.value.length==0) 
   {   
     alert("<%=vOrd_jsCode%>"); 
     fmName.ChkCode.focus();
     eflag = 1; break;
   }
   
   // LnkTel JobName,InfRem1
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ fmName.submit(); }
}

</script>
</body>
</html>
