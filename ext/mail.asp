<!--#include file="../page/_config.asp"-->
<!--#include file="../upfile/sys/pcfg/link.asp"--> 
<!--#include file="../sadm/func2/func_obj.asp"-->

<%

'LinkMaster = "dgweb@dg.gd.cn" 
send = Request("send")
ChkCode = uCase(Request("ChkCode"))
SendTo = LinkMaster 
MD = "TraJ124"

If send="send" Then
  'sSubj = "在线订单" 'RequestS("sSubj","C",48)
  'SendTo = RequestS("SemdTo","C",48)
  sSubj = RequestS("sSubj","C",48) '"在线订单" '
  sType = "InfType"'RequestS("InfType","C",48)
  sCode = rs_OrdID(conn,"yyyymmdd-r3",6,3) 'RequestS("KeyCode","C",48)
  'SendTo = RequestS("SemdTo","C",48)
  hCont = Send_HCont("hCont","<!-- FlagStart -->","<!-- FlagEnd -->")
  'hCont = "开会内容" 
  'Response.Write hCont
  If sSubj<>"" Then 
    If ChkCode=Session("ChkCode")&"" Then
	  'Call Send_Mails(Config_Name,SendTo,sSubj,hCont,"T02") 'T01,T02,GB2312
	  sql = " INSERT INTO [GboSend] (" 
	  sql = sql& "  KeyID" 
	  sql = sql& ", KeyMod" 
	  sql = sql& ", KeyCode" 
	  sql = sql& ", InfType" 
	  sql = sql& ", InfSubj" 
	  sql = sql& ", InfCont"  
	  sql = sql& ", LogAddIP" 
	  sql = sql& ", LogATime" 
	  sql = sql& ")VALUES(" 
	  sql = sql& "  '" & Get_AutoID(24) &"'" 
	  sql = sql& ", '" & MD &"'" 
	  sql = sql& ", '" & sCode &"'" 'ObjID
	  sql = sql& ", '" & sType &"'" 
	  sql = sql& ", '" & sSubj &"'" 
	  sql = sql& ", '" & hCont &"'" 
	  sql = sql& ", '" & Get_CIP() &"'" 
	  sql = sql& ", '" & Now() &"'" 
	  sql = sql& ")"
	  Call rs_DoSql(conn,sql)
      Response.Write js_Alert("提交成功！","Redir","?")
	Else
      Response.Write js_Alert("认证码错误！提交失败！","Redir","?")
	End If
  End If
  'Response.Write send
End If
'Call Send_t01()
'Response.Write send&LinkMaster
'Response.Write SendTo

MDName = "站长信箱"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%>-<%=vPMsg_WName%></title>
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>

<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_01.jpg" width="10" /></td>
    <td width="120" align="center" background="../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle"><%=MDName%></td>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="210" height="30" align="left" class="SysTit02">&nbsp;</td>
          <td align="right" class="SysTit03"><%=vPMsg_WSite%><%=vPMsg_WHome%> &gt;&gt; <%=MDName%> &gt;&gt;&nbsp; <%=TPName%> &nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="450" align="left" valign="top">
		<div style="line-height:12px;">&nbsp;</div>
            <form action='?' method="post" name='fmMail' id="fmMail">
              <input type="hidden" name='send' value="send" />
              <input type="hidden" name='hCont' value="" />
              <div id="SMailID">
                <table width="700" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E0E0E0">
                  <tr>
                    <td colspan="2" bgcolor="#F0F0F0" class="line05">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="20%" align="center" bgcolor="#FFFFFF">主题名称</td>
                    <td bgcolor="#FFFFFF"><input name="sSubj" type="text" id="sSubj" size="60" maxlength="48" />
                    </td>
                  </tr>
                  <tr>
                    <td align="center" bgcolor="#FFFFFF">投递地址</td>
                    <td bgcolor="#FFFFFF"><%=LinkMaster%> &nbsp; （站长信箱）
                      <input name='SemdTo' type="hidden" id="SemdTo" value="<%=LinkMaster%>" /></td>
                  </tr>
                  <tr>
                    <td align="center" bgcolor="#FFFFFF">详细内容</td>
                    <td bgcolor="#FFFFFF"><textarea name="sCont" cols="48" rows="8" id="sCont"></textarea></td>
                  </tr>
                  <tr>
                    <td align="center" nowrap="nowrap" bgcolor="#FFFFFF">联系人</td>
                    <td bgcolor="#FFFFFF"><input name="sName" type="text" id="sName" size="18" maxlength="24" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>电话</span>
                      <input name="sTel" type="text" id="sTel" size="18" maxlength="24" /></td>
                  </tr>
                  <tr>
                    <td align="center" nowrap="nowrap" bgcolor="#FFFFFF">邮件地址</td>
                    <td bgcolor="#FFFFFF"><input name="sEmail" type="text" id="sEmail" size="18" maxlength="60" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>ＱＱ</span>
                      <input name="sQQ" type="text" id="sQQ" size="18" maxlength="24" /></td>
                  </tr>
                  <!-- FlagStart -->
                  <tr>
                    <td align="center" bgcolor="#FFFFFF">认证码</td>
                    <td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="55%" nowrap="nowrap">
                            <input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" />
                          <img src="../sadm/pcode/img_frnd.asp" alt="如果看不清楚或停留时间过长，请点击图片换" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')"/></span></td>
                          <td><input name="button" type="button" id="button" style="width:60px;" value="提   交" onclick="chkData()" />
                            &nbsp;
<input name="button2" type="reset" id="button2" style="width:60px;" value="重   置"/>
<input name="ModID" type="hidden" id="ModID" value="LinkSMaster" /></td>
                        </tr>
                      </table></td>
                  </tr>
                  <!-- FlagEnd -->
                  <tr>
                    <td colspan="2" align="center" bgcolor="#F0F0F0" class="line05">&nbsp;</td>
                  </tr>
                </table>
              </div>
            </form></td>
      </tr>
    </table></td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240">&nbsp;</td>
  </tr>
</table>
<script type="text/javascript">
var f = getElmID('fmMail');
function chkData(){
  f.hCont.value=getElmID('SMailID').innerHTML;
  if(f.sSubj.value==''){ alert('请填写 主题名称!'); f.sSubj.focus(); return false; }
  if(f.sCont.value==''){ alert('请填写 详细内容!'); f.sCont.focus(); return false; }
  if(f.ChkCode.value==''){ alert('请填写 认证码!'); f.ChkCode.focus(); return false; }
  f.submit();
}
</script>
</body>
</html>
