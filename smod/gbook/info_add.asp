<!--#include file="config.asp"-->
<!doctype html>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jquery.js"></script>
<script type="text/javascript" charset="utf-8" src="../../inc/home/jsInfo.js"></script>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtAct=mainLoad"></script>
<style type="text/css">
tr, td {
	background-color:#FFF;
}
</style>
</head>
<body>
<%

send = Request("send")
ReEnd = Request("ReEnd")
InfType = RequestS("InfType",C,24) 

Dim sys27_Rnd(10)
If send = "ins" Then
  sys27_RVal = Request(App27Random)&""
  If sys27_RVal&"" = "" Then
    Response.End()
  Else
    For i = 1 To 9
     sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
    Next
  End If
Else
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next
End If

If send="ins" Then 

ChkCode = uCase(Request.form("ChkCode"))
If PrmFlag="(Mem)" Then
 If Session("ChkCode")<>ChkCode Or ChkCode&""="" Then
  Response.Write "<h1 style='line-height:180%; text-align:center'>认证码错误！</h1>"
  set up_file = Nothing
  Response.End()
 End If
End If

KeyID = rs_AutID(conn,ModTab,"KeyID",upPart,"1","")
InfSubj = RequestS("InfSubj"&sys27_Rnd(1),C,255)
InfCont = Show_Html(RequestS("InfCont"&sys27_Rnd(2),C,24000))
InfReply = Show_Html(RequestS("InfReply",C,24000))
 If Config_Cont="DB" Then
  xxxCont = InfCont
  xxxReply = InfReply
 Else
  xxxCont = ""
  xxxReply = ""
 End If
sql = " INSERT INTO "&ModTab&" (" 
sql = sql& "  KeyID" 
sql = sql& ", KeyMod" 
sql = sql& ", KeyFlag" 
sql = sql& ", InfType" 
sql = sql& ", InfSubj" 
sql = sql& ", InfCont" 
sql = sql& ", InfReply" 
sql = sql& ", LnkName" 
sql = sql& ", LnkEmail" 
sql = sql& ", SetRead" 
sql = sql& ", SetShow" 
sql = sql& ", ImgName" 
sql = sql& ", LogAddIP" 
sql = sql& ", LogAUser" 
sql = sql& ", LogATime" 
sql = sql& ", LogEditIP" 
sql = sql& ", LogEUser" 
sql = sql& ", LogETime" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & KeyID &"'" 
sql = sql& ", '" & ModID &"'" 
sql = sql& ", '" & RequestS("KeyFlag",C,12) &"'" 
sql = sql& ", '" & InfType &"'" 
sql = sql& ", '" & InfSubj &"'" 
sql = sql& ", '" & xxxCont &"'" 
sql = sql& ", '" & xxxReply &"'" 
sql = sql& ", '" & RequestS("LnkName",C,24) &"'" 
sql = sql& ", '" & RequestS("LnkEmail",C,255) &"'" 
sql = sql& ", 0" 
sql = sql& ", 'Y'" 
sql = sql& ", '" & RequestS("ImgName",C,255) &"'" 
sql = sql& ", '" & Get_CIP() &"'" 
sql = sql& ", '" & Get_PUser(PrmFlag) &"'" 
sql = sql& ", '" & Now() &"'" 
sql = sql& ", '-'" 
sql = sql& ", '-'" 
sql = sql& ", '1900-12-31'" 
sql = sql& ")"
 If Len(InfSubj)>0 And Len(InfCont)>0 Then
  Call rs_DoSql(conn,sql) 
  upPath = upRoot&Replace(KeyID,"-","/") 
  Call add_sfFile()
  jsMsg = "信息填写成功！" 
 Else
  jsMsg = "信息填写失败！" 
 End If
 
 ReDir = "info_list.asp?PrmFlag="&PrmFlag&""
 If ReEnd="Y" Then ReDir="info_add.asp?ReEnd="&ReEnd&"&InfType="&InfType&"&InfTyp2="&InfTyp2&"&PrmFlag="&PrmFlag&""
 Response.Write js_Alert(jsMsg,"Redir",ReDir) 
 
End If


LnkName = RequestS("LnkName","C",48)
If (ModID="MemB224" AND LnkName="") OR (ModID="MemB524" AND LnkName="") Then
  LnkName = "(Public)"
ElseIf ModID="MemB224" AND LnkName<>"" Then
  LnkName = LnkName 
ElseIf PrmFlag = "(Inn)" Then
  LnkName = Session("InnID")
ElseIf PrmFlag = "(Mem)" Then
  LnkName = Session("MemID")
Else
  LnkName = Session("UsrID") 
End If

If (ModID="MemB224" AND PrmFlag="(Mem)") OR (ModID="MemB524" AND PrmFlag="(Inn)") Then
  Response.Write "<h1 style='line-height:180%; padding:50px'>系统通知 不能增加</h1>"
  Response.End()
End If

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr>
      <td align="center"><p>&nbsp;</p></td>
      <td align="right"><strong>信息增加</strong></td>
    </tr>
    <tr>
      <td align="center" nowrap>主题</td>
      <td><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="60" maxlength="120"></td>
    </tr>
    <tr>
      <td align="center">类别</td>
      <td><select name="InfType" style="width:120px; " id="InfType">
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&ModID&"' ORDER BY TypID",InfType)%>
        </select>
        &nbsp; &nbsp;
        姓名
        <input name="LnkName" type="text" id="LnkName" value="<%=LnkName%>" size="24" maxlength="24" <%If ModID="Gbo936" Then Response.Write("readonly") %> ></td>
    </tr>
    <tr>
      <td align="center" nowrap>审核</td>
      <td><select name="SetShow" id="SetShow" style="width:120px;">
          <option value="Y"  <%if SetShow="Y" then Response.Write("selected")%>>已审核</option>
          <option value="N"  <%if SetShow="N" then Response.Write("selected")%>>未审核</option>
        </select>
        &nbsp; &nbsp; 邮箱
        <input name="LnkEmail" type="text" id="LnkEmail" value="<%=LnkEmail%>" size="24" maxlength="60"></td>
    </tr>
    <tr>
      <td align="center">内容</td>
      <td><div style="width:580px">
      
<textarea id="InfCont<%=sys27_Rnd(2)%>" name="InfCont<%=sys27_Rnd(2)%>" style="width:580px;height:360px;visibility:hidden;display:none"></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID01<%=sys27_Rnd(2)%>&edtCont=InfCont<%=sys27_Rnd(2)%>&edtTool=Basic"></script>

        </div></td>
    </tr>
    <%If (PrmFlag="(Mem)" OR PrmFlag="(Inn)") AND ModID="MemB124" Then%>
    <%Else%>
    <tr>
      <td align="center">回复</td>
      <td><div style="width:580px">
      
<textarea id="InfReply" name="InfReply" style="width:580px;height:240px;visibility:hidden;display:none"></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID02&edtCont=InfReply&edtTool=Basic"></script>

      </div></td>
    </tr>
    <%End If%>
    <tr>
      <td align="center" nowrap>返回</td>
      <td><input name="ReEnd" type="radio" id="ReEnd1" value="N" <%If ReEnd="N" Then Response.Write("checked")%>>
        添加资料后返回列表
        &nbsp;&nbsp;&nbsp;
        <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
        添加资料后继续</td>
    </tr>
    
    <%If PrmFlag="(Mem)" Then%>
    <tr>
      <td align="center" nowrap>认证码</td>
      <td><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" xonfocus="javascript:PicReLoad('../../');"/>
        <img src="../../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../../')"/></td>
    </tr>
    <%End If%>
    
    <tr>
      <td align="center" nowrap><input name="ModID" type="hidden" id="ModID" value="<%=ModID%>"></td>
      <td><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="send" type="hidden" id="send" value="ins">
        <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>">
        <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
        注意:留言12K字内 </td>
    </tr>
  </form>
</table>
<script type="text/javascript">

 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj<%=sys27_Rnd(1)%>.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     document.fm01.InfSubj<%=sys27_Rnd(1)%>.focus();
     eflag = 1; break;
   }
 document.fm01.InfCont<%=sys27_Rnd(2)%>.value = apiGetValEditID01<%=sys27_Rnd(2)%>();
 <%If NOT ((PrmFlag="(Mem)" OR PrmFlag="(Inn)") AND ModID="MemB124") Then%>
 document.fm01.InfReply.value = apiGetValEditID02();
 <%End If%>
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length>=12000) 
   {   
     alert(" 内容 不能超过 12K字!"); 
	 document.fm01.InfCont<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }


         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
</body>
</html>
