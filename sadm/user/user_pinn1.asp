<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.pItem {
	width:120px;
	height:auto;
	margin:0px;
	padding:0px 2px 0px 2px;
	float:left;
	overflow:hidden;
	border-right:1px dashed #EEEEEE;
}
.pItm2 {
	width:20px;
	height:auto;
	margin:0px;
	float:left;
}
.pIRow1 {
	border-top:1px dashed #333333;
}
.pIRow2 {
	padding-bottom:5px;
}
.unLine {
    text-decoration:underline;	
}
.ovLine {
	text-decoration:overline;
}
</style>
</head>
<body>

        <%

PrmUser = RequestS("PrmUser",3,48)
ModID = RequestS("ModID",3,48)
sql = "SELECT UsrName,UsrPerm,UsrType FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&PrmUser&"'"
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
If NOT rs.EOF Then
  UsrName = rs("UsrName")
  UsrPerm = rs("UsrPerm")
  UsrType = rs("UsrType")
  sPerm = UsrPerm
End If
rs.close()
set rs = nothing


send  = RequestS("send",3,48) 
If send="send" Then
  PrmList  = RequestS("PrmList",3,8000)
  PrmArr = Split(PrmList,",")
  s = ""
For i = 0 To uBound(PrmArr)
  If PrmArr(i)&""<>"" Then
    s = s &"("& Trim(PrmArr(i)) &");"
  End If
Next
  s = "{"&UsrType&"};"&s
  Call rs_DoSql(conn,"UPDATE [AdmUser"&Adm_aUser&"] SET UsrPerm='"&s&"' WHERE UsrID='"&PrmUser&"'")
  sPerm = s
  msg = "权限修改成功!!"
End If

s = ""

'rUInner
pList = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rUInner'")
sql = " SELECT * FROM [AdmSyst] WHERE SysID IN("&pList&") ORDER BY SysTop,SysID " 
Set rs=Server.Createobject("Adodb.Recordset")
rs.Open Sql,conn,1,1
i=0
Do While NOT rs.EOF
  i=i+1
  SysID = rs("SysID")
  SysName = rs("SysName")
    ch1 = ""
  If inStr(sPerm,SysID)>0 Then
    ch1 = "checked"
  End If
  s=s& "<span class='pItem'><font style='color:#6666FF'><input type='checkbox' name='[ChkBox01]' value='"&SysID&"' "&ch1&">"&SysName&"</font></span>"

rs.MoveNext
Loop
s = Replace(s,"[ChkBox01]","PrmList")


%>
        <table width="620" border="0" align="center" cellpadding="3" cellspacing="1">
          <tr align="center">
            <td width="30%" rowspan="2" align="center"><strong>管理者权限设定</strong> <br>
              <font color="#0000FF">[<%=PrmUser%>]<%=UsrName%></font></td>
            <td align="center"><font color="#FF0000"><%=msg%></font></td>
            <td width="20%"><a href="user.asp?ModID=Inner">返回用户列表</a></td>
          </tr>
          <tr align="center">
            <td colspan="2" align="right"><a href="user_pinn1.asp?PrmUser=<%=PrmUser%>&ModID=Inner">Ver1</a></td>
          </tr>
		  <form name="fm01" action="?">
          <tr align="center">
            <td colspan="3" align="center">
              
              
              
              <table width="100%" border="1">
                <tr>
                  <td>
                  
              <table width='100%' border=0 align='center'>
           
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * 1.
                      <input name='xx0List' type='checkbox' id="xx0List" value='x0System' disabled>
                  内部公文 ----- [ModDocs]</font></td>
                </tr>

                <tr>
                  <td class='pIRow2'>
                  
                    <span class='pItem'><font style='color:#6666FF'>
                    <input type='checkbox' name='PrmList' value='DocD124' <%If inStr(sPerm,"DocD124")>0 Then Response.Write("checked")%>>
                    公文查看</font></span><span class='pItem'><font style='color:#6666FF'>
                    <input type='checkbox' name='PrmList' value='DocDPub' <%If inStr(sPerm,"DocDPub")>0 Then Response.Write("checked")%>>
                    公文发布</font></span></td>
                </tr>
                
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * 2.
                      <input name='xx1List' type='checkbox' id="xx1List" value='xx1System' disabled>
                  留言反馈 ----- [MemB*3]</font></td>
                </tr>
                <tr>
                  <td class='pIRow2'>
                  
                    <span class='pItem'><font style='color:#6666FF'>
                    <input type='checkbox' name='PrmList' value='MemB424' <%If inStr(sPerm,"MemB424")>0 Then Response.Write("checked")%>>
                    会员反馈</font></span><span class='pItem'><font style='color:#6666FF'>
                    <input type='checkbox' name='PrmList' value='MemB524' <%If inStr(sPerm,"MemB524")>0 Then Response.Write("checked")%>>
                    会员通知</font></span><span class='pItem'><font style='color:#6666FF'>
                    <input type='checkbox' name='PrmList' value='MemB624' <%If inStr(sPerm,"MemB624")>0 Then Response.Write("checked")%>>
                    个人笔记</font></span>
                  
                  </td>
                </tr>
                
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * 3.
                      <input name='xx2List' type='checkbox' id="xx2List" value='xx2System' disabled>
                  模块权限 ----- [Config_Model]</font></td>
                </tr>
                <tr>
                  <td class='pIRow2'><%=s%></td>
                </tr>
                
                <%If SwhDepSubs="Y" Then %>
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * D.
                    <input name='xx3List' type='checkbox' id="xx3List" value='xx3System' disabled>
                    科室部门 ----- [SysDepart]</font></td>
                </tr>
                <tr>
                  <td class='pIRow2'><%=CPartD()%></td>
                </tr>
                <%End If%>
                <!--InsHere-->
                <tr>
                  <td class='pIRow1'><font style='color:#6666FF'> * X.
                    <input name='xx4List' type='checkbox' id="xx4List" value='xx4System' disabled>
                    附加权限 ----- [SysAppend]</font></td>
                </tr>
                <tr>
                  <td class='pIRow2'>
                    <input type='checkbox' name='PrmList' value='SetShow' <%If inStr(sPerm,"SetShow")>0 Then Response.Write("checked")%>>
                    发布信息不需要审核(默认都需要审核)
                    &nbsp;
                    <input type='checkbox' name='PrmList' value='SetOher' <%If inStr(sPerm,"SetOher")>0 Then Response.Write("checked")%>>
                    管理查看他人信息(默认只管理查看个人的信息) 
                  </td>
                </tr>
                <!--Ins End-->
                <!--TestHere--/>
                <tr>
                  <td class='pIRow2'><=CPartT("PicS124")></td>
                </tr>
                <!--Test End-->
              </table>

                  
                  </td>
                </tr>

                <tr>
                  <td align="center"><input name="SetOK" type="submit" id="SetOK" value="确认">
                    <input name="send" type="hidden" id="send" value="send">
                    <input name="PrmUser" type="hidden" id="PrmUser" value="<%=PrmUser%>"></td>
                </tr>
            </table></td>
          </tr>
          
		  </form>
        </table>
        
<!-- ;(N110088
<%=sPerm%>
-->        
</body>
</html>
