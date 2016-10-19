<!--#include file="config.asp"-->
<!--#include file="../func1/md5_func.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../func1/Func_JS.js" type="text/javascript"></script>
<script src="../func1/WinFunc.js" type="text/javascript"></script>
</head>
<body>

<%

send = Request("send")
PrmFlag = Request("PrmFlag")

If PrmFlag="(Inn)" Then
  Call Chk_Perm3("",Config_Path&"doc/") 
  PrmUID = Session("InnID")
Else
  Call Chk_Perm1("","") 
  PrmUID = Session("UsrID")
End If

if send = "send" then
  OldPW =  RequestS("OldPW",3,512)
  PW0 =  RequestS("PW0",3,48)
  PW1 =  RequestS("PW1",3,48)
  PW2 =  RequestS("PW2",3,48)
  UsrName = RequestS("UsrName",3,24)
  eStr0 = MD5_Adm(PW0&PrmUID)
  eStr1 = MD5_Adm(PW1&PrmUID)
  sql = "UPDATE [AdmUser"&Adm_aUser&"] SET UsrPW='"&eStr1&"',UsrName='"&UsrName&"' WHERE UsrID='"&PrmUID&"'"
  If OldPW = eStr0 Then
    Call rs_DoSql(conn,sql)
    Msg = "修改成功！"
	Call Add_Log(conn,PrmUID,"修改个人密码_"&"["&PrmUID&"]",PrmFlag,"[user_editpw.asp]")
  Else
    Msg = "旧密码错误，修改不成功！"
  End If
  'Response.Write "<BR>"&OldPW&"<BR>"&eStr0
else

end if
		  
  OldPW = rs_Val("","SELECT UsrPW FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&PrmUID&"'")
  UsrID  = PrmUID
  USNM  = rs_val("","SELECT UsrName FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&UsrID&"'")
  UStr  = "["&UsrID&"] "&USNM
  USIP  = Get_CIP()
  
'////////////////////////////////////////////////
GrpDef = RequestS("GrpDef","N",-1)
GrpSub = RequestS("GrpSub","C",48)
if(send="gSet") Then
  if (GrpDef=-1)OR(GrpSub="") Then
	sql = "DELETE FROM AdmPara WHERE ParCode='Grop_"&PrmUID&"'"
	Msg = "取消设置成功!"
  ElseIf GrpSub<>"" Then
	if(rs_Count(conn,"AdmPara WHERE ParCode='Grop_"&PrmUID&"'")>0) Then
	  sql = "UPDATE AdmPara SET ParText='"&GrpDef&","&GrpSub&"' WHERE ParCode='Grop_"&PrmUID&"'"
	  Msg = "修改设置成功!"
	Else
	  sql = "INSERT INTO AdmPara (ParCode,ParFlag,ParName,ParText)VALUES('Grop_"&PrmUID&"','ParGroup','user_"&PrmUID&"','"&GrpDef&","&GrpSub&"')"
	  Msg = "新增设置成功!"
	End If
  Else
	sql = "" 
  End If
  if sql<>""  Then
	Call rs_DoSql(conn,sql)
  End If
  'echo "$GrpDef,$GrpSub<br>$sql";
End If

sGroup = rs_Val(conn,"SELECT ParText FROM AdmPara WHERE ParCode='Grop_"&PrmUID&"'")
if sGroup="" Then sGroup="-1,无"
aGroup = Split(sGroup,",") 'echo $sGroup;
fGroup = rs_Val(conn,"SELECT ParText FROM AdmPara WHERE ParCode='GropSwitch'")
aGroup(0) = RequestSafe(aGroup(0),"N",-8)
%>
        <br>
        <script type="text/javascript">

function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

		  if (document.ff.UsrName.value.length==0)
           { alert('[错误]\n 请输入 姓名!');
             document.ff.UsrName.focus();
             errflag=0;
             break;
           }
		   
		  if (document.ff.PW0.value.length==0)
           { alert('[错误]\n 请输入 旧密码!');
             document.ff.PW0.focus();
             errflag=0;
             break;
           }

		  if (document.ff.PW1.value.length<6)
           { alert('[错误]\n密码格式不正确!');
             document.ff.PW1.focus();
             errflag=0;
             break;
           }
		   
		  if (document.ff.PW1.value!=document.ff.PW2.value)
           { alert('[错误]\n两次密码不一致!');
             document.ff.PW2.focus();
             errflag=0;
             break;
           }

		  if (document.ff.PW1.value==document.ff.UsrID.value)
           { alert('[错误]\n密码不能和帐号相同!');
             document.ff.PW1.focus();
             errflag=0;
             break;
           }
       
        }
          if (errflag==1)
          {    document.ff.submit()
          }
}
</script>
        <table width="540" border="0" align="center" cellpadding="8" cellspacing="5" bgcolor="f8f8f8">
          <tr>
            <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#EEEEEE">
                <form name="ff" method="post" action="user_editpw.asp">
                  <tr align="center">
                    <td height="27" colspan="2">个人密码设定 / 定制管理菜单</td>
                  </tr>
                  <tr bgcolor="999999">
                    <td colspan="2" align="right"></td>
                  </tr>
                  <tr bgcolor="F8F8F8">
                    <td colspan="2">个人密码设定: 当前用户(<font color=blue><%=UStr%></font>)</td>
                  </tr>
                  <tr>
                    <td width="30%" align="right" bgcolor="F8F8FF">管理者帐号:</td>
                    <td width="70%" bgcolor="F8F8FF"><input name="UsrID" type="text" id="UsrID" value="<%=UsrID%> [不能修改]" readonly="">
                    </td>
                  </tr>
                  <tr>
                    <td align="right" bgcolor="F8F8FF">管理者姓名:</td>
                    <td bgcolor="F8F8FF"><input name="UsrName" type="text" id="UsrName" value="<%=USNM%>" maxlength="15">
                    </td>
                  </tr>
                  <tr>
                    <td align="right" bgcolor="F8F8FF">旧密码:</td>
                    <td bgcolor="F8F8FF"><input name="PW0" type="password" id="PW0" maxlength="15"></td>
                  </tr>
                  <tr>
                    <td align="right" bgcolor="F8F8FF">新密码:</td>
                    <td bgcolor="F8F8FF"><input name="PW1" type="password" id="PW1" maxlength="15"></td>
                  </tr>
                  <tr>
                    <td align="right" bgcolor="F8F8FF">确认新密码:</td>
                    <td bgcolor="F8F8FF"><input name="PW2" type="password" id="PW2" maxlength="15"></td>
                  </tr>
                  <tr bgcolor="F8F8F8">
                    <td nowrap><font color="#FF0000">&nbsp;</font></td>
                    <td><input type="button" name="Submit" value="个人密码设定" onClick="chkData()">
                      <input name="send" type="hidden" id="send" value="send">
                      <input name="OldPW" type="hidden" id="OldPW" value="<%=OldPW%>">
                      <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
                  </tr>
          <tr bgcolor="999999">
            <td colspan="2" align="right"></td>
          </tr>
                </form>
                
          <form name="ff2" method="post" action="?">
          <tr bgcolor="F8F8F8">
            <td colspan="2">定制管理菜单: [<%=sGroup%>]</td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">管理菜单组:</td>
            <td bgcolor="F8F8FF">
            <select name='GrpDef' id="GrpDef" style="width:150px" onChange="setGrpSub()">
            
<% 



if(fGroup="Y") Then
  Response.Write "<option value=''>(...取消定制...)</option>"
  sql = "SELECT * FROM AdmPara WHERE ParFlag='ParGroup' AND ParName<'2' ORDER BY ParName, ParCode"
  s2 = " var nGroup = (n);"&vbcrlf
  s2 =s2& " var aGroup = new Array();"
  s2 =s2& " var bGroup = new Array();" 
  i=0 : dno = 100
  SET rs=Server.CreateObject("Adodb.Recordset") 
  rs.Open sql,conn,1,1 
	if NOT rs.eof then 
	Do While Not rs.EOF
	  pCode = rs("ParCode")
	  pName = rs("ParName") : pNFlg = Mid(pName,1,2)
	  pText = rs("ParText") : pName = Mid(pName,3)
	  'If(pNFlg<"20") Then
		If i=Int(aGroup(0)) Then
		  def = " selected "
		  dno = i
		Else
		  def = ""
		End If
		Response.Write vbcrlf&"<option value='"&i&"'"&def&">"&pName&"</option>"
		s2 = s2&vbcrlf&"aGroup["&i&"] = '"&pText&"';"
		s2 = s2&vbcrlf&"bGroup["&i&"] = '"&getMName(pText)&"';"
		i = i+1
	  'End If
	rs.MoveNext
	Loop
	end if 	
  rs.Close()
  SET rs=Nothing 
  s2 = Replace(s2,"(n)",i)
Else
  s2 = ""
  Response.Write "<option value='8'>(默认分组)</option>"
End If
function getMName(Mods)
  Dim a,i,s
  a = Split(Mods,",") : s="" 
  For i=0 To uBound(a)
    s = s& rs_Val(conn,"SELECT SysName FROM AdmSyst WHERE SysID='"&a(i)&"'")&","  
  Next
  getMName = s
End function
%>
            
            
            </select></td>
          </tr>
          <tr>
            <td align="right" bgcolor="F8F8FF">管理菜单模块:</td>
            <td bgcolor="F8F8FF">
            <select name='GrpSub' id="GrpSub" style="width:150px">
            <option value=''>(...取消定制...)</option>
            <% if(fGroup<>"Y") Then Response.Write Get_rsOpt(conn,"SELECT SysID,SysName FROM AdmSyst WHERE SysType='Module' ORDER BY SysTop",aGroup(1)) %>
            </select>
            </td>
          </tr>

          <tr bgcolor="F8F8F8">
            <td nowrap bgcolor="#F8F8F8"><font color="#FF0000">&nbsp;</font></td>
            <td bgcolor="#F8F8F8"><input type="submit" name="Submit" value="定制管理菜单" xxClick="chkData()">
              <a href="?">刷新</a>              <input name="send" type="hidden" id="send" value="gSet"></td>
          </tr>
          
          <tr bgcolor="999999">
            <td colspan="2" align="right"></td>
          </tr>
          
          <tr>
            <td colspan="2" align="left" bgcolor="#FFFFFF">说明：<font color="#FF0000"><b><%=Msg%></b></font><br>
              &nbsp;&nbsp;&nbsp;&nbsp;1. 左边默认打开的管理菜单,可能不是你经常使用的管理菜单, 那么你可以使用[定制管理菜单]来设置, 设置好后重新重新登陆可见效果; <br>
              &nbsp;&nbsp; &nbsp;2. 在模块较多的较大系统中,可能不同管理员管理不同模块,那么各个管理员都自定义管理菜单,此功能就显得非常有用！</td>
            </tr>
          <tr bgcolor="999999">
            <td colspan="2" align="right"></td>
          </tr>
          
        </form>
                
              </table></td>
          </tr>
        </table>

<script type="text/javascript">

<%=s2 %>

var fmDef = document.ff2.GrpDef;
var fmSub = document.ff2.GrpSub;
function setGrpSub(){
  var id = fmDef.value; //alert(id);
  if(id=="8"){ return; }
  else{ fmSub.options.length = 1; }
  if(id==""){ return; }
  var cGroup = aGroup[id].split(",");
  var dGroup = bGroup[id].split(",");
  for(i=0;i<cGroup.length;i++){
	var iText = cGroup[i].substring(3)+":"+dGroup[i];
	var iItem = new Option(iText,cGroup[i]); 
	fmSub.options.add(iItem); //Text,Value
  }
}
setGrpSub();
fmSub.value = '<%=aGroup(1) %>';

function chkData()
{   errflag=1
        for(i=0;i<1;i++)
        {  

		  if (document.ff.UsrName.value.length==0)
           { alert('[错误]\n 请输入 姓名!');
             document.ff.UsrName.focus();
             errflag=0;
             break;
           }
		   
		  if (document.ff.PW0.value.length==0)
           { alert('[错误]\n 请输入 旧密码!');
             document.ff.PW0.focus();
             errflag=0;
             break;
           }

		  if (document.ff.PW1.value.length<6)
           { alert('[错误]\n密码格式不正确!');
             document.ff.PW1.focus();
             errflag=0;
             break;
           }
		   
		  if (document.ff.PW1.value!=document.ff.PW2.value)
           { alert('[错误]\n两次密码不一致!');
             document.ff.PW2.focus();
             errflag=0;
             break;
           }

		  if (document.ff.PW1.value==document.ff.UsrID.value)
           { alert('[错误]\n密码不能和帐号相同!');
             document.ff.PW1.focus();
             errflag=0;
             break;
           }
       
        }
          if (errflag==1)
          {    document.ff.submit()
          }
}
</script>
        
</body>
</html>
