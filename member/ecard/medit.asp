<!--#include file="config.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

send   = Request("send") 
ID = RequestS("ID",3,48)

PG = RequestS("PG","N",1)
KW = RequestS("KW",3,48)
KT = RequestS("KT",3,48)
ID = RequestS("ID",3,48)
TP = RequestS("TP",3,48)
T1 = RequestS("T1","D","1900-12-31")
T2 = RequestS("T2","D","2100-12-31")

If send = "ins" Then
sql = " UPDATE [MemCard] SET " 
sql = sql& " CrdNO = '" & RequestS("CrdNO",C,48) &"'" 
sql = sql& ",CrdType = '" & RequestS("CrdType",C,48) &"'" 
sql = sql& ",CrdName = '" & RequestS("CrdName","C",48) &"'" 
sql = sql& ",CrdNCode = '" & RequestS("CrdNCode","C",48) &"'" 
sql = sql& ",CrdTime = '" & RequestS("CrdTime","D","1900-12-31") &"'"
sql = sql& ",CrdSpeci = '" & RequestS("CrdSpeci","C",48) &"'" 
sql = sql& ",CrdQty = " & RequestS("CrdQty","N",48) &"" 
sql = sql& ",CrdRem = '" & RequestS("CrdRem","C",255) &"'" 
sql = sql& ",CrdDate = '" & RequestS("CrdDate","D","1900-12-31") &"'" 
sql = sql& ",CrdC48 = '" & RequestS("CrdC48","C",48) &"'" 
sql = sql& ",CrdInt = " & RequestS("CrdInt","N",48) &"" 
sql = sql& " WHERE CrdID='"&ID&"' "
Call rs_DoSql(conn,sql)
Response.Write js_Alert("修改成功！","Redir","member.asp?KW="&KW&"&KT="&KT&"&TP="&TP&"&T1="&T1&"&T2="&T2&"&Page="&PG&"")
'Response.Write sql
End If

sql = "SELECT "
sql = sql & " * FROM [MemCard] "
sql =sql& " WHERE CrdID='"&ID&"' "

   Set rs=Server.Createobject("Adodb.Recordset")
   rs.Open Sql,conn,1,1
if NOT rs.EOF then
CrdID = rs("CrdID")
CrdNO = rs("CrdNO")
CrdNCode = rs("CrdNCode")
CrdType = rs("CrdType")
CrdName = rs("CrdName")
CrdTime = rs("CrdTime")
CrdSpeci = rs("CrdSpeci")
CrdQty = rs("CrdQty")
CrdRem = rs("CrdRem")
CrdC48 = rs("CrdC48")
CrdDate = rs("CrdDate")
CrdInt = rs("CrdInt")

end if
  rs.close()
  set rs = nothing

If CrdTime="1900-12-31" Then
 CrdTime="" '<font color=red>无资料</font>
End If

%>



        <br>
        <table width="540" align="center" cellpadding="2" cellspacing="1">
          <form name='fm01' method='post' action='?'>
            <tr>
              <td colspan="2" bgcolor="#999999"></td>
            </tr>
            <tr bgcolor="f4f4f4">
              <td height="24" colspan="2"><strong>&nbsp;查询修改</strong> <font color="#FF0000"><%=Msg%></font></td>
            </tr>
            <tr>
              <td>身份证编号</td>
              <td><input name='CrdNO' type='text' value="<%=CrdNO%>" size='24' maxlength='24' Xreadonly>
              (查询編號)</td>
            </tr>
            <tr>
              <td>科室</td>
              <td><select name="CrdType" id="CrdType" style="width:120px; ">
                <option value="">[选择类别]</option>
				<%=Get_rsOpt(conn,"SELECT TypName AS TypID,TypName FROM WebTyps WHERE TypMod='MemC124' ORDER BY TypName ",CrdType)%>
              </select></td>
            </tr>
            <tr>
              <td>姓名</td>
              <td><input name='CrdName' type='text' value="<%=CrdName%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td>性别</td>
              <td><select name="CrdNCode" id="CrdNCode" style="width:120px; ">
                <option value="">[性别]</option>
                <option value="男" <%If CrdNCode="男" Then Response.Write("selected")%>>[男]</option>
                <option value="女" <%If CrdNCode="女" Then Response.Write("selected")%>>[女]</option>
              </select>                <input name='xCrdNCode' type='hidden' id="xCrdNCode" value="<%=CrdNCode%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td>生日</td>
              <td><input name='CrdTime' type='text' id="CrdTime" value="<%=CrdTime%>" size='24' maxlength='48'></td>
            </tr>
            <tr>
              <td>头衔</td>
              <td><input name='CrdSpeci' type='text' id="CrdSpeci" value="<%=CrdSpeci%>" size='24' maxlength='48'>              </td>
            </tr>
            <!--
            <tr>
              <td>数量</td>
              <td><input name='CrdQty' type='text' id="CrdQty" value="<%=CrdQty%>" size='24' maxlength='48'>              </td>
            </tr>
            <tr>
              <td>备用C48</td>
              <td><input name='CrdC48' type='text' id="CrdC48" value="<%=CrdC48%>" size='24' maxlength='48'></td>
            </tr>
            <tr>
              <td>备用时间</td>
              <td><input name='CrdDate' type='text' id="CrdDate" value="<%=CrdDate%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td>备用数字</td>
              <td><input name='CrdInt' type='text' id="CrdInt" value="<%=CrdInt%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td>备注<br>
                (120字内)</td>
              <td><textarea name="CrdRem" cols="48" rows="4" id="CrdRem"><%=CrdRem%></textarea></td>
            </tr>
            -->
            <tr bgcolor="f4f4f4">
              <td>&nbsp;</td>
              <td bgcolor="f4f4f4"><input type='button' name='Button' value='提交' onClick="chkData()">
&nbsp;&nbsp;<a href="member.asp">返回</a> &nbsp;
        <input type='reset' name='Reset' value='重设'>
        <input name='send' type='hidden' id='send' value='ins'>              <input name='ID' type='hidden' id='ID' value='<%=ID%>'>
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="KT" type="hidden" id="KT" value="<%=KT%>">
<input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="T1" type="hidden" id="T1" value="<%=T1%>">
        <input name="T2" type="hidden" id="T2" value="<%=T2%>"></td>
            </tr>
          </form>
          <tr>
            <td colspan="2" bgcolor="#999999"></td>
          </tr>
        </table>
		
<script type="text/javascript">


 function chkData()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////

 if (document.fm01.CrdName.value.length==0) 
   {   
     alert(" 姓名 不能为空！"); 
     document.fm01.CrdName.focus();
     eflag = 1; break;
   }
   
 if (document.fm01.CrdType.value.length==0) 
   {   
     alert(" 科室 不能为空！"); 
     document.fm01.CrdType.focus();
     eflag = 1; break;
   }
   
 if (document.fm01.CrdNCode.value.length==0) 
   {   
     alert(" 性别 不能为空！"); 
     document.fm01.CrdNCode.focus();
     eflag = 1; break;
   }

 if (document.fm01.CrdNO.value.length<15) 
   {   
     alert("身份证编号 至少15位");
	 document.fm01.CrdNO.focus();
	 eflag = 1; break;
   }

 tmv = chkF_Date(document.fm01.CrdTime,"日期 不规范！");
 if (tmv=='ER') { eflag = 1; break; }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
		
		
</body>
</html>