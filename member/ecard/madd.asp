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

  ReEnd = Request("ReEnd")
  CrdType = RequestS("CrdType",C,48)

If Request("send")="ins" Then		
  CrdID = Get_AutoID(24)'RequestS("CrdID",C,48)
  CrdNO = RequestS("CrdNO",C,48)
sql = " INSERT INTO [MemCard] (" 
sql = sql& "  CrdID" 
sql = sql& ", CrdNO" 
sql = sql& ", CrdNCode" 
sql = sql& ", CrdName" 
sql = sql& ", CrdType" 
sql = sql& ", CrdTime" 
sql = sql& ", CrdSpeci" 
sql = sql& ", CrdQty"
sql = sql& ", CrdRem" 
sql = sql& ", CrdC48" 
sql = sql& ", CrdDate" 
sql = sql& ", CrdInt" 
sql = sql& ")VALUES(" 
sql = sql& "  '" & CrdID &"'" 
sql = sql& ", '" & CrdNO &"'" 
sql = sql& ", '" & RequestS("CrdNCode",C,48) &"'" 
sql = sql& ", '" & RequestS("CrdName",C,48) &"'" 
sql = sql& ", '" & CrdType &"'" 
sql = sql& ", '" & RequestS("CrdTime","D","1900-12-31") &"'" 
sql = sql& ", '" & RequestS("CrdSpeci",C,48) &"'" 
sql = sql& ", " & RequestS("CrdQty","N",1) &"" 
sql = sql& ", '" & RequestS("CrdRem",C,255) &"'" 
sql = sql& ", '" & RequestS("CrdC48",C,48) &"'" 
sql = sql& ", '" & RequestS("CrdDate","D","2009-02-17") &"'" 
sql = sql& ", " & RequestS("CrdInt","N",1) &"" 
sql = sql& ")"

   sqlE = "SELECT CrdNO FROM [MemCard] WHERE CrdID='"&CrdNO&"' "
 If CrdNO="" Then
   Msg = "增加失败，数据不规范！"
 ElseIf rs_Exist(conn,sqlE) = "EOF" Then
   'Response.Write sql
   Call rs_DoSql(conn,sql)
   Msg = "增加成功！"
   'Response.Write js_Alert(msg,"Redir","?")
 Else
   Msg = "增加失败，证书编号已经存在！"
 End If
 
  If ReEnd="Y" Then
    Response.Write js_Alert(Msg,"Redir","madd.asp?KW="&KW&"&PG="&PG&"&CrdType="&CrdType&"&ReEnd="&ReEnd&"") 
  Else
    Response.Write js_Alert(Msg,"Redir","member.asp?KW="&KW&"&PG="&PG&"&CrdType="&CrdType&"&ReEnd="&ReEnd&"")   
  End If
 
End If
'Get_rsOpt2(conn,sqli,MemGrade)

'DefID = "Crd"&Year(Now)&"_"&Get_FmtID("mdhnsx","")

%>



        <br>
        <table width="540" align="center" cellpadding="2" cellspacing="1">
          <form name='fm01' method='post' action='?'>
		  
            <tr>
              <td colspan="2" bgcolor="#999999"></td>
            </tr>
            <tr bgcolor="f4f4f4">
              <td height="24" colspan="2"><strong>&nbsp;查询增加</strong> <font color="#FF0000"><%=Msg%></font></td>
            </tr>
            <tr>
              <td>身份证编号</td>
              <td><input name='CrdNO' type='text' value="<%=DefID%>" size='24' maxlength='24'>
              (查询編號)</td>
            </tr>
            <tr>
              <td>科室</td>
              <td>
              
      <select name="CrdType" id="CrdType" style="width:120px; ">
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
                <option value="男">[男]</option>
                <option value="女">[女]</option>
              </select>                <input name='xCrdNCode' type='hidden' id="xCrdNCode" value="<%=CrdNCode%>" size='24' maxlength='24'></td>
            </tr>
            <tr>
              <td>生日</td>
              <td><input name='CrdTime' type='text' id="CrdTime" value="<%=Date()%>" size='24' maxlength='48'></td>
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
              <td><a href="../madmin/ma_list.asp">返回</a></td>
              <td bgcolor="f4f4f4"><input name="ReEnd" type="radio" id="ReEnd1" value="N" <%If ReEnd="N" Then Response.Write("checked")%>>
添加资料后返回列表
      &nbsp;&nbsp;&nbsp;
      <input type="radio" name="ReEnd" id="ReEnd2" value="Y" <%If ReEnd="Y" Then Response.Write("checked")%>>
添加资料后继续</td>
            </tr>
            <tr bgcolor="f4f4f4">
              <td><%=EncPW%></td>
              <td bgcolor="f4f4f4"><input type='button' name='Button' value='提交' onClick="chkData()">
&nbsp;&nbsp;<a href="member.asp">返回</a> &nbsp;
<input type='reset' name='Reset' value='重设'>
                  <input name='send' type='hidden' id='send' value='ins'> </td>
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