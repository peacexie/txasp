<!--#include file="config.asp"-->

<%

t1 = Timer()
Call Chk_URL()
Call Chk_Perm1("","") 

If Request("send")="OutMember" Then
  Session("MemPerm")=""
  Session("MemID")=""
ElseIf Request("send")="OutInner" Then 
  Session("InnPerm")=""
  Session("InnID")=""
End If


' User Info //////////////////////////////////////////
UsrID  = Session("UsrID")
UsrNM  = rs_val("","SELECT UsrName FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&UsrID&"'")
UsrStr  = "["&Session("UsrID")&"] "&UsrNM
UsrIP  = Get_CIP()

''X^ remind.js
'scRem = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='remind.js'")
'scF1 = inStr(scRem,"X^") ': Response.Write scF1
'scF2 = Config_Mode ': Response.Write scF2
''tmp_File.asp文件模板:sTmpFile_Flag = false
'scTFile = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='tmp_File.asp'")
'scF3 = inStr(scTFile ,"sTmpFile_Flag = false") ': Response.Write scF3
''rKeyID.asp  [ParRem]保留字,刷新项
'scKeyID = rs_Val("","SELECT ParRem FROM AdmPara WHERE ParCode='rKeyID.asp'")
'scF4 = inStr(scKeyID,"Docs;") ': Response.Write scF4
'Response.Write Request.Servervariables("HTTP_REFERER")

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>后台管理中心</title>
<link rel="stylesheet" href="../../inc/adm_img/style.css" type="text/css" />
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
table {
	border:1px solid #999; 
}

.songti {
	font-family:'宋体';
	color:#CCC;
}
#u1,#u2,#u3,#u4 {
    background-color:#FFFFCC;
}
#sc1,#sc2 {
    color:#FFF;
	background-color:#FF0000;
}
#main_body td{
    padding:2px 8px;
	margin:0px;
	line-height:180%;
}
</style>
</head>
<body>
<div id="main_body" style="width:600px; margin:auto">
  <ul class="main_top">
    <li class="main_top_left left"><%=Config_Name%> 后台管理中心 --- 管理首页</li>
    <li class="main_top_right right"> </li>
  </ul>
  <div class="main_content_rightbg">
    <div class="main_content_leftbg">
      <table  border="0" align="center" cellpadding="0" cellspacing="1">
        <tr align="center">
          <td colspan="2" align="left"><div style="float:right"> 快捷键:
             <A href="../../smod/adupd/upd.asp" target="mainFrame">信息刷新</A>
<%
  If Config_Mode="isExpert" Then
%>
           - <A href="../../smod/adupd/upd.asp?send=TimTest" target="mainFrame">效率测试</A>
           - <A href="../../smod/adupd/upd_data.asp" target="mainFrame">数据转化</A>
<%
  End If
%> 
           - <A href="../../tools/help/xhelp.asp" target="mainFrame">管理帮助</A>
           
           </div>
          <strong><span class="songti">&#8226;</span>管理员信息：</strong>
            <!--<%=Session(UsrPStr)%>--></td>
        </tr>
        
              <!--  Or scF1+scF3+scF4>0 -->
			  <%If Config_Mode="isExpert" AND Chk_PermSP() Then%>
              <tr>
                <td id="sc1" align="right">设置检测:</td>
                <td id="sc2">
                <%
				If scF1>0 Then
				  Response.Write " remind.js | "
				End If
				If Config_Mode="isExpert" Then
				  Response.Write " Config_Mode=isExpert | "
				End If
				If scF3>0 Then
				  Response.Write " tmp_File.asp | "
				End If
				If scF4>0 Then
				  Response.Write " rKeyID.asp | "
				End If
				%>
                (是否需要设置？)
                </td>
              </tr>
              <%End If%>
        <%If Len(Session("InnID")&"")>0 Then%>
              <tr>
                <td align="right" id="u1">内部用户:</td>
                <td id="u2">&nbsp;<a href="?send=OutInner"><span class="fnt00F">退出 内部用户[<%=Session("InnID")%>]</span></a></td>
              </tr>
              <%End If%>
        <%If Len(Session("MemID")&"")>0 Then%>
              <tr id="u2">
                <td align="right" id="u3">会员用户:</td>
                <td id="u4">&nbsp;<a href="?send=OutMember"><span class="fnt00F">退出 会员用户[<%=Session("MemID")%>]</span></a></td>
              </tr>
              <%End If%>
        
        <tr>
          <td width="20%" align="right">当前用户:</td>
          <td align="left">&nbsp;<%=UsrStr%> &nbsp;( IP: <%=UsrIP%> ) </td>
        </tr>
        <!--tr>
          <td align="right">来源地址:</td>
          <td align="left" nowrap><%=Request.Servervariables("HTTP_REFERER")%></td>
        </tr-->
        <tr>
          <td align="right">服务器时间:</td>
          <td align="left" nowrap>
          <!--#include file="chk_time.asp"-->
          </td>
        </tr>
        <tr>
          <td align="right">密码设定:</td>
          <td align="left"><a href="user_editpw.asp">修改管理密码/定制管理菜单</a> (<span class="colF00">第一次使用,建议修改密码</span>)</td>
        </tr>
      </table>
      <div style="line-height:5px;">&nbsp;</div>
      <table  border="0" align="center" cellpadding="0" cellspacing="1" id="help">
        <tr>
          <td colspan="2" align="left"><strong> <span class="songti">&#8226;</span>管理帮助 &amp; 配置向导&nbsp;&nbsp;</strong> 初次使用建议仔细阅读以下标记为[<span class="colF00">重要</span>]的说明 </td>
        </tr>
        <tr>
          <td align="center"><p>帮助文件[<span class="colF00">重要</span>]</p></td>
          <td align="left"><a href="../../tools/help/xhelp.asp" target="_blank">通用帮助</a> -- <a href="../../tools/help/xhelp.asp#FlagComm" target="_blank">常用配置</a> -- <a href="../../smod/gbook/info_list.asp?ModID=GboU124">站务笔记</a> -- <a href="../../tools/out/admlogs.asp" target="_blank">导出笔记</a></td>
        </tr>
        <tr>
          <td align="center"><p>重点说明[<span class="colF00">重要</span>]</p></td>
          <td align="left"><a href="../../tools/help/xhelp.asp#FlagComm_1" target="_blank">设置栏目</a> -- <a href="../../tools/help/xhelp.asp#FlagComm_2" target="_blank">版权信息</a> -- <a href="../../tools/help/xhelp.asp#FlagComm_3" target="_blank">友情连接</a> -- <a href="../../tools/help/xhelp.asp#FlagComm_4" target="_blank">刷新说明</a></td>
        </tr>
        <tr>
          <td width="20%" align="center"><p>非必要工具</p></td>
          <td align="left"><a href="../../tools/tools.asp">补充应用</a> -- <a href="../../ext/api/scan/aspcheck.asp" target="_blank">阿江ASP 探针</a> -- <!--<a href="../../tools/base/clear.asp" target="_blank">木马扫把</a>--> [需要时谨慎操作！] </td>
        </tr>
      </table>
      <div style="line-height:5px;">&nbsp;</div>
      <table  border="0" align="center" cellpadding="0" cellspacing="1" id="biji">
        <tr>
          <td colspan="2" align="left">
          <!--div style="float:right"><a href="#">提醒设置</a></div-->
          <strong> <span class="songti">&#8226;</span>站务笔记(管理日志)</strong> -----------------------------------------------------=&gt; <a href="../../smod/gbook/info_list.asp?ModID=GboU124">更多站务笔记&gt;&gt;&gt;</a> </td>
        </tr>
        <%
 If SwhShowSpace="Y" Then
   nTop = 8
 Else
   nTop = 12
 End If
 sql = " SELECT TOP "&nTop&" * FROM [GboInfo] "
 sql =sql& " WHERE KeyMod='GboU124' " 'LogAUser='"&Session("MemID")&"' 1=1(Admin)
 sql =sql& " ORDER BY KeyID DESC"  ':Response.Write sql&SetTop
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 i = 0
 Do While Not rs.EOF
KeyID = rs("KeyID")
KeyFlag = rs("KeyFlag")
InfType = Show_Text(rs("InfType"))
InfSubj = Show_SLen(rs("InfSubj"),22) 'Show_Text(rs("InfSubj"))
If i mod 2 = 0 Then Response.Write("<tr>")
%>
          <td align="left" width="50%"><%=(i+1)%>&middot; <a href="../../smod/gbook/info_view.asp?ID=<%=KeyID%>" target="_blank"><%=InfSubj%></a></td>

        <%
If (i mod 2 = 1) and i>1 Then Response.Write("</tr>")
i = i + 1
 rs.Movenext
 Loop
 rs.Close()
 Set rs = Nothing
%>
      </table>
      <%If SwhShowSpace="Y" Or Request("SwhShowSpace")="Y" Then%>
      <div style="line-height:5px;">&nbsp;</div>
      <!--#include file="chk_space.asp"-->
      <%End If%>
    </div>
  </div>
  <ul class="main_end">
    <li class="main_end_left left"></li>
    <li class="main_end_right right"></li>
  </ul>
</div>

</body>
</html>
