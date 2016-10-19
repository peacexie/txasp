<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="config.asp"-->
<script src="/sadm/func1/Func_JS.js" type="text/javascript"></script>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>

<%

Set rs=Server.Createobject("Adodb.Recordset")

yAct = Request("yAct") 
Page = RequestS("Page","N",1)
KW = RequestS("KW",3,24)
TP = RequestS("TP",3,24)
ID = RequestS("ID",3,48)
OD = RequestS("OD",3,48)
If ID&"" <> "" Then
  sqlK = " AND (VoteLogs.KeyMod='"&ID&"') " 
End If
If KW&"" <> "" Then
 If TP="IP" Then
  sqlK = " AND (VoteLogs.LogAddIP LIKE '%"&KW&"%') " 
 ElseIf TP="SJ" Then 
  sqlK = " AND (VoteInfo.InfSubj LIKE '%"&KW&"%') " 
 ElseIf TP="NM" Then 
  sqlK = " AND (VoteLogs.InfName LIKE '%"&KW&"%') " 
 ElseIf TP="TL" Then 
  sqlK = " AND (VoteLogs.InfTel LIKE '%"&KW&"%') " 
 ElseIf TP="CD" Then 
  sqlK = " AND (VoteLogs.InfCard LIKE '%"&KW&"%') " 
 Else
  sqlK = " AND (VoteLogs.LogAddIP LIKE '%"&KW&"%') " 
 End If
End If

cID = 0
sID = ""

If yAct="del_sel" then
  cID=0
  For iy = 1 To Request.Form("yID").Count
    cID=cID+1
	iID = RequestSafe(Request.Form("yID").item(iy),3,96) 
    Call rs_DoSql(conn,"DELETE FROM VoteLogs WHERE KeyID='"&iID&"'")
  Next
    Msg = cID&" 条记录删除成功!"
ElseIf yAct="v_Cancel" AND ID<>"" Then '取消贺奖
  For iy = 1 To Request.Form("yID").Count
    Call rs_DoSql(conn,"UPDATE VoteLogs SET InfLuck='X' WHERE KeyID='"&Request.Form("yID").item(iy)&"'")
  Next
  Msg = " 取消贺奖成功!"
ElseIf yAct="d_Clear" Then '清理数据
  Call rs_DoSql(conn,"DELETE FROM VoteLogs WHERE KeyMod NOT IN(SELECT KeyID FROM VoteInfo)")
  Call rs_DoSql(conn,"DELETE FROM VoteItem WHERE KeyMod NOT IN(SELECT KeyID FROM VoteInfo)")
  Msg = " 清理数据成功!"
ElseIf yAct="r_Vote" AND ID<>"" Then '重计选票 
  Call ReVote(ID)
  Msg = " 重计选票成功!"
ElseIf yAct="r_Rate" AND ID<>"" Then '重计吻合度
  nTop = rs_Val("","SELECT InfNum2 FROM VoteInfo WHERE KeyID='"&ID&"'") 
  Call ReRate(ID,nTopID(ID,nTop),nTop)
  Msg = " 重计吻合度成功!"
ElseIf yAct="fLuck" AND ID<>"" Then '自动抽奖    
  vOrd = RequestS("vOrd","C",2) ' Nul,0,1,2,3,4
  vNum = RequestS("vNum","N",1) ' 1,2,3...
  vDat = RequestS("vDat","C",4) ' Nul,05,10,20,AA
  If vOrd<>"" And vDat<>"" Then
    If vDat="AA" Then vDat="99"
	nLuck = SetLucks(ID,vDat,vNum,vOrd)
	Msg = ""&nLuck&"条记录 自动抽奖成功!"
  Else
    Msg = " 抽奖设置错误？！"
  End If
End If
    sql = " SELECT VoteLogs.*,VoteInfo.InfSubj FROM [VoteLogs],[VoteInfo] "
	sql =sql& " WHERE VoteLogs.KeyMod=VoteInfo.KeyID "&sqlK
	sqlO = " VoteLogs.KeyID DESC "
	If OD="InfRate" Then sqlO = " VoteLogs.InfRate DESC "
	If OD="InfLuck" Then sqlO = " VoteLogs.InfLuck ASC "
	sql =sql& " ORDER BY "&sqlO&"" ': Response.Write sqlk'SetTop,
   
   rs.Open Sql,conn,1,1
   rs.PageSize = 15 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If

'Call rs_DoSql(conn,"UPDATE VoteInfo SET SetShow='Y' ")

%>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="E0E0E0">
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="9" align="center" bgcolor="f8f8f8"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr align="center" bgcolor="#FFFFFF">
          <td width="210" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>[投票记录]管理</strong></td>
          <td width="210" rowspan="2" align="center" bgcolor="#FFFFFF"><font color="#FF0000"><%=msg%></font></td>
          <form name="fm01" method="post" action="?">
            <td width="50%" align="right" nowrap bgcolor="#FFFFFF"><input name="yAct" type="hidden" id="yAct" value="fLuck">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
              抽奖: 
              <select class=form id=vOrd name=vOrd style="width:90px; ">
                <option value="">奖项名</option>
                <option value="0">特等奖</option>
                <option value="1">一等奖</option>
                <option value="2">二等奖</option>
                <option value="3">三等奖</option>
                <option value="4">参与奖</option>
              </select>
              <input name="vNum" type="text" id="vNum" style="width:35px;" value="3" maxlength="2">
              个(贺奖名额)
              <select class=form id=vDat name=vDat style="width:90px; ">
                <option value="">[数据来源]</option>
                <option value="05">前05个吻合度</option>
                <option value="10">前10个吻合度</option>
                <option value="20">前20个吻合度</option>
                <option value="AA">[ 所有 ]</option>
              </select>
            <input type="submit" name="Submit2" value="执行" <%If ID="" Then Response.Write("disabled")%>></td>
          </form>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
<form name="fm02" method="post" action="?">
            <td width="50%" align="right" nowrap bgcolor="#FFFFFF">&nbsp;
              <input name="send" type="hidden" id="send" value="sch">
              <input name="ID" type="hidden" id="ID" value="<%=ID%>">
              搜索:
              <select class=form id=select name=TP style="width:90px; ">
			  <option value="">选择</option>
              <option value="IP">[IP]</option>
              <option value="SJ">主题</option>
              <option value="NM">姓名</option>
              <option value="TL">电话</option>
              <option value="CD">证件</option>
              </select>
              <input name="KW" type="text" id="KW" value="<%=KW%>" size="12" maxlength="12">
              <select class=form id=OD name=OD style="width:90px; ">
                <option value="LogATime">[排序]</option>
                <option value="LogATime">时间</option>
                <option value="InfRate">吻合度</option>
                <option value="InfLuck">贺奖名次</option>
              </select>
              <input type="submit" name="Submit" value="搜索">            </td>
          </form>
        </tr>
      </table></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" colspan="9" align="center" bgcolor="f8f8f8"><%= RS_Page(rs,Page,"?send=pag&TP="&TP&"&KW="&KW&"",1)%></td>
  </tr>
  <tr align="center" bgcolor="E0E0E0">
    <td height="27" align="center">NO</td>
    <td width="30%" height="27" align="center" bgcolor="E0E0E0">主题</td>
    <td align="center" nowrap>贺奖</td>
    <td align="center" nowrap>吻合度</td>
    <td height="27" align="center">时间</td>
    <td align="center">IP</td>
    <td align="center">姓名</td>
    <td align="center">电话</td>
    <td height="27" align="center">身份证号</td>
  </tr>
  <tr bgcolor="#333333">
    <td colspan="11" align="right" nowrap></td>
  </tr>
  <form name="flist" method="post" action="?">
    <%
  If not rs.eof then
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfSubj = rs("InfSubj")
InfName = rs("InfName")
InfCard = rs("InfCard")
InfTel = rs("InfTel")
InfRate = rs("InfRate")
InfLuck = rs("InfLuck")&""
LogAddIP = rs("LogAddIP")
LogATime = rs("LogATime")
If NOT isNumeric(InfRate) Then InfRate=0
InfRate = FormatNumber(InfRate,4)
If InfLuck="0" Then
 sLuck = "<font color='#FF00FF'>特等奖</font>"
ElseIf InfLuck="1" Then
 sLuck = "<font color='#FF0000'>一等奖</font>"
ElseIf InfLuck="2" Then
 sLuck = "<font color='#CC0000'>二等奖</font>"
ElseIf InfLuck="3" Then
 sLuck = "<font color='#990000'>三等奖</font>"
ElseIf InfLuck="4" Then
 sLuck = "<font color='#999999'>参与奖</font>"
Else
 sLuck = "<font color='#CCCCCC'>"&InfLuck&"</font>"
End If
If KeyMod="BBSVA124" Then
 sUrl = "vote.asp"
Else
 sUrl = "research.asp"
End If
	  %>
    <tr bgcolor="<%=col%>">
      <td align="right"><%=i%>
      <input 
			name="yID" type="checkbox" id="yID" value="<%=KeyID%>"></td>
      <td><a href="../../vote/<%=sUrl%>?ID=<%=KeyMod%>" target="_blank"><%=InfSubj%></a>      </td>
      <td align="center" nowrap><%=sLuck%></td>
      <td align="center" nowrap><%=InfRate%></td>
      <td align="center" nowrap bgcolor="<%=col%>" class="siz009L"><%=LogATime%></td>
      <td align="center" bgcolor="<%=col%>" class="siz009L"><%=LogAddIP%></td>
      <td align="center" nowrap><%=InfName%></td>
      <td align="center"><%=InfTel%></td>
      <td align="center" class="siz009L"><%=InfCard%></td>
    </tr>
    <%
  rs.Movenext
  If rs.Eof Then Exit For

  Next
%>
    <tr bgcolor="E0E0E0">
      <td height="21" align="right" nowrap>
        <span id="yFlag" style="visibility:hidden ">N</span>全选
      <input name="yAll" type="checkbox" id="yAll" onClick="ySel()"></td>
      <td><input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td align="center">&nbsp;</td>
      <td colspan="2" align="right"><select name="yAct" id="yAct" >
          <option value="del_sel">删除.所选</option>
          <option value="d_Clear">清理数据</option>
          <%If ID<>"" Then%>
          <option value="v_Cancel">取消贺奖</option>
          <option value="r_Vote">计算.重计选票</option>
          <option value="r_Rate">计算.吻合度</option>
          <!--<option value="Report">报表.贺奖报表</option>-->
          <%End If%>
      </select></td>
      <td align="left"><input type="submit" name="Submit" value="执行">      </td>
    </tr>
    <%  
  Else
  %>
    <tr align="center" bgcolor="#f4f4f4">
      <td colspan="11">无信息</td>
    </tr>
    <%
  End If
	  
	  rs.Close()
	  'Set rs = Nothing
	  
	  %>
    <tr bgcolor="#999999">
      <td colspan="11" align="right"></td>
    </tr>
  </form>
</table>


<%

Function SetLucks(xID,xTop,xNum,xOrd)
Dim sql,sID : sID=""
Dim aID,i,n,r,t
 sql = " SELECT TOP "&xTop&" KeyID FROM [VoteLogs] "
 sql = sql& " WHERE KeyMod='"&xID&"' AND InfLuck='X' ORDER BY InfRate DESC "
 rs.Open Sql,conn,1,1
 Do While NOT rs.EOF
  sID=sID&rs("KeyID")&";"
 rs.MoveNext()
 Loop
 rs.Close()
 If Len(sID)>24 Then
  sID=Left(sID,Len(sID)-1)
  If Len(sID)=24 Then '/////////////////// 一条记录
	Call rs_DoSql(conn,"UPDATE VoteLogs SET InfLuck="&xOrd&" WHERE KeyID='"&sID&"'") 'Response.Write sID
	SetLucks=1
	Exit Function
  End If
  aID=Split(sID,";") : n=uBound(aID)'///// 多条记录
  Randomize ' Int(N*Rnd())=0~(N-1)
  For i=0 To n
	r=Int(n*Rnd()) : t=aID(r) 
	aID(r)=aID(i) : aID(i)=t
  Next
  r=Int(xNum)-1 : If r>=n Then r=n
  For i=0 To r
	Call rs_DoSql(conn,"UPDATE VoteLogs SET InfLuck="&xOrd&" WHERE KeyID='"&aID(i)&"'") 'Response.Write aID(i)&", "
  Next
  SetLucks=r+1
 Else '/////////////////////////////////// 无记录
  SetLucks=0
 End If
End Function

Function ReVote(xID)
Dim sql,iID
 sql = " SELECT KeyID FROM [VoteItem] WHERE KeyMod='"&xID&"' "
 rs.Open Sql,conn,1,1
 Do While NOT rs.EOF
  iID = rs("KeyID") : iCount = rs_Count(conn,"VoteLogs WHERE KeyItems LIKE '%"&iID&"%'")
  Call rs_DoSql(conn,"UPDATE VoteItem SET SetVote="&iCount&" WHERE KeyID='"&iID&"'")
 rs.MoveNext()
 Loop
 rs.Close()
End Function

Function nTopID(xID,xTop)
Dim sql,iID,sID : sID=""
 sql = " SELECT TOP "&xTop&" KeyID FROM [VoteItem] WHERE KeyMod='"&xID&"' ORDER BY SetVote DESC "
 rs.Open Sql,conn,1,1
 Do While NOT rs.EOF
  sID=sID&rs("KeyID")&";"
 rs.MoveNext()
 Loop
 rs.Close()
nTopID=sID
End Function

Function ReRate(xID,xItems,xTop)
Dim sql,iID,aID,aN,i
 aID=Split(xItems,";") :aN = uBound(aID)
 sql = " SELECT KeyID,KeyItems FROM [VoteLogs] WHERE KeyMod='"&xID&"' "
 rs.Open Sql,conn,1,1
 Do While NOT rs.EOF
  iID = rs("KeyID") : iStr = rs("KeyItems")
  iCount = 0
  For i=0 To aN
    If inStr(iStr,aID(i))>0 And Len(aID(i))>8 Then
	  iCount=iCount+1
	End If
  Next
  Call rs_DoSql(conn,"UPDATE VoteLogs SET InfRate="&iCount/xTop&" WHERE KeyID='"&iID&"'")
 rs.MoveNext()
 Loop
 rs.Close()

End Function

 Set rs = Nothing
%>

<script type="text/javascript">
function ySel()
{
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
   yFlag.innerText = "Y";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=true;}
   }else{
   yFlag.innerText = "N";
   for(var i=0;i<document.flist.yID.length;i++)
   {document.flist.yID.item(i).checked=false;}
   }
}  

</script>
</body>
</html>
