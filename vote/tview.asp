<!--#include file="_config.asp"-->
<%

ID = RequestS("ID","C",48)
Page = RequestS("Page","N",1)
PSize = 24

  rs.Open "SELECT * FROM [VoteInfo] WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID") : js="VotCount=0;"
KeyMod = rs("KeyMod")
InfSubj = Show_Text(rs("InfSubj"))
InfTim1 = rs("InfTime1") 
InfTim2 = rs("InfTime2")
InfCard = rs("InfCard")  
InfVNum = rs("InfVNum") 
InfNum1 = rs("InfNum1") 
InfNum2 = rs("InfNum2")  
ImgName = rs("ImgName")
  else
KeyID = ""  
  end if 
  rs.Close()
If KeyID = "" Then Response.Redirect("rlist.asp")

If DateDiff("d",InfTim2,Date())>0 Then
   fVNotes = "调查已结束! ( <a href='vprize.asp?ID="&ID&"' target=_blank>查看贺奖名单</a>) "
ElseIf DateDiff("d",Date(),InfTim1)>0 Then
   fVNotes = "调查还未开始!"
Else
   fVNotes = "调查进行中... ( <a href='vprize.asp?ID="&ID&"' target=_blank>查看贺奖名单</a>) " 
End If

Dim aLogs
sql = "SELECT KeyItems FROM [VoteLogs] WHERE KeyMod='"&ID&"' ORDER BY KeyID "
rs.Open sql,conn,2,1 
If Not rs.EOF Then
aLogs = rs.Getrows()
aRecs = UBound(aLogs,2)+1
shRec = aRecs  '显示
Else
'ReDim aLogs(1,1) ' = Split("","|")
aRecs = -1     ' 计算时，除数避开0，用-1
shRec = 0      ' 用于显示
End If
rs.Close() ':Response.Write sql
If Int(Page)>Int(shRec) Or Int(Page)<1 Then
  Page = 1
End If
'GetRows方法通常比一次读一笔记录的回圈要来得快些，但使用这方法时，必须确定Recordset未包含太多记录；
'否则，会很容易以一个非常大的变数阵列来填满所有记忆体。基於相同的原因，得小心不要包括任何
'BLOB（Binary Large Object）或CLOB（Character Large Object）栏位；若如此做的化，
'应用程式一定会爆掉，特别是对於较大的Recordset而言。最后，记住此方法传回的变数阵列是以0为基底的；
'传回记录的笔数是UBound(values,2)+1，传回栏位数是UBound(value, 1)+1。
Function CntItems(xPos)
 Dim n,i,c
  n = 0
  If aRecs>0 Then
   For i=0 To UBound(aLogs,2) 'uBound(aLogs)
    c = Mid(aLogs(0,i),xPos,1)
	If c="1" Then n = n + Int(c)
   Next
  Else
    c=0
  End If
 CntItems = n
End Function



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线调查投票系统-<%=Config_Name%></title>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
</HEAD>
<BODY>
<!--#include file="_xtop.asp"-->
                  <!--////////////////////////////////////////Start主体-->
                  <table border="1" align="center" cellpadding="0" cellspacing="5">
                    <form action="?" method="post" name="fm01">
                      <tr>
                        <td align="left"><FIELDSET style="margin:5px; padding:5px;">
                          <LEGEND align="center"> &nbsp;<b><%=InfSubj%></b> (<%=InfTim1%> ~ <%=InfTim2%>) &nbsp; </LEGEND>
                          说明：<span class="fnt00F"><%=fVNotes%></span>； 共有(<span class="fnt00F"><%=shRec%>人次</span>)调查记录<!--；因为有可能为多选，每一项<span class="fnt00F">百分比结果相加不一定为100%</span>-->。<span style="float:right; padding-right:5%;">
                          <input name="" type="button" value="返回调查页" onclick="javascript:location.href='tvote.asp?ID=<%=ID%>';"/>
                          </span>
                          </FIELDSET></td>
                      </tr>
                      <tr>
                        <td width="650" align="left" valign="bottom" style="padding:5px 2px 5px 5px;"><%
sql = " SELECT * FROM [VoteItem] WHERE KeyMod='"&ID&"' "
sql = sql & " ORDER BY SetTop,KeyID " 
rs.Open Sql,conn,1,1
i=0 : ii=0 : kk=1
Do While NOT rs.EOF
i=i+1
iKeyID = rs("KeyID")
iSubj = rs("InfSubj")
iRem = rs("InfRem")&""
aRem = Split(iRem,"|")
iTop = rs("SetTop")
iVote = rs("SetVote")
iImage = rs("ImgName")&""
jsItem = jsItem&iKeyID&"|"
jsITyp = jsITyp&iImage&"|"
If iImage<>"Select" Then
 If iImage="FBlank" Then
  iImgExt = " (120个字以内 <font color=red>必填项！！！</font>)"
 Else
  iImgExt = " (120个字以内 <font color=gray>可空白</font>)"
 End If
End If
%>
                          <div style="padding-bottom:12px;">
                            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td align="left"><b><span id="Msg<%=iKeyID%>"><%=i%>. <%=iSubj%></span></b> <%=iImgExt%> </td>
                              </tr>
                              <tr>
                                <td><%
								If iImage="Select" Then
								  For j=0 To uBound(aRem)
								jItem = Trim(Show_Text(aRem(j)))
								jChar = Chr(65+j)
								iii = iii+1
								ivn = CntItems(iii)
								iw1 = 100*(ivn/aRecs)
								iw0 = FormatNumber(iw1,2)
								iw1 = Int(iw1)
								iw2 = 100-iw1
								%>
                                  <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:5px;">
                                    <tr>
                                      <td colspan="2"><%=jChar&". "&jItem%> <span class="fnt00F"><%=ivn%>票 (<%=iw0%>%) </span> </td>
                                    </tr>
                                    <tr>
                                      <td width="<%=iw1%>%" background="../img/vote/vote52.gif"><img src="../img/vote/vote52.gif" width="1" height="15" align="absmiddle" /></td>
                                      <td width="<%=iw2%>%" background="../img/vote/vote62.gif"><img src="../img/vote/vote62.gif" width="1" height="15" align="absmiddle" /></td>
                                    </tr>
                                  </table>
                                  <%
								    Next
								  End If    ' iImage="Select" Then
								  %>
                                </td>
                              </tr>
                            </table>
                          </div>
                          <%
						   rs.MoveNext()
						   Loop
						   rs.Close()
						   %>
                        </td>
                      </tr>

                    </form>
                  </table>
                  <!--////////////////////////////////////////End主体-->
                  <!--#include file="_xbot.asp"-->
<script type="text/javascript">var jsItem='<%=jsItem%>';var jsITyp='<%=jsITyp%>';</script>
<script src="research.js" type="text/javascript"></script>
</BODY>
</HTML>
