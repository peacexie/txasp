<table width="480" border="0" align="center" cellpadding="3" cellspacing="8">
  <tr>
    <td width="50%" align="left" valign="top"><a href="/bbs/showList.Asp?Action=hot" target="_blank">热贴排行</a>
      <table width="96%" border="0" align="center" cellpadding="3" cellspacing="1">
        <%
sql = "SELECT TOP 5 * FROM [YX_Topic] WHERE 1=1 ORDER BY Hits DESC " 
rs.Open sql,cbbs,1,1 
i = 0
Do While NOT rs.EOF
 i = i + 1
 aID = rs("TopicID")
 aSubj = Show_SLen(rs("Caption"),15)
 aName = rs("Name")
 If i<=5 Then
%>
        <tr>
          <td width="5%" align="center"><img src="../pimg/jjj.jpg" width="3" height="3" /></td>
          <td align="left"><a href="/bbs/Show.Asp?ID=<%=aID%>" target="_blank"><%=aSubj%></a></td>
        </tr>
        <%
  End If
rs.MoveNext()
Loop
rs.Close()
%>
      </table></td>
    <td width="50%" align="left" valign="top"><a href="/bbs/showList.Asp?Action=new">最新帖子</a>
      <table width="96%" border="0" align="center" cellpadding="3" cellspacing="1">
        <%
sql = "SELECT TOP 5 * FROM [YX_Topic] WHERE 1=1 ORDER BY AddTime DESC " 
rs.Open sql,cbbs,1,1 
i = 0
Do While NOT rs.EOF
 i = i + 1
 aID = rs("TopicID")
 aSubj = Show_SLen(rs("Caption"),15)
 aName = rs("Name")
 If i<=5 Then
%>
        <tr>
          <td width="5%" align="center"><img src="../pimg/jjj.jpg" width="3" height="3" /></td>
          <td align="left"><a href="/bbs/Show.Asp?ID=<%=aID%>" target="_blank"><%=aSubj%></a></td>
        </tr>
        <%
  End If
rs.MoveNext()
Loop
rs.Close()
%>
      </table></td>
  </tr>
  <tr>
    <td align="left" valign="top"><a href="/bbs/Members.Asp">发贴英雄</a>
      <table width="96%" border="0" align="center" cellpadding="3" cellspacing="1">
        <%
sql = "SELECT TOP 5 * FROM [YX_User] WHERE 1=1 ORDER BY EssayNum DESC " 
rs.Open sql,cbbs,1,1 
i = 0
Do While NOT rs.EOF
 i = i + 1
 uID = rs("ID")
 uName = rs("Name")
 uENum = rs("EssayNum")
 If i<=5 Then
%>
        <tr>
          <td width="5%" align="center" nowrap="nowrap"><img src="../pimg/jjj.jpg" width="3" height="3" /></td>
          <td align="left" nowrap="nowrap"><a href="/bbs/user_id2name.Asp?id=<%=uID%>" target="_blank"><%=uName%></a></td>
          <td width="15%" align="left" nowrap="nowrap"><%=uENum%>篇</td>
        </tr>
        <%
  End If
rs.MoveNext()
Loop
rs.Close()
%>
      </table></td>
    <td align="left" valign="top"><a href="/bbs/List.Asp?BoardID=15" target="_blank">旅游话题</a>
      <table width="96%" border="0" align="center" cellpadding="3" cellspacing="1">
        <%
sql = "SELECT TOP 5 * FROM [YX_Topic] WHERE BoardID=3 ORDER BY AddTime DESC " 
rs.Open sql,cbbs,1,1 
i = 0
Do While NOT rs.EOF
 i = i + 1
 aID = rs("TopicID")
 aSubj = Show_SLen(rs("Caption"),15)
 aName = rs("Name")
 If i<=5 Then
%>
        <tr>
          <td width="5%" align="center"><img src="../pimg/jjj.jpg" width="3" height="3" /></td>
          <td align="left"><a href="/bbs/Show.Asp?ID=<%=aID%>" target="_blank"><%=aSubj%></a></td>
        </tr>
        <%
  End If
rs.MoveNext()
Loop
rs.Close()
%>
      </table></td>
  </tr>
</table>
