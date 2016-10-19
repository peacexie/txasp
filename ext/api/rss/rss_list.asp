<?xml version="1.0" encoding="utf-8"?>
<rss version="2.0">
  <channel>
    <!--#include file="rss_config.asp"-->
    <%

TypID = Request("TypID") ''InfN124
nID = fChkRID(TypID)

If nID<0 Then
  sSubjs = "(错误提示:[参数错误]!!!)"
Else
  sSubjs = aTitle(nID)
End If

%>
    <title><%=sSubjs%>-<%=Config_Name%></title>
    <link>
    <%=Config_URL%>
    </link>
    <description>
    欢迎订阅 PeaceAsp综合网站系统 RSS
    </description>
    <language>zh-cn</language>
    <generator>Dreamweaver,Notepad</generator>
    <ttl>120</ttl>
    <%
  If nID<0 Then
	Response.Write "  </channel>"&vbcrlf&"</rss>" 
    Response.End() 
  End If

 sql = " SELECT KeyID,InfSubj,InfPara"
 sql =sql& ",LogATime,LogAUser FROM ["&aTabID(nID)&"] "
 sql =sql& " WHERE KeyMod='"&aModID(nID)&"' "
 If aTypID(nID)<>"" Then
 sql =sql& " AND InfType LIKE '%"&aTypID(nID)&"%' " 
 End If
 sql =sql& " ORDER BY LogATime DESC " 
 Set rs=Server.Createobject("Adodb.Recordset")
 rs.Open Sql,conn,1,1
 Do While NOT rs.Eof 
 
  KeyID = rs("KeyID")
  InfSubj = fFilRss(rs("InfSubj"))
  LogATime = rs("LogATime") 'FormatDateTime(rs("LogATime"),2) 
  LogATGMT = fToGMT(LogATime)
  LogAUser = rs("LogAUser")
  LogAName = rs_Val("","SELECT UsrName FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&LogAUser&"'")
  InfRem = "Publish:"&LogATime&"" 
  InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
  If aPara(1)<>"" Then
    InfRem = InfRem&"; Keywords:"&aPara(1)&"" 
  End If
  If aPara(2)<>"" Then
    InfRem = InfRem&"; From:"&aPara(2)&"" 
  End If
  If aPara(3)<>"" Then
    InfRem = InfRem&"; Speci.:"&aPara(3)&"" 
  End If
%>
    <item>
      <title>
      <![CDATA[<%=InfSubj%>]]>
      </title>
      <link>
      <![CDATA[rss_view.asp?TypID=<%=TypID%>&KeyID=<%=KeyID%>]]>
      </link>
      <pubDate><%=LogATGMT%></pubDate>
      <source>
        <![CDATA[<%=InfFrom%>]]>
      </source>
      <author>
        <![CDATA[<%=LogAUser%>]]>
      </author>
      <description>
        <![CDATA[<%=InfRem%>]]>
      </description>
    </item>
    <%
 rs.MoveNext()
 Loop
 rs.Close()
 Set rs = Nothing
%>
  </channel>
</rss>
