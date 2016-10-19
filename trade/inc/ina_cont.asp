<div class="PubClear"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="left" class="SysCont" style="padding:5px 8px;">

  <%
iCont = 0 : sList = ""
If Flag="Cont" Then
  sql = "SELECT KeyID,InfCont FROM "&ModTab&" WHERE KeyMod='"&MD&"' AND InfTyp2='"&TP&"' "
  sql = sql& " ORDER BY LogATime ASC"
  rs.Open sql,conn,1,1
  Do While Not rs.EOF
	ID = rs("KeyID")
    xxxCont = rs("InfCont")
  rs.MoveNext
  Loop
  rs.Close()
  'Response.Write sql&ID
ElseIf tmpList="Next" Then
   sql="SELECT KeyID,InfSubj,InfCont FROM "&ModTab&" WHERE KeyMod='"&MD&"' AND InfTyp2='"&TP&"' "
   sql = sql& " ORDER BY SetTop,LogATime DESC"
   rs.Open sql,conn,1,1
   Do While Not rs.EOF
	 iCont = iCont+1
	 KeyID = rs("KeyID")
     KeyMod = rs("KeyMod")
     InfType = rs("InfType")
     InfSubj = rs("InfSubj")
	 If iCont=1 AND ID="" Then 
	   ID = KeyID
	   xxxCont = rs("InfCont")
       InfSBak = Show_Text(InfSubj)
	 ElseIf ID=KeyID Then
	   xxxCont = rs("InfCont")
       InfSBak = Show_Text(InfSubj)
	 End If
     t="<span class='ItemAbout'><A href='?KeyID="&KeyID&"&ModID="&KeyMod&"&TypID="&TP&"&DepID="&DP&"'>"&InfSubj&"</A></span>"
     sList = sList&vbcrlf&t
   rs.MoveNext
   Loop
   rs.Close()
%>
  <table width="90%" border="0" align="center" cellpadding="5" cellspacing="2">
    <tr>
      <td align="left"><%=sList%></td>
    </tr>
  </table>
  <%
End If
  Call Show_sfData(ID,"fcont.htm")  
%>
    </td>
  </tr>
</table>