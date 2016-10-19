
<FIELDSET style="padding:3px;">
<LEGEND style="padding:5px;"><%=vIRem_SchMsg%></LEGEND>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr align="center" bgcolor="#FFFFFF">
    <td align="left" bgcolor="#FFFFFF"><%=vIRem_OrgPage%>: <a href="<%=ObjUrl%>"><%=ObjSubj%></a></td>
    <form name="fm02" method="post" action="?">
      <td align="right" nowrap><font color="#FF0000">
        <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
        <input name="ObjID" type="hidden" id="ObjID" value="<%=ObjID%>" />
        <input name="ObjSubj" type="hidden" id="ObjSubj" value="<%=ObjSubj%>" />
        <input name="ObjUrl" type="hidden" id="ObjUrl" value="<%=ObjUrl%>" />
        <%=rMsg%></font>&nbsp;<%=vIRem_SchKey%>:
        <input name="KW" type="text" id="KW" value="<%=KW%>" size="12">
        <input type="submit" name="Submit" value="<%=vIRem_Search%>">
        </td>
    </form>
  </tr>
</table>
</FIELDSET>
<div style="line-height:8px;">&nbsp;</div>
<%
sqlK = ""
KW = RequestS("KW","C",24)
If KW<>"" Then
  sqlK=sqlK&" AND (InfSubj LIKE '%"&KW&"%' OR InfCont LIKE '%"&KW&"%') "
End If
isCheck = " AND SetShow='Y' "

    sql = " SELECT GboSend.* FROM [GboSend] "
	sql =sql& " WHERE KeyMod='"&MD&"' "&isCheck&" AND KeyCode='"&ObjID&"' "&sqlK
	sql =sql& " ORDER BY KeyID DESC" 'SetTop,
   rs.Open Sql,conn,1,1
   rs.PageSize = 24 
If Int(Page)>rs.PageCount Or Int(Page)<1 Then
  Page = 1
End If
%>
<%
  
  If rs.eof then
  %>
<FIELDSET style="padding:3px;">
<LEGEND style="padding:5px;"><%=vPMsg_NoRec%></LEGEND>
</FIELDSET>
<%
  Else
  rs.AbsolutePage = Page
  For i = 1 To rs.PageSize
		If i mod 2 = 1 Then
		  col = "#ffffff"
		Else
		  col = "#F8F8F8"
		End If
KeyID = rs("KeyID")
InfType = rs("InfType")
InfSubj = Show_Text(rs("InfSubj"))
LogATime = rs("LogATime")
LogAUser = rs("LogAUser")
LogAddIP = rs("LogAddIP") 
InfSubj = Show_Text(InfSubj) 
xxxCont = rs("InfCont")
	  %>
<FIELDSET style="background:<%=col%>">
<LEGEND><B><%=InfSubj%></B> <a href="<%=ObjUrl%>"><%=vIRem_OrgPage%></a></LEGEND>
<table width="100%"  border="0" align="center" cellpadding="5" cellspacing="1">
  <tr>
    <td align="left" valign="top" class="SysCont" style="padding:8px;"><%Call Show_sfGbook(KeyID,".out.htm")%></td>
    <td width="15%" align="left" valign="top" nowrap style="line-height:120%;"><%=LogATime%><br>
      <%=vIRem_Name%>: <%=LogAUser%><br>
      IP:<%=LogAddIP%> </td>
  </tr>
</table>
</FIELDSET>
<div style="line-height:8px;">&nbsp;</div>
<%
  rs.Movenext
  If rs.Eof Then Exit For
  Next
%>
<FIELDSET style="padding:3px;">
<%= RS_Page(rs,Page,"?send=pag&ModID="&MD&"&ObjID="&ObjID&"&ObjSubj="&ObjSubj&"&KW="&KW&"",1)%>
</FIELDSET>
<%    
  End If
	  rs.Close()
	  
	  %>
<div style="line-height:8px;">&nbsp;</div>
<FIELDSET style="padding:3px;">
<LEGEND style="padding:5px;"><B><%=vIRem_Add%></B> <a href="<%=ObjUrl%>"><%=vIRem_OrgPage%></a></LEGEND>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><%=vIRem_Subj%></td>
      <td bgcolor="#FFFFFF"><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" value="<%=vPMsg_Remark%>: <%=ObjSubj%>" size="60" maxlength="120">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><%=vIRem_Name%></td>
      <td bgcolor="#FFFFFF"><input name="LogAUser<%=sys27_Rnd(4)%>" type="text" id="LogAUser<%=sys27_Rnd(4)%>" value="<%=Session("MemID")%>" size="18" maxlength="24" />
        &nbsp;&nbsp;<%=vIRem_Attitude%>:
        <select name="InfType" id="InfType">
          <option value="<%=vIRem_AttOK%>"><%=vIRem_AttOK%></option>
          <option value="<%=vIRem_AttNG%>"><%=vIRem_AttNG%></option>
          <option value="<%=vIRem_AttMid%>"><%=vIRem_AttMid%></option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" bgcolor="#FFFFFF"><%=vIRem_Cont%></td>
      <td bgcolor="#FFFFFF"><textarea name="InfCont<%=sys27_Rnd(2)%>" cols="50" rows="8" id="InfCont<%=sys27_Rnd(2)%>"></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><%=vPMsg_ChkCode%></td>
      <td bgcolor="#FFFFFF"><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12" xonfocus="javascript:PicReLoad('../');"/>
        <img src="../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onclick="PicReLoad('../')"/></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="center" nowrap bgcolor="#FFFFFF"><input name="send" type="hidden" id="send" value="send" />
      </td>
      <td bgcolor="#FFFFFF"><input name=view type=button id="Button1" value="<%=vIRem_BtnOK%>" onClick="chkData()">
        <input name=view type=reset id="Button1" value="<%=vIRem_BtnNO%>">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="Page" type="hidden" id="Page" value="<%=Page%>">
        <input name="ModID" type="hidden" id="ModID" value="<%=MD%>" />
        <input name="ObjID" type="hidden" id="ObjID" value="<%=ObjID%>">
        <input name="ObjSubj" type="hidden" id="ObjSubj" value="<%=ObjSubj%>" />
        <input name="ObjUrl" type="hidden" id="ObjUrl" value="<%=ObjUrl%>" />
        
        <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
        
        </td>
    </tr>
  </form>
</table>
</FIELDSET>
<script type="text/javascript">
 function chkData()
 {
       var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (document.fm01.InfSubj<%=sys27_Rnd(1)%>.value.length==0) 
   {   
     alert(" <%=vGbo_jsSubj%>"); 
     document.fm01.InfSubj<%=sys27_Rnd(1)%>.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length==0) 
   {   
     alert(" <%=vGbo_jsCont%>"); 
	 document.fm01.InfCont<%=sys27_Rnd(2)%>.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length>=1200) 
   {   
     alert(" <%=vGbo_jsCMax%>1.2 K!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
