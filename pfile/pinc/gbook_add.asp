<script type="text/javascript" src="../sadm/edfck/fckeditor.js"></script>
<%
'Config_Path = "/u/demo/"

 aCode = Split(strDepCode,"|")
 aName = Split(strDepName,"|")
 aFlag = Split(strDepFlag,"|")
 TreeOpt = "" : TreeArr = ""
 For i=0 To uBound(aCode)-1
    TreeOpt=TreeOpt&vbcrlf&"<option value='"&aName(i)&"'>"&aName(i)&"</option>"
    TreeArr=TreeArr&"new Array("&gdArr(aCode(i))&"),"&vbcrlf
 Next

Function gdArr(xID)
  Dim rs,s : s="'不限'" 
  Dim sql,KeyID,InfSubj
  sql = " SELECT * FROM [InfoPics] WHERE KeyMod='PicR124' AND InfTyp2='"&xID&"' " ' AND SetShow='Y'
  Set rs=Server.CreateObject("Adodb.Recordset")
  rs.Open Sql,conn,1,1 '("&rs.RecordCount&"人)
  If NOT rs.EOF Then
  Do While NOT rs.EOF
    KeyID = rs("KeyID")
    InfSubj = rs("InfSubj")
     s=s&",'"&InfSubj&"'"
  rs.MoveNext
  Loop
  End If
  Set rs=Nothing
  gdArr = s
End Function

  Dim sys27_Rnd(10)
  sys27_RVal = Rnd_Base("5678",9)&Rnd_Base("",64)
  For i = 1 To 9
    sys27_Rnd(i)=Mid(sys27_RVal,i*6,Mid(sys27_RVal,i,1))
  Next

%>

<a name="Write" id="Write"></a>
<table width="686"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tdbg">
  <form name="fm01" id="fm01" action="gbadd.asp" method="post">
    <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=sys27_RVal%>">
    <tr>
      <td width="10%" align="center"><%=vGbo_Subj%></td>
      <td><input name="InfSubj<%=sys27_Rnd(1)%>" type="text" id="InfSubj<%=sys27_Rnd(1)%>" size="36" maxlength="60">
        &nbsp;<%=vInf_Type%>
        <select name="InfType" id="InfType" style="width:210px;">
          <%=Get_rsOpt(conn,"SELECT TypID,TypName FROM WebTyps WHERE TypMod='"&MD&"'",TP)%>
        </select></td>
    </tr>
    <tr>
      <td align="center"><%=vGbo_Name%></td>
      <td><input name="LnkName<%=sys27_Rnd(4)%>" type="text" id="LnkName<%=sys27_Rnd(4)%>" value="<%=Session("MemID")%>" size="36" maxlength="24">
        &nbsp;<%=vGbo_Email%>
        <input name="LnkEmail<%=sys27_Rnd(5)%>" type="text" id="LnkEmail<%=sys27_Rnd(5)%>" size="36" maxlength="120" /></td>
    </tr>
    <%If SwhDepSubs="Y" Then%>
    <tr>
      <td align="center">部门</td>
      <td><select name="TreeTypA" size="1" id="TreeTypA" style="width:210px;" onchange="TreeChange()">
          <option value="不限">选择部门</option>
          <%=TreeOpt%>
          <option value="不限">不限</option>
        </select>
        &nbsp;人员
        <select name="TreeTypB" size="1" id="TreeTypB" style="width:210px;">
          <option value="不限">选择医生</option>
        </select></td>
    </tr>
    <%End If%>
    <%If SwhGbkEditor="Y" Then%>
    <tr>
      <td align="center"><%=vGbo_Cont%></td>
      <td>
	  
<textarea id="InfCont<%=sys27_Rnd(2)%>" name="InfCont<%=sys27_Rnd(2)%>" style="width:600px;height:300px;visibility:hidden;display:none"></textarea>
<script type="text/javascript" charset="utf-8" src="../smod/file/edt_api.asp?edtID=EditID01<%=sys27_Rnd(2)%>&edtCont=InfCont<%=sys27_Rnd(2)%>&edtTool=Basic"></script>

	  </td>
    </tr>
    <%Else%>
    <tr>
      <td align="center"><%=vGbo_Cont%></td>
      <td><textarea name="InfCont<%=sys27_Rnd(2)%>" cols="72" rows="12" id="InfCont<%=sys27_Rnd(2)%>"></textarea></td>
    </tr>
    <%End If%>
    <tr>
      <td align="right" nowrap><%=vPMsg_ChkCode%></td>
      <td nowrap><input name="ChkCode" type="text" id="ChkCode" size="6" maxlength="12">
        <img src="../sadm/pcode/img_frnd.asp" alt="<%=vPMsg_RelCode%>" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;" onClick="PicReLoad('../')"/>
        <input type="button" name="Button" value="<%=vGbo_Send%>" onclick="chkData()" />
        <input type="reset" name="Submit2" value="<%=vGbo_Reset%>" />
        <input name="send" type="hidden" id="send" value="ins" />
        <input name="ID" type="hidden" id="ID" value="<%=ID%>" />
        <input name="TP" type="hidden" id="TP" value="<%=TP%>" /></td>
    </tr>
  </form>
</table>
<script type="text/javascript">

var TreeForm = document.getElementById('fm01');
var TreeOptA = TreeForm.TreeTypA;
var TreeOptB = TreeForm.TreeTypB;
var TreeArr = new Array(
<%=TreeArr%>
new Array("不限")
);

function TreeChange(){
  index = TreeOptA.options.selectedIndex-1;
  if(index<0){
      for(var i=(TreeOptB.options.length-1);i>0;i--){TreeOptB.options.remove(i);}
	  return;
  };
  TreeOptB.length = TreeArr[index].length;
  for(var i = 0;i<TreeArr[index].length;i++)
  {
    var iVal = TreeArr[index][i];
    TreeOptB.options[i].text = iVal;
    TreeOptB.options[i].value = iVal;
  }
}

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

    <%If SwhGbkEditor="Y" Then%>
  document.fm01.InfCont<%=sys27_Rnd(2)%>.value = apiGetValEditID01<%=sys27_Rnd(2)%>(); 
    <%End If%>
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length==0) 
   {   
     alert(" <%=vGbo_jsCont%>"); 
	 document.fm01.InfCont.focus();
     eflag = 1; break;
   }
 if (document.fm01.InfCont<%=sys27_Rnd(2)%>.value.length>=12000) 
   {   
     alert(" <%=vGbo_jsCMax%> 12 K!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ document.fm01.submit(); }
}</script>
