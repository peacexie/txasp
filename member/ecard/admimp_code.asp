<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%

act = Request("act") ':Response.Write act

mstr1 = Request("mstr1") ':Response.Write mstr1
mstr2 = Request("mstr2") ':Response.Write mstr2

fstr = Request("fstr") :if fstr="" Then fstr="test_no item_no item_name item_code"
fstr = Trim(Replace(Replace(fstr,"["," "),"]"," ")) ':Response.Write fstr
fstr = Replace(Replace(Replace(fstr,","," "),";"," "),"	"," ")
fstr = Replace(Replace(Replace(fstr,vbcrlf," "),vblf," "),vbcr," ")
fstr = Replace(Replace(Replace(fstr,"    "," "),"  "," "),"  "," ")
'Response.Write fstr

If act<>"" Then 
  mstr1 = Replace(Replace(mstr1,"[%=","<%="),"%]","%"&">")
  mstr2 = Replace(Replace(mstr2,"[%=","<%="),"%]","%"&">")
  mstr1 = Replace(mstr1,"`","""") :If inStr(lCase(mstr1),"</td>")<=0 Then mstr1=mstr1&"<br>"
  mstr2 = Replace(mstr2,"`","""") :If inStr(lCase(mstr2),"</td>")<=0 Then mstr2=mstr2&"<br>"
  a = Split(fstr," ") : s1="" : s2=""
  for i=0 to uBound(a) 
	  imod = mstr1
	  imod = Replace(imod,"fID",a(i))
	  imod = Replace(imod,"(n)","("&i&")")
	  s1 = s1&vbcrlf&imod&""
	If mstr2<>"<br>" Then
	  imod = mstr2
	  imod = Replace(imod,"fID",a(i))
	  imod = Replace(imod,"(n)","("&i&")")
	  s2 = s2&vbcrlf&imod&""
	End If
  next
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>代码生成器</title>
<style type="text/css">
<!--
body, td, th {
	font-size: 14px;
	font-family:"Courier New", Courier, monospace;
}
select, option {
	font-family:"Courier New", Courier, monospace;
}
td {
	background-color:#FFF;
}
-->
</style>
</head>
<body>
<form id="fm01" name="fm01" method="post" action="?">
  <table width="680" border="1" align="center" cellpadding="2" cellspacing="1">
    <tr>
      <td align="center">Fields<br />
        <select name="act1" id="act1" onchange="set1(this.value)">
          <option value="0">(set)</option>
          <option value="1">Dome1</option>
          <option value="2">Dome2</option>
          <option value="3">Dome3</option>
          <option value="4">Dome4</option>
          <option value="4">Dome4</option>
        </select></td>
      <td><textarea name="fstr" cols="60" rows="4" id="fstr" style="width:580px"><%=fstr%></textarea></td>
    </tr>
    <tr id="rowID">
      <td align="center">Custom1</td>
      <td nowrap="nowrap"><textarea name="mstr1" rows="2" id="mstr1" style="width:580px"><%=Replace(mstr,"""","&quot;")%></textarea></td>
    </tr>
    <tr id="rowID2">
      <td align="center">Custom2</td>
      <td nowrap="nowrap"><textarea name="mstr2" rows="2" id="mstr2" style="width:580px"><%=Replace(mstr,"""","&quot;")%></textarea></td>
    </tr>
    <tr>
      <td align="center">Model</td>
      <td nowrap="nowrap"><select name="act" id="act" style="width:480px" onchange="set(this)">
          <option value="mImp">(mImp) --- fID = RequestSafe(rs(`fID`),`C`,48)</option>
          <option value="mIns">(mIns) --- sql = sql&amp; `, fID` values ...</option>
          <option value="mAdd">(mAdd) --- rs.Fields("fID").Value = fID</option>
          <option value="mEdit">(mEdit) ----- sql = sql&amp; `,fID = '` &amp; RequestS(`fID`,`C`,48) &amp;"'"</option>
          <option value="mShow">(mShow) ----- fID = rs("fID")</option>
          <option value="mForm">(mForm) ----- fID: &lt;input type='text'&gt; </option>
          
          <option value="setCust">(setCust) ----- ...... ...... (手工设置) ...... ...... </option>
          
          <option value="cmImp">(cmImp) --- .Parameters.Append .CreateParameter(`@fID`,202,1,48,fID)</option>
          <option value="cmItem">(cmCell) ----- fID = rs(0)(`fID`)</option>
          
        </select>
        <input type="submit" name="doAct" id="doAct" value="Build" />
        <a href="?fstr=<%=fstr%>">reLoad</a></td>
    </tr>
  </table>
  
</form>
<table border='0' align='center' cellpadding='0' cellspacing='1' bgcolor='#CCCCCC' width="680">
  <% 

Response.Write vbcrlf&s1&vbcrlf
If s2<>"" Then Response.Write vbcrlf&"<BR>"&vbcrlf
Response.Write vbcrlf&s2&vbcrlf
%>
</table>
<script type="text/javascript">
fid = document.fm01.fstr;
mid = document.fm01.mstr;
function set1(v){
  var s1 = new Array();
  sd = "<%=fstr%>";
  s1[0] = "MemID MemPW MemQu MemAn MemType MemTyp2 MemGrade MemName MemNam2 MemSex MemCard MemBirth MemFrom MemMobile MemTel MemEmail MemExp MemFlag LogAddIP LogAUser LogATime LogEditIP LogEUser LogETime"
  s1[1] = "OrdID OrdDate OrdWeek OrdDptID OrdDptName OrdDocID OrdDoctor OrdDocTitle OrdType OrdTime OrdNFee OrdNMax OrdNNow"
  s1[2] = "test_no patient_id visit_id working_id execute_date name name_phonetic sex age test_cause relevant_clinic_diag specimen notes_for_spcm spcm_received_date_time spcm_sample_date_time requested_date_time ordering_dept ordering_provider performed_by results_rpt_date_time transcriptionist verified_by billing_indicator print_indicator ic_id send_man test_man"
  s1[3] = "test_no item_no item_name item_code billing_indicator"
  s1[4] = "test_no item_no print_order report_item_name report_item_code result units abnormal_indicator instrument_id result_date_time print_context"
  fid.value = s1[v-1];
  if(v==0){ fid.value = sd; }
  else{     fid.value = s1[v-1]; }
}

function set(e){
  var v = e.value;
  for(i=1;i<3;i++){
    var itm = document.getElementById("mstr"+i);
	try{
	  var idat = document.getElementById(v+i).innerHTML;
	  for(j=0;j<8;j++) { idat = idat.replace("&amp;","&"); }
	  idat = idat.replace("value=[%=fID%]","value='[%=fID%]'");
	  idat = idat.replace("[%","<"+"%");
	  idat = idat.replace("%]","%"+">");
	  itm.value = idat;
    }catch(err){
	  itm.value = "";	
	}
  }
}
document.fm01.act.value = 'setCust';
//set(document.fm01.act)
//document.fm01.act2.value = '<%=Request("act2")%>';
//set1(1);
</script>
<div style="display:none">

<div id="mImp1">fID = RequestSafe(rs(`fID`),`C`,48)</div>
<div id="mImp2">fID = RequestSafe(ra(n),`C`,48)</div>
<div id="mIns1">sql = sql& `, fID`</div>
<div id="mIns2">sql = sql& `, '` & fID &`'`</div>
<div id="mAdd1">rs.Fields(`fID`).Value = fID</div>
<div id="mEdit1">sql = sql& `,fID = '` & RequestS(`fID`,`C`,48) &`'`</div>
<div id="mShow1">fID = rs(`fID`)</div>
<div id="mShow2"><td>[%=fID%]</td></div>
<div id="mForm1"><tr><td>fID</td><td><input name='fID' type='text' id='fID' value='[%=fID%]' size='24' maxlength='48' /></td></tr></div>

<div id="cmImp1">.Parameters.Append .CreateParameter(`@fID`,202,1,48,fID)</div>
<div id="cmItem1">rs(0)(`fID`)</div>
<div id="cmItem2">cm.Parameters.Item(n).Value = fID</div>

</div>
</body>
</html>
