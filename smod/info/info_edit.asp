<!--#include file="config.asp"-->
<!--#include file="../../sadm/func2/func_obj.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
<!doctype html>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" charset="utf-8" src="../../inc/home/jquery.js"></script>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtAct=mainLoad"></script>
<script src="../../inc/home/jskeys.js" type="text/javascript"></script>
<style type="text/css">
tr, td {
	background-color:#FFF;
}
</style>
</head>
<body>

<%

  reUrl1 = Config_Path&"smod/info/info_list.asp"
  reUrl2 = Config_Path&"smod/info/info_edit.asp"
  reUrl3 = Config_Path&"smod/info/set_top.asp"
  If Chk_URL3(reUrl1)="OK" Or Chk_URL3(reUrl2)="OK" Or Chk_URL3(reUrl3)="OK" Then
  Else
    Response.End()
  End If

KW = RequestS("KW",3,24)
TP = RequestS("TP",3,240)
TP2 = RequestS("TP2",3,38)
PG = RequestS("PG","N",1)
ID = RequestS("ID",3,48)

ModImgCount = Eval(ModID&"ImgCount")
ModImgCount = RequestSafe(ModImgCount,"N",1)

send = Request("send")

If send = "send" Then

KeyCode = RequestS("KeyCode",3,48) 
InfCont = RequestS("InfCont",3,960000) 
InfCont = Show_Html(InfCont)
If SwhRemSave="Y" Then
InfCont = RemoteReplaceUrl(InfCont, upRoot, ID)
End If
If Config_Cont="DB" Then
  xxxCont = InfCont
Else
  xxxCont = ""
End If
InfSubj = RequestS("InfSubj",3,255) 
InfType = RequestS("InfType",3,255)
InfTyp2 = RequestS("InfTyp2",3,48)
InfPara = PrmFlag
For i=1 To 96 
  iPara = RequestS("InfPara"&i,3,1200) 
  iPara = Replace(iPara,"^","")
  InfPara = InfPara&"^"&iPara
Next
SetSubj = RequestS("SetSubj",3,12)
SetHot = RequestS("SetHot",3,2)
SetTop = RequestS("SetTop",3,12)
SetShow = RequestS("SetShow",3,2)
LogATime = RequestS("LogATime","D",Now())

sql = " UPDATE "&ModTab&" SET " 
sql = sql& " InfType = '" & InfType &"'" 
sql = sql& ",InfTyp2 = '" & InfTyp2 &"'" 
sql = sql& ",KeyCode = '" & KeyCode &"'"
sql = sql& ",InfSubj = '" & InfSubj &"'" 
sql = sql& ",InfCont = '" & xxxCont &"'" 
sql = sql& ",InfPara = '" & InfPara &"'" 
sql = sql& ",SetSubj = '" & SetSubj &"'" 
sql = sql& ",SetHot = '" & SetHot &"'" 
sql = sql& ",SetTop = '" & SetTop &"'" 
sql = sql& ",SetShow = '" & SetShow &"'" 
sql = sql& ",LogATime = '" & LogATime &"'"
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Get_PUser(PrmFlag) &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' "
  Call rs_DoSql(conn,sql)
  upPath = upRoot&Replace(ID,"-","/")&"/" 
  KeyID = ID
  ImgName = rs_Val("","SELECT ImgName FROM "&ModTab&" WHERE KeyID='"&ID&"'")&""
  If Len(ImgName)<15 Then
    If Request("ImgSCopy")="Y" Then
      ImgName = ImgSUpd(ModTab,KeyID,InfCont,Request("ImgSmall1"),Request("ImgSmall2"))
    End If
    'ImgName = ImgSUpd(ModTab,ID,InfCont,120,60)
  End If
  Call add_sfFile()
  If Request("Act")="Close" Then
    Session("KeyID") = ""
	Response.Write js_Alert("修改成功!","Close","")
  Else
	Session("KeyID") = ""
	Response.Redirect "info_list.asp?TP="&TP&"&TP2="&TP2&"&KW="&KW&"&Page="&PG&"&PrmFlag="&PrmFlag
  End If
End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM "&ModTab&" WHERE KeyID='"&ID&"'",conn,1,1 
  if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
KeyCode = rs("KeyCode")
InfType = rs("InfType")
InfTyp2 = rs("InfTyp2")
InfSubj = Show_Form(rs("InfSubj"))
xxxCont = rs("InfCont") 
InfCont = Show_sfRead(ID,"/fcont.htm")
InfPara = rs("InfPara")&"" : aPara = Split(InfPara,"^")
SetRead = rs("SetRead")
SetSubj = rs("SetSubj")
SetHot = rs("SetHot")
SetTop = rs("SetTop")
SetShow = rs("SetShow")
ImgName = rs("ImgName")
LogATime = rs("LogATime") 
End If
rs.Close()
SET rs=Nothing 

If KeyID = "" Then 
  Response.Redirect("info_list.asp")
Else
  Session("KeyID") = KeyID
End If

tmrPara = get_TmrPara(false)
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="F0F0F0">
  <form name="fm01" id="fm01" action="?" method="post">
    <tr>
      <td align="center"><p>&nbsp;</p></td>
      <td align="center"><strong>[<%=rs_Val("","SELECT SysName FROM AdmSyst WHERE SysID='"&ModID&"'")%>]编辑</strong></td>
    </tr>
    <tr>
      <td align="center" nowrap>主题</td>
      <td><input name="InfSubj" type="text" id="InfSubj" value="<%=InfSubj%>" size="60" maxlength="120">
          <select name='SetSubj' size=1 class="Input_Text" id="SetSubj" style="width:90 ">
            <OPTION style='COLOR:#<%=SetSubj%>; BACKGROUND-COLOR:#<%=SetSubj%>' value='<%=SetSubj%>'>#<%=SetSubj%></OPTION>
            <OPTION style='COLOR:#000000; BACKGROUND-COLOR:#000000' value='000000'>#000000</OPTION>
            <OPTION style='COLOR:#FF0000; BACKGROUND-COLOR:#FF0000' value='FF0000'>#FF0000</OPTION>
            <OPTION style='COLOR:#00FF00; BACKGROUND-COLOR:#00FF00' value='00FF00'>#00FF00</OPTION>
            <OPTION style='COLOR:#0000FF; BACKGROUND-COLOR:#0000FF' value='0000FF'>#0000FF</OPTION>
            <OPTION style='COLOR:#00FFFF; BACKGROUND-COLOR:#00FFFF' value='00FFFF'>#00FFFF</OPTION>
            <OPTION style='COLOR:#FF00FF; BACKGROUND-COLOR:#FF00FF' value='FF00FF'>#FF00FF</OPTION>
            <OPTION style='COLOR:#FFFF00; BACKGROUND-COLOR:#FFFF00' value='FFFF00'>#FFFF00</OPTION>
          </select>
      </td>
    </tr>
    <tr>
      <td align="center" nowrap><%=Get_TypeLable(ModID)%></td>
      <td><select name="InfType" id="InfType">
        <%=Get_TypeOpt(ModID,InfType)%>
      </select>
        <input name=Button type=button id="Button2" value="选择模板" onClick="owTemp()" />
&nbsp; &nbsp;
  		<%=Get_Typ2Opt(ModID,InfTyp2) %>
</td>
    </tr>
    <%=tmrPara(0)%>
    <tr>
      <td align="center">内容</td>
      <td>
<textarea id="InfCont" name="InfCont" style="width:600px;height:360px;visibility:hidden;display:none"><%=Show_Form(InfCont)%></textarea>
<script type="text/javascript" charset="utf-8" src="../../smod/file/edt_api.asp?edtID=EditID01&edtCont=InfCont"></script>
      </td>
    </tr>
    <tr>
      <td align="center">属性</td>
      <td><select name="SetHot" id="SetHot">
        <option value="N" <%If SetHot="N" Then Response.Write("selected")%>>一般</option>
        <option value="Y" <%If SetHot="Y" Then Response.Write("selected")%>>推荐</option>
      </select>
          <%If Flg_PCheck() Then%>
        <select name="SetShow" id="SetShow">
          <option value="Y" <%If SetShow="Y" Then Response.Write("selected")%>>审核</option>
          <option value="N" <%If SetShow="N" Then Response.Write("selected")%>>未审</option>
        </select>
          <%Else%>
          <input name="SetShow" type="hidden" value="<%=SetShow%>">
		  <%End If%>
        <select name="SetTop" id="SetTop">
          <option value="<%=SetTop%>" selected >顺序</option>
          <%
	  For i = 0 to 9
	      sel = ""
	    If CStr(i)=CStr(SetTop) Then
		  sel = "selected"
		End If
	  %>
          <option value="<%=i%>" <%=sel%>><%=i%></option>
          <%Next%>
        </select>
         &nbsp; &nbsp; 
        时间
        <input name="LogATime" type="text" id="LogATime" value="<%=LogATime%>" size="20" maxlength="20"> 
        &nbsp; <a href="#" onClick="owFile('../file/img_list.asp?send=PicList&ID=<%=ID%>&EdtID=EditID01')">附件管理</a></td>
    </tr>
    <%=tmrPara(1)%>
    <%
	For i=1 To ModImgCount
	If ModImgCount=1 Then
	  snPic = "附图"
	Else
	  snPic = "附图"&i
	End If
	%>
    <tr>
      <td align="center" nowrap><%=snPic%></td>
      <td>
        <IFRAME marginWidth=0 marginHeight=0 type="text/javascript" 
          src="../file/img_set.asp?TabID=<%=ModTab%>&upPath=<%=ID%>&ID=<%=ID%>&NO=<%=i%>"
          width="420" height='24' frameBorder=0 scrolling=no>
        </IFRAME>
      </td>
    </tr>
    <%Next%>
    <%
	If i=2 And Eval(ModID&"ImgSCopy")="Y" Then
	  sa = Split(Eval(ModID&"ImgSAtuo")&"x","x")
	  s1 = RequestSafe(sa(0),"N",0)
	  s2 = RequestSafe(sa(1),"N",0)
	%>
    <tr title="附图为空时,此选项有效">
      <td align="center" nowrap>提取</td>
      <td><input name="ImgSCopy" type="checkbox" id="ImgSCopy" value="Y">
        自动取内容第一个图为附图 &nbsp;&nbsp;自动缩略
        <input name="ImgSmall1" type="text" id="ImgSmall1" value="<%=s1%>" size="3" maxlength="4" style="width:30px">
        x
        <input name="ImgSmall2" type="text" id="ImgSmall2" value="<%=s2%>" size="3" maxlength="4" style="width:30px">
      0为不缩略</td>
    </tr>
    <%End If%>
    <tr>
      <td align="center" nowrap><input name="send" type="hidden" id="send" value="send">      </td>
      <td><input name=view type=button id="Button1" value="确    定" onClick="chkData()">
        <input name=view type=reset id="Button1" value="重    写">
        <input name="KW" type="hidden" id="KW" value="<%=KW%>">
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="PG" type="hidden" id="PG" value="<%=PG%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
        <input name="Act" type="hidden" id="Act" value="<%=Request("Act")%>">
      <input name="PrmFlag" type="hidden" id="PrmFlag" value="<%=PrmFlag%>"></td>
    </tr>
  </form>
</table>


<script type="text/javascript">

var tmpCont = "";
var fm = document.fm01;

function setKeys(){
  try{
	str = apiGetValEditID01();
	if(str.length==0) { alert("请先填写内容!"); }
	fm.InfPara1.value = getKeyN(str,3).replace(" ... ","");
  }catch(objError){ }
}

function owFile(url){ 
  window.open(url,"winFiles",",width=720,height=560,scrollbars=yes");
}
function owTemp(){
  var opt = fm.InfType;
  var typ = fm.InfType.value;
  window.open("set_temp.asp?tMod=<%=ModID%>&tTyp="+typ+"");
}
function getTemp() 
{
	apiSetValEditID01(tmpCont);
}

var xFile,xSize;
function owFileGet()
{
	  var tFile = xFile; 
	  var nPos = tFile.indexOf('/upfile/');
	  var sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  fm.InfPara7.value = tFile;
	  nPos = tFile.indexOf('.');
	  sSub = tFile.substring(0,nPos);
	  tFile = tFile.replace(sSub,''); 
	  fm.InfPara3.value = tFile;
	  fm.InfPara4.value = xSize;
}

aPara = "<%=InfPara%>".split("^");
for(i=0;i<96;i++) {
  try{
	eval("fm.InfPara"+i).value = aPara[i];
  } catch (er) { }
}

function chkData()
{
       
	   var eflag = 0;
       for(i=0;i<1;i++)
         {  ////////// //////////////// Start For ////////////////
 if (fm.InfSubj.value.length==0) 
   {   
     alert(" 主题 不能为空！"); 
     fm.InfSubj.focus();
     eflag = 1; break;
   }
   
   fm.InfCont.value = apiGetValEditID01(); 
 if (fm.InfCont.value.length==0) 
   {   
     //alert(" 内容 不能为空！"); 
     //eflag = 1; break;
   } 
   
 if (fm.InfCont.value.length>=480000) 
   {   
     alert(" 内容 不能超过 480 K字!"); 
     eflag = 1; break;
   }

         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		 fm.submit();
		 }
}

</script>

</body>
</html>
