<!--#include file="_config.asp"-->
<%
'Response.Write vMem_Btn01

Call sf_Guard()
Call ChkSpider(1)
Call Chk_URL()
Call App25Set()
goPage = Request("goPage")


ChkCod11 = Rnd_ID("KEY",24)
ChkCod12 = MD5_User(ChkCod11,AppRandom)
ChkCod21 = Now()
ChkCod22 = MD5_User(ChkCod21,"Set1234")

If SwhCodRead = "Y" Then
  App24Code = Rnd_ID("0",4)
  Session("App24Code") = App24Code
  flgApp24 = "XXX"
  cmdApp24 = ""
Else
  flgApp24 = ""
  cmdApp24 = "ChkAj22('ChkAj22')"
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<TITLE><%=" [ "&vPMsg_WName&" ] "&vMem_RD12&" "%></TITLE>
<link rel="stylesheet" type="text/css" href="../pfile/pimg/style.css">
<meta name="robots" content="noindex,nofollow,noarchive" />
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../sadm/func1/Func_JS.js" type="text/javascript"></script>
<style type="text/css">
.gapLeft {
   padding-left:60px;	
}
.gapRight {
   padding-left:30px;	
}
</style>
</HEAD>
<BODY>
<%TPName=vMem_Btn01%>
<%
If SwhMemApp<>"Y" And Session("UsrID")&""="" Then 
  Response.Write "<br><br><center>"&vMem_RD21&"</center><br><br>"
  Response.End()
End If

'rDir = Config_Path&"member/mu_"&AppRand12&".asp"
'If Chk_URL3(rDir)="eUrl" Then
  'Response.Redirect "/"
'End If
'Response.Write "<br>ar_xxxx:"&rDir&Chk_URL3(rDir)

%>
<!--#include file="_ftop2.asp"-->
      <table width="520" border="0" align="center" cellpadding="2" cellspacing="1">
        <form action="mu_<%=AppRand12%>.asp" method="post" name="fm01" id="fm01">
          <%
		  Msg = Request("Msg")
		  If Msg<>"" Then 
		  %>
          <tr align="middle" bgcolor="#FFFFDD">
            <td height="30" colspan="2" align="center"><%=vMem_RD11%><%=Msg%></td>
          </tr>
          <%End If%>
          <tr align="middle">
            <td colspan="2" align="center"><span style="line-height:180%;"><%="[ "&vPMsg_WName&" ]"%><%=vMem_RD12%></span></td>
          </tr>
          <tr align="middle">
            <td colspan="2" align="center"><div style="width:480px; height:300px; overflow-y:auto; text-align:left; padding:5px; border:1px solid #CCC;">
              <%If verMemb="2" Then%>
              <!--#include file="../upfile/sys/para/rread2.txt"-->
              <%Else%>
			  <!--#include file="../upfile/sys/para/rread.txt"-->
              <%End If%>
            </div></td>
          </tr>
          <tr align="middle">
            <td height="30" align="left" nowrap="nowrap" class="gapLeft"><input type="radio" name="ChkFlag" id="ChkFlag1" value="OK" onclick="ChkVals(this);<%=cmdApp22%>" />
              <%=vMem_RD13%>&nbsp;&nbsp;&nbsp;
              <input name="ChkFlag" type="radio" id="ChkFlag2" onclick="ChkVals(this)" value="XX" checked="checked" />
              <%=vMem_RD14%></td>
            <td align="left" nowrap="nowrap" class="gapRight"><select name="MemType" id="MemType" style="width:130px;" <%=flgApp22%>onchange="ChkAj22('ChkAj22')">
                <option value="">(<%=vMem_RD15%>)</option>
                <%=Get_SOpt(mCfgCode,mCfgName,Show_Text(Request("MemType")),"")%>
              </select></td>
          </tr>
          <%If SwhCodRead = "Y" Then%>
          <tr align="middle">
            <td align="left" nowrap="nowrap" class="gapLeft"><%=vMem_RD25%><%=App24Code%> (<%=vMem_RD26%>)</td>
            <td align="left" nowrap="nowrap" class="gapRight">
              <input name='App24Code' type='text' id="App24Code" onblur="ChkAj22('ChkAj22')" size='4' maxlength='8' />
            (<%=vMem_RD27%>)<span class="pt9"> 
            </span></td>
          </tr>
          <%End If%>
          <tr align="middle">
            <td width="50%" align="left" nowrap="nowrap" class="gapLeft"><input type="button" name="btnOK" value="<%=vMem_RD17%>" id="btnOK" style="width:130px;" onclick="chkData()" /></td>
            <td align="left" nowrap="nowrap" class="gapRight"><input type="button" name="btnXX" value="<%=vMem_RD18%>" id="btnXX" style="width:130px;" onclick="Dir_Addr('login.asp')" /></td>
          </tr>
          <tr align="middle">
            <td colspan="2" align="left" nowrap="nowrap">&nbsp;&nbsp;&nbsp;<%=vMem_RD19%><span style="line-height:180%;">[<a href="?goPage=<%=goPage%>&verMemb=<%=verMemb%>&Act=Clear"><%=vMem_Btn13%></a>] </span></td>
          </tr>
          <tr align="middle">
            <td colspan="2">&nbsp;</td>
          </tr>
          <input name="<%=app30Tab(1)%>" type="hidden" id="<%=app30Tab(1)%>" value="<%=ChkCod11%>" />
          <input name="<%=app30Tab(3)%>" type="hidden" id="<%=app30Tab(3)%>" value="<%=ChkCod12%>" />
          <input name="<%=app30Tab(5)%>" type="hidden" id="<%=app30Tab(5)%>" value="<%=ChkCod21%>" />
          <input name="<%=app30Tab(7)%>" type="hidden" id="<%=app30Tab(7)%>" value="<%=ChkCod22%>" />
          <%If Chk_URL3(Config_Path&"member/mu_app2.asp")="OK" Then%>
          <input name="MemID" type="hidden" id="MemID" value="<%=Show_Text(Request("MemID"))%>" />
          <input name="MemQu" type="hidden" id="MemQu" value="<%=Show_Text(Request("MemQu"))%>" />
          <input name="MemAn" type="hidden" id="MemAn" value="<%=Show_Text(Request("MemAn"))%>" />
          <input name="MemName" type="hidden" id="MemName" value="<%=Show_Text(Request("MemName"))%>" />
          <input name="MemNam2" type="hidden" id="MemNam2" value="<%=Show_Text(Request("MemNam2"))%>" />
          <input name="MemTyp2" type="hidden" id="MemTyp2" value="<%=Show_Text(Request("MemTyp2"))%>" />
          <input name="MemCard" type="hidden" id="MemCard" value="<%=Show_Text(Request("MemCard"))%>" />
          <input name="MemFrom" type="hidden" id="MemFrom" value="<%=Show_Text(Request("MemFrom"))%>" />
          <input name="MemBirth" type="hidden" id="MemBirth" value="<%=Show_Text(Request("MemBirth"))%>" />
          <input name="MemTel" type="hidden" id="MemTel" value="<%=Show_Text(Request("MemTel"))%>" />
          <input name="MemMobile" type="hidden" id="MemMobile" value="<%=Show_Text(Request("MemMobile"))%>" />
          <input name="MemEmail" type="hidden" id="MemEmail" value="<%=Show_Text(Request("MemEmail"))%>" />
          <%End If%>
          <input name="ChkFRes" type="hidden" id="ChkFRes" value="---" />
          <input name="AppRCode" type="hidden" id="AppRCode" value="(-)" />
          <input name="<%=App22Code%>" type="hidden" id="<%=App22Code%>" value="(-)" />
          <input name='<%=App25A%>' type="hidden" value="<%=App25B%>" />
          <input name="<%=App25C%>" type="hidden" value='<%=App25D%>' />
          <input name='goPage' type='hidden' id='goPage' value='<%=goPage%>'>
          <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=Rnd_Base("5678",9)&Rnd_Base("",64)%>" />
          <input name="verMemb" type="hidden" id="verMemb" value="<%=verMemb%>" />
          
          <input name="<%=App30Code%>0" type="hidden" id="<%=App30Code%>0" value="<%=app30Arr(0)%>" />
          <input name="<%=App30Code%>1" type="hidden" id="<%=App30Code%>1" value="<%=app30Arr(1)%>" />
          <input name="<%=App30Code%>2" type="hidden" id="<%=App30Code%>2" value="<%=app30Arr(2)%>" />
          
          <input name="_trakeTime" type="hidden" value='<%=Now()%>' />
          
        </form>
      </table>
      <script type="text/javascript">

//Reset Form

var xmlHttp = getXmlHttp(); //ChkAjID,ChkAjCode
function ChkAj22(xType) {
  var url = "check_id.asp?yAct=ChkAj22"; 
  xmlHttp.open("GET", url, true); //这里的true代表是异步请求 ChkAjCode
  xmlHttp.onreadystatechange = ChkAjUpd;
  xmlHttp.send(null);
}
function ChkAjUpd(){
  if (xmlHttp.readyState == 4) {
	var rData = xmlHttp.responseText; 
	document.fm01.<%=App22Code%>.value = rData;
	//var rMsg = "";
	//if(rData=="N.Code"){ rMsg = "认证码 错误！";}
	//if(rData.substring(0,2)=='N.'){alert(rMsg);}
  }
}

function App_Read()
{
    AppRCode = "<%=App24Read%>";
	AppMsg = "<%=vMem_RD22%>";
	if(confirm(AppMsg))
     {
		 document.getElementById('AppRCode').value = AppRCode;
         return true;
     }
         location.href = "login.asp";
		 return false;
}
//App_Read();

function ChkVals(e)
{
	var eFlag = e.value;
	if (eFlag=='OK'){
		document.getElementById('btnOK').disabled = false;
		document.getElementById('ChkFRes').value = "OK";
		document.getElementById('AppRCode').value = "<%=App24Read%>";
	}else{
		document.getElementById('btnOK').disabled = true;
		document.getElementById('ChkFRes').value = "XX";
		document.getElementById('AppRCode').value = "-";
	}
}
document.getElementById('btnOK').disabled = true;


 function chkData()
 {
       var eflag = 0;
       for(ii=0;ii<1;ii++)
         {  ////////// //////////////// Srart For ////////////////
 
 if (document.fm01.MemType.value.length==0)
   {   
     alert('<%=vMem_RD23%>');
     document.fm01.MemType.focus();
     eflag = 1; break;
   }
 <%If SwhCodRead = "Y" Then%>  
 if (document.fm01.App24Code.value.length==0)
   {   
     alert('<%=vMem_RD24%>');
     document.fm01.App24Code.focus();
     eflag = 1; break;
   }
 <%End If%>
         }  ////////// //////////////// End For //////////////////
         if (eflag==0){ 
		   document.getElementById('ChkFlag2').checked = true;
		   document.getElementById('btnOK').disabled = true;
		   document.fm01.submit(); 
		 }
}

if(top.location!==self.location){top.location=self.location;}

</script>
      <!--#include file="_fbot2.asp"-->
</BODY>
</HTML>
