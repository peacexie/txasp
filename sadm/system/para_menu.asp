<a href='upd_para.asp'><font color="#0000FF">刷新</font></a> | 
<%
  fExpert = rs_Val("","SELECT ParText FROM AdmPara WHERE ParCode='Config_Mode'")
  If fExpert="isExpert" Or Chk_PermSP() Then 
%>
<!--#include file="../../upfile/sys/config/sf_Para.htm"-->
<%Else%> 
<a href=para_yno.asp?ModID=ParTYN>开关</a> | 
<a href=para_text.asp?ModID=ParText>随机</a> | 
<a href=para_date.asp?ModID=ParDate>日期</a> | 
<a href=para_text.asp?ModID=ParLink>联系</a> | 
<!--<a href=para_text.asp?ModID=ParEmail>Jmail</a> |  -->
<a href=para_rem.asp?ModID=ParSEO>SEO</a> | 
<a href=para_rem.asp?ModID=ParTemp>模版</a> | 
<!--<a href=para_rem.asp?ModID=ParBBS>论坛</a> | -->
<a href=para_rem.asp?ModID=ParFil>过滤</a> | 
<!--<a href=para_rem.asp?ModID=ParRe2>说明</a> | --> 
<a href=para_rem.asp?ModID=ParRem>备注</a> | 
<%End If%>