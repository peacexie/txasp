<link rel="stylesheet" type="text/css" href="style.css">
<table width="100%" border='0' cellpadding='0' cellspacing='1' bgcolor='#CCCCFF'>
    <tr bgcolor='#DDDDFF'>
      <td align="left"> 选择查看对象
        　 公开
        <input name="sAll01" type="radio" value="(__Public__)" onclick="ChkAll('Pub')" <%If InfTo="(__Public__)" Or PrmFile="info_add.asp" Then Response.Write("checked='checked'")%> />        
        　 全选
        <input name="sAll01" type="radio" value="Y" onClick="ChkAll('All')">
         　  
        取消全选        
        <input name="sAll01" type="radio" value="N" onclick="ChkAll('Nul')" /></td>
      <td align="center" nowrap="nowrap">全选</td>
      <td align="center" nowrap="nowrap">部门</td>
    </tr>
    <%=GetUList(InfTo)%>
</table>
