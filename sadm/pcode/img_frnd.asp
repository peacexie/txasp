<!--#include file="../../sadm/func2/func_const.asp"-->
<%

'' iNum=Int((iMax - iMin + 1) * Rnd + iMin)
'' 2233444477(2:2:4:2)数字,字母,运算; 234
'' 5566(2:2)汉字
'' hhiijj(2:2:2)小号数字,字母
'' x(1)OUT数字,字母
'' 234567hijx(全部)
If Request("Config_PCode")<>"" Then
  Config_PCode = "hij"
End If
'Config_PSess 

Randomize 
Dim CfgStr  :CfgStr  = Config_PCode '"234567hijx" '223444477
Dim CfgMax  :CfgMax  = Len(CfgStr)
Dim CfgNum  :CfgNum  = Int((CfgMax)*Rnd + 1)
Dim CfgChar :CfgChar = Mid(CfgStr,CfgNum,1)
Dim CfgFile :CfgFile = "img_cod"&CfgChar&".asp"
'' Response.Write CfgChr
'' Response.Redirect CfgFile
Server.Transfer(CfgFile)
'' Server.Execute(CfgFile) 

'' Server.Transfer 比 Response.Redirect 快很多！！！ 
'' Server.Transfer(),Response.Redirect(),Server.Execute()的区别 
'' http://hi.baidu.com/%CC%A4%C0%CB%CB%A7/blog/item/cb351c3e0353d5cb7c1e71eb.html
'' 如果要将执行流程转入同一Web服务器的另一个ASPX页面，应当使用Server.Transfer，能够避免不必要的网络通信。       
'' 如果要捕获一个ASPX页面的输出结果，然后将结果插入另一个ASPX页面的特定位置，则使用Server.Execute。       
'' 如果要确保HTML输出合法，请使用Response.Redirect，不要使用Server.Transfer或Server.Execute方法。 

'' 2: H28 : 2~4 : 2AC3EF4HJ5KL6NP7RST8UVW9XYZ
'' 3: H28 : 3~5 : 数字3~5 Layen support@ssaw.net 84815733(QQ) 
'' 4: H28 : 2~4 : 数字+字母
'' 5: H28 : 2~4 : 汉字350个 4改进
'' 6: H28 : 2~4 : 汉字85个"好一路..." DV-BBS
'' 7: H28 : 5~6 : 99+9=? 数字运算 4改进
'' h: H11 : 2~4 : 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ@:/.-_^
'' i: H10 : 4 : 数字,黑色
'' j: H10 : 4 : 数字,彩色
'' x: H22 : 3~5 : .NET : OUT eweb.dg.gd.cn 

%>
