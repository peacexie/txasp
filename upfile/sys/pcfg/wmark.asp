<%
'    Peace Para: [wmark]   
Dim wmk_Mark, wmk_Logo, wmk_Trans, wmk_TColor, wmk_TScope, wmk_Text, wmk_Pos, wmk_Padding, wmk_Color2, wmk_Color1, wmk_Size, wmk_Font, wmk_Width, wmk_Height, wmk_HLine, wmk_Notes10, wmk_Notes40
wmk_Mark = "X"    '10水印控制
wmk_Logo = "upfile/myfile/logo/logo-peace.gif"    '20水印图片
wmk_Trans = "0.8"    '23图片透明度(0.1~1.0)
wmk_TColor = "&H0000FE"    '25图片透明颜色(&RGB)
wmk_TScope = "72"    '27图片透明范围(1~65536)
wmk_Text = "[X2X测试网站]|domain.com"    '30水印文字
wmk_Pos = "3"    '41水印位置
wmk_Padding = "10"    '42水印边距
wmk_Color2 = "&HFFFFFF"    '51描边颜色
wmk_Color1 = "&H00FF00"    '51文字颜色
wmk_Size = "24"    '52文字大小
wmk_Font = "Arial"    '53文字字体
wmk_Width = "150"    '56文字宽度
wmk_Height = "60"    '57文字高度
wmk_HLine = "25"    '58文字行高(2行)
wmk_Notes10 = "T:文字; P:图片; X:不启用"    'x10控制说明
wmk_Notes40 = "0:中间; 1234:左上右下顺时针"    'x40位置说明
%>