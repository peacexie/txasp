<%
'    Peace Para: [tmsPara.asp]模版参数   
'//////////////////////////////////////////////////////////////////////////////////////

Dim PicT124ImgCount,PicT125ImgCount
Dim PicT124ImgRelat
' tmsPics.asp 图片个数/相关图片
' 设置某个模块 图片个数; 请严格按现有模式填写!
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！

PicT124ImgCount = 3 
PicT125ImgCount = 0
PicV125ImgCount = 0
TraN124ImgCount = 0
TraJ124ImgCount = 0

PicT124ImgRelat = true

'自动缩略图
Dim PicR124ImgSCopy,PicR124ImgSAtuo : PicR124ImgSCopy = "Y" : PicR124ImgSAtuo = "90x120"

'//////////////////////////////////////////////////////////////////////////////////////

Dim tmsListTab(24) ' 列表模版 Cont,Next,FAQ,NTab,PTab,News,Pics,Jobs,Vdos
' 设置某个模块/类别 列表信息 显示方式; 请严格按现有模式填写! 
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！

tmsListTab(1) = "Cont:TraA124;A110012;A110040;D110024;C110024;" '单页
tmsListTab(2) = "Next:A110020;" '上下页
tmsListTab(3) = "FAQ:A110016;"  '问答
tmsListTab(4) = "NTab:A110024;A110026;" '列表
tmsListTab(5) = "PTab:D110028;C110028;" '图片

tmsListTab(6) = "PicA:;" '宽4:3 --- 2x3列行
tmsListTab(7) = "PicB:;" '长3:4 --- 3x2列行
tmsListTab(8) = "PicC:A110028;" '一行一个 lightbox 特效

tmsListTab(11) = "News:TraN124;" '新闻
tmsListTab(12) = "Jobs:TraJ124;A110030;" '职位
tmsListTab(13) = "Pics:TraT124;PicS124;PicS224;PicR124;PicT124;" '产品
tmsListTab(14) = "Vdos:PicV124;" '视频
tmsListTab(15) = "Down:PicV125;" 'Down


'//////////////////////////////////////////////////////////////////////////////////////
Dim tmsLinkTab(24) ' 连接模版
' 设置某个模块/类别 列表信息 连接方式; 请严格按现有模式填写! 
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！
' InfPara (1~8) 
'通用:  InfKey,InfFrom,InfSpeci,InfPrice,
'       InfPric2,InfDate,InfUrl,InfRem,
'新闻:  .Key,来源,规格,价格,
'       原价,日期,地址,.Rem,
'下载:  导演,版权,格式,大小,
'       片长,日期,附件,主演,
'  酒:  ----,产地,规格,价格
'       原价,年份,级别

tmsLinkTab(1) = "(Null):R110020;" '无连接
tmsLinkTab(2) = "(ImgNam1):R110024;" '连接图片
tmsLinkTab(3) = "(InfPar2):A110026;"  '外部连接

tmsLinkTab(4) = "[dir/:;"  '目录 用于生成静态...
tmsLinkTab(5) = "[dir/index.stm:;"  '目录+index.stm 用于生成静态...
tmsLinkTab(6) = "[dir/index.asp:;"  '目录+index.asp 用于生成静态...

tmsLinkTab(9) = "iview.asp?KeyID=(KeyID):InfA124;" '本页 &InfType=(InfType)
tmsLinkTab(10) = "?KeyID=(KeyID):TraA124;" '公共
tmsLinkTab(11) = "nview.asp:" '新闻
tmsLinkTab(12) = "pview.asp:" '图片

'//////////////////////////////////////////////////////////////////////////////////////

Dim tmsShowTab(24) ' 显示模版
' 设置某个模块/类别 详细信息 显示方式; 请严格按现有模式填写! 
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！

tmsShowTab(1) = "1News:" '经典新闻：标题+内容
tmsShowTab(2) = "1Pics:TraT124;PicS124;" '经典产品：标题+内容+参数+图片
tmsShowTab(3) = "1Pic3:PicT124;" '标题+内容+3图片+相关图片
tmsShowTab(4) = "1Vdos:PicV124;" '视频+内容 

tmsShowTab(6) = "6UD:TraA124;R110010;" '图片+内容(上下)
tmsShowTab(7) = "6Left:R110012;" '图片(左)+内容
tmsShowTab(8) = "6Right:R110016;" '图片(右)+内容

tmsShowTab(11) = "2Pics:S110016;" '产品无订购：标题+内容+参数+图片 magic zoom特效

'//////////////////////////////////////////////////////////////////////////////////////

Dim tmsRemTab(6) ' 评论模版
' 设置某个模块/类别 评论 显示方式; 请严格按现有模式填写! 
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！

tmsRemTab(1) = "Y:;" '有评论
tmsRemTab(2) = "N:;" '无评论
tmsRemTab(3) = "X:Y;" '默认

'//////////////////////////////////////////////////////////////////////////////////////

Dim tmsVoteTab(6) ' 投票模版
' 设置某个模块/类别 投票 显示方式; 请严格按现有模式填写! 
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！

tmsVoteTab(1) = "Y:InfN124;" '有投票
tmsVoteTab(2) = "N:;" '无投票
tmsVoteTab(3) = "X:Y;" '默认

'//////////////////////////////////////////////////////////////////////////////////////

Dim tmsNextTab(6) ' 上下篇模版
' 设置某个模块/类别 上下篇 显示方式; 请严格按现有模式填写! 
' 正常运行的系统，请不要随便执行该操作！否则后果自负！！！

tmsNextTab(1) = "M:PicS124;" '模块上下篇
tmsNextTab(2) = "T:InfN124;" '类别上下篇
tmsNextTab(3) = "X:InfD124;" '无上下篇
tmsNextTab(4) = "D:M;" '默认

'//////////////////////////////////////////////////////////////////////////////////////
%>