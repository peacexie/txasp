<%
Dim Config_Code, Config_Path, Config_Name, Config_Neng, Config_Cont, Config_Langs, Config_URL, Config_PCopy, Config_Nstyle, Config_RAdm, Config_PCode, Config_Mode, Config_Vers
Dim intSpace, intDAGroup, intPSize, intGPSize, intLink, nPics124, nPics224, nInfN124, nBBSI124, nBBSP124, nBBSVC24
Dim AppRandom, AppRand12, AppRMemID, AppRName, App21Code, App22Code, App22Data, App23Date, App24Read, App25Code, App26Code, App27Random, App28Code, App29Code, App30Code, App31Code, App32Mode, App32Code
Dim SwhRemSave, SwhDepSubs, SwhShowSpace, SwhGuestPub, SwhGbkEditor, SwhBBSStop, SwhMemApp, SwhCodRead, SwhMemChk, SwhBBSChk, SwhGbkChk, SwhRemChk

'    Peace Para: [Config:  配置 | 系统 | 数字 | 开关]   

Config_Code = "0BB1C657-483F-00CC-22BD-PEACEV25"    '10初始化ID(SN)
Config_Path = "/"    '20系统路径
Config_Name = "贴心ASP系统(V2.5)"    '31站点中文名
Config_Neng = "(ASP 2.4)Peace Dove Web System"    '32站点英文名
Config_Cont = "DB"    '46内容存文件
Config_Langs = "1,2,3;中文,英文,日文"    '51语言版本
Config_URL = "http://127.0.0.1/txmao/txasp/"    '52首页URL地址
Config_PCopy = "PicS124,PicS224;InfN124,InfN224;Pic_999,Pic_999"    '53图片同步模块
Config_Nstyle = "01"    '62计数器风格
Config_RAdm = "adm"    '81管理入口文件名
Config_PCode = "2233444477"    '82认证码配置
Config_Mode = "isExpert"    '83系统管理模式
Config_Vers = "2.5.161016"    '84系统版本号

intSpace = 200    '11空间大小(M)
intDAGroup = 0    '12默认管理菜单组
intPSize = 36000    '21最大附件大小(KB)
intGPSize = 180    '22Guest附件大小(KB)
intLink = 12    '22连接数目
nPics124 = 2    '66产品类别级数
nPics224 = 2    '66产品类别级数
nInfN124 = 1    '66新闻类别级数
nBBSI124 = 2    '66论坛类别级数
nBBSP124 = 2    '66论坛类别级数
nBBSVC24 = 2    '88BBSVC24

AppRandom = "read_123"    '11防注随机码-Reade
AppRand12 = "app_456"    '12防注随机码-App
AppRMemID = "c714_MBD0KR9J_1ebukt"    '16防注会员ID
AppRName = "cRP0_W1V3FP_qa8g6jwf"    '17防注会员名称
App21Code = "bAJP_6P8G91EJ_np3d6d"    '21防注ChkCode
App22Code = "aJax_714422_N7XN1KDH"    '22防注AJaxCode
App22Data = "r2D6_T38U8UWG_chb73s"    '22防注AJaxData
App23Date = "20110401"    '23防注日期版本
App24Read = "App-Read-1234567"    '24防注阅读协议
App25Code = "2_1_3_9_8"    '25协议-附加码(5x1)
App26Code = "!$()*+,-/:;=?@^_`|~"    '26防注-特殊码
App27Random = "sysRand_CUE46PAJ2UXQ"    '27系统表单随机码
App28Code = "tim01S99NM1E79A"    '28限时1随机码
App29Code = "timTEEG1JCPPBWH06"    '29限时2随机码
App30Code = "r30_2MGQ851VMFYT"    '30通用限时码
App31Code = "F35N0S9DRG4EQA2CKTJVM1W87UXPH6YB"    '31通用随机码
App32Mode = "XY(Auto,Email,Mobile,ABCXYZ,Stop)"    '32A激活模式
App32Code = "DB587FFKK8FR"    '32B激活模式码

SwhRemSave = "Y"    '12保存远程图片
SwhDepSubs = "N"    '13分公司(部门)
SwhShowSpace = "N"    '13显示占用空间
SwhGuestPub = "Y"    '21Guest发布
SwhGbkEditor = "Y"    '22留言多媒体框
SwhBBSStop = "N"    '23论坛系统关闭
SwhMemApp = "Y"    '31允许会员注册
SwhCodRead = "Y"    '32协议认证码
SwhMemChk = "Y"    '41自动审核会员
SwhBBSChk = "Y"    '42自动审核帖子
SwhGbkChk = "Y"    '43自动审核留言
SwhRemChk = "N"    '44自动审核评论
%>