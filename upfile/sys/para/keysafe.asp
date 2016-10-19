<%

ParFKeyStr = ""
ParFIPStr = ""
ParFUrlStr = ""

Config_sfIPAct = "Stop" 'Trace:跟踪,不影响浏览; Stop:禁止IP,不能正常浏览; Exit:不启用；
Config_sfInjAct = "Trap" 'Stop:停止操作,显示空白; Warn:警告信息; Trap:反攻击; Exit:不启用； 

ParFKeyStr = ParFKeyStr&"and char(|and exists ( )|ascii( substring(|len( select from|update set =|alter database full|"
ParFKeyStr = ParFKeyStr&"drop table|grant exec xp_|declare char( @|backup disk \|backup database \|create tabel (|"
ParFKeyStr = ParFKeyStr&"insert xp_subdirs (|insert sqloledb dbo.|insert exec xp_|insert values(|select dbo sys|select openrowset(|"
ParFKeyStr = ParFKeyStr&"select sh"&"ell(|select count(*)|select sqloledb (|select cmd.exe (|exec sp_|exec xp_|"
ParFKeyStr = ParFKeyStr&"exec \|xp_getfiledetails \|xp_makecab \|delete from|and > ;|and = ;|"
ParFKeyStr = ParFKeyStr&"and = '|and < ;|or = '|execute request(|cmd.exe /c request|script /script|"
ParFKeyStr = ParFKeyStr&"iframe /iframe|" 

ParFIPStr = ParFIPStr&"61.170.150.187|61.170.156|61.142.246.153|61.151.239.174|127.0.X.1|58.217.104.41|"
ParFIPStr = ParFIPStr&"58.47.108.63|58.251.29.70|59.41.253.7|59.42.176.149|60.255.76.11|61.151.239.174|"
ParFIPStr = ParFIPStr&"61.166.144.147|74.222.6.95|116.253.245.199|116.11.32.178|117.36.239.130|118.113.89.11|"
ParFIPStr = ParFIPStr&"120.0.91.11|121.34.124.178|121.124.77.124|121.8.69.98|122.4.208.14|123.123.131.134|"
ParFIPStr = ParFIPStr&"124.118.168.210|202.96.96.162|202.70.55.160|211.153.19.240|219.157.96|219.113.217.42|"
ParFIPStr = ParFIPStr&"219.148.50.162|221.231.138.90|221.212.165.100|222.222.50.179|" 

ParFUrlStr = ParFUrlStr&"/|http://www.dg.gd.cn/|http://www.dgchr.com/|"
ParFUrlStr = ParFUrlStr&"http://www.changanedu.com/|http://www.txjia.com/|" 
%>