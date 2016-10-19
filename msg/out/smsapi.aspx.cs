using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class send_smsapi : System.Web.UI.Page
{
    
	public static string smaTitle = "短信接口API";     //短信接口
	public static string smaState = "isOpen";          //状态(isStop,isOpen)
	public static string smaUser = "eweb";             //接口帐号(预设)
	public static string smaCode = "0BA9A93D-6024-00CA-0F79-1155AB078E8F";             //接口SN(预设)
	public static string smaAdda = "http://peace.96327.com/msg/user/sapi.asp";         //接口地址(远程)
	public static string smaAddr = "eweb.dg.gd.cn/member/send/smsapi.aspx";            //发送地址(本地)
	public static string smaTels = "13537432147";            //发送号码(网管)
	
	public static string id = smaUser; 
	public static string snOrg = smaCode; 
	public static string smsABase = "http://peace.96327.com/"; 
	
	
	protected void Page_Load(object sender, EventArgs e)
    {
        
		//权限设置.
		CheckUrl("/member/");
		
		//act=AppOK
		//ID=ID
		string ID = Request.Params["ID"];
		
		
		string tm = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); //Now()
		string sn = outEncSN(id,tm,snOrg); 
		string urlAdd0 = smaAdda; //(...&"msg/user/sapi.asp");
		string urlPar0 = "?tm="+tm+"&id="+id+"&sn="+sn+""; //'Peace_Sms_RndID="&Timer&"&
		
		string tel = smaTels;
		string ct1 = "自助建站:注册一个新会员("+ID+"). asp.net-utf-8-peace测试";
		string csp = "&cs=(out)"; ct1 = outEncode(ct1);
		string urlPara = urlPar0+"&act=Send&tel="+tel+"&ct1="+ct1+csp+"";
		
		string frm = "<IFRAME name=LeftMenu src='"+urlAdd0+urlPara+"' frameBorder=0 width='420' scrolling='no' height='100%'></IFRAME>";
		
		string msg = "<a href='../login.aspx?ID="+ID+"'>请返回登陆...</a>";
		this.lbFlag.Text = msg;
		this.lbFrame.Text = frm;
		
		/*
		string sFlag = "Run ASP.NET OK! \n";
		sFlag += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
		this.lbFlag.Text = sFlag;
		this.lbFlag.Text += "<br>"+outEncode("测试byPeace鸽子!");
		this.lbFlag.Text += "<br>"+outEncSN("demo","2011-4-25 11:11:13","0BA9A93D-7304-00CA-0F79-B25FAB273C1B");
		this.lbFlag.Text += "<br>"+System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
		*/
		
    }
	
    public string outEncode(string xStr)
    {
		string s=""; int iAsc;
		char[] aStr = xStr.ToCharArray(0, xStr.Length);
		for (int i=0;i<aStr.Length;i++)
		{
			iAsc = Convert.ToInt32(aStr[i]);
			s += iAsc+";";
		}
		return s;
	}
	
    public string outEncSN(string xid,string xtm,string xsn)
    {
		// using System.Security.Cryptography;
		//string snOrg = "";
		//string smaAddr = "";
		string str = xid+"+"+snOrg+"+"+xtm+"+"+smaAddr;
		str = FormsAuthentication.HashPasswordForStoringInConfigFile(str, "md5");
		return str.ToLower();
	}
	
	public static string CheckUrl(string xDirPage)
	{
		////
	}
	
	
}


