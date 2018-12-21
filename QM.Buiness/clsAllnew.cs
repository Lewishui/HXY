using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using Gecko;
using ISR_System;
using QM.DB;
using System.Runtime.InteropServices;
using System.IO;
using WEBIBM;
using System.Threading;
using System.Text.RegularExpressions;
using mshtml;
using System.ComponentModel;
using China_System.Common;
using System.Diagnostics;

namespace QM.Buiness
{
    public enum ProcessStatus
    {
        初始化,
        登录界面,
        确认YES,
        第一页面,
        第二页面,
        Filter下拉,

        结束页面

    }
    public enum EapprovalProcessStatus
    {
        初始化,
        Current_Task,
        Task_Queue,
        Search,
        Process
    }


    public class clsAllnew
    {
        private readonly string xulrunnerPath = Application.StartupPath + "\\xulrunner";
        private string testUrl = "https://publish.caasdata.com/";
        private ProcessStatus isrun = ProcessStatus.初始化;
        private EapprovalProcessStatus isrun1 = EapprovalProcessStatus.初始化;
        private bool isOneFinished = false;
        private bool isOneFinished_VI = false;
        private Form viewForm;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        private WbBlockNewUrl MyWebBrower;
        WbBlockNewUrl myDoc = null;
        GeckoWebBrowser GmyDoc = null;
        private GeckoWebBrowser Browser;
        private DateTime strFileName;
        List<string> xuanzhezhanghao;
        private bool isReadyForSearch = false;
        string FileDownURL = string.Empty;
        string FileName = "";
        private string publicPDFName;
        static nsICookieManager CookieMan;
        public int run_type;
        public BackgroundWorker bgWorker1;
        public List<clsDatabaseinfo> RResult;
        public clsDatabaseinfo Item_RResult;
        int yinhangrun_index = 0;
        private DateTime StopTime;
        #region Import API
        System.Timers.Timer aTimer = new System.Timers.Timer(100);//实例化Timer类，设置间隔时间为10000毫秒； 
        System.Timers.Timer t = new System.Timers.Timer(1000);//实例化Timer类，设置间隔时间为10000毫秒； 

        private int ScreenStatus, intCnt;
        private bool RUNING = false;
        private const int WM_KEYDOWN = 0x100;
        private const int WM_KEYUP = 0x101;
        private const int VK_TAB = 0x9;
        private const int VK_CONTROL = 0x11;
        private const int VK_PRIOR = 0x21;
        private const int VK_UP = 0x26;
        private const int VK_HOME = 0x24;
        private const int BM_CLICK = 0xF5;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int SYSKEYDOWN = 0x104;
        private const int WM_SETTEXT = 0x000C;
        private bool WebSiteStatus = false;
        private bool IntialFinish = false;
        private IntPtr hwnd_main, hwnd_ReportTree, hwnd_ReportTree1, hwnd_Control;
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint = "GetParent")]
        public static extern IntPtr GetParent(IntPtr hwndChild);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);


        [DllImport("User32.dll ")]
        public static extern IntPtr GetDlgItem(IntPtr parent, long childe);


        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);
        #endregion
        #region  Webbroswer
        public List<clsWEBINFO> ReadWEN()
        {

            try
            {
                InitialWebbroswer();

                HtmlElement Scope = null;
                HtmlElement View = null;
                int io = 0;

                while (!isOneFinished)
                {
                    System.Windows.Forms.Application.DoEvents();


                }
                System.Windows.Forms.Application.DoEvents();
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);

                throw;
            }
        }
        public void InitialWebbroswer()
        {
            try
            {

                MyWebBrower = new WbBlockNewUrl();
                //不显示弹出错误继续运行框（HP方可）
                MyWebBrower.ScriptErrorsSuppressed = true;
                #region new add

                //MyWebBrower.AllowWebBrowserDrop = false;
                //MyWebBrower.IsWebBrowserContextMenuEnabled = false;
                //MyWebBrower.WebBrowserShortcutsEnabled = false;
                //MyWebBrower.ObjectForScripting = this;
                //Uncomment the following line when you are finished debugging.
                //webBrowser1.ScriptErrorsSuppressed = true;

                //MyWebBrower.DocumentText =
                //  "<html><head><script>" +
                //  "function test(message) { alert(message); }" +
                //  "</script></head><body><button " +
                //  "onclick=\"window.external.Test('called from script code')\">" +
                //  "call client code from script code</button>" +
                //  "</body></html>";

                #endregion
                MyWebBrower.BeforeNewWindow += new EventHandler<WebBrowserExtendedNavigatingEventArgs>(MyWebBrower_BeforeNewWindow);
                MyWebBrower.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(AnalysisWebInfo);
                MyWebBrower.Dock = DockStyle.Fill;

                //显示用的窗体
                viewForm = new Form();
                //viewForm.Icon=
                viewForm.ClientSize = new System.Drawing.Size(550, 600);
                viewForm.StartPosition = FormStartPosition.CenterScreen;
                viewForm.Controls.Clear();
                viewForm.Controls.Add(MyWebBrower);
                viewForm.FormClosing += new FormClosingEventHandler(viewForm_FormClosing);
                //显示窗体

                viewForm.Show();

                MyWebBrower.Url = new Uri("https://publish.caasdata.com/");
                //MyWebBrower.Url = new Uri("https://www.baidu.com/");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private void viewForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // if (tsStatusLabel1.Text != " Search Finished  !")
            {
                //if (MessageBox.Show("正在进行，是否中止?", "CCW", MessageBoxButtons.OKCancel) == DialogResult.OK)
                //{
                //    if (MyWebBrower != null)
                //    {
                //        if (MyWebBrower.IsBusy)
                //        {
                //            MyWebBrower.Stop();
                //        }
                //        MyWebBrower.Dispose();
                //        MyWebBrower = null;
                //    }
                //}
                //else
                //{
                //    e.Cancel = true;
                //}
            }
        }

        void MyWebBrower_BeforeNewWindow(object sender, WebBrowserExtendedNavigatingEventArgs e)
        {
            #region 在原有窗口导航出新页
            e.Cancel = true;//http://pro.wwpack-crest.hp.com/wwpak.online/regResults.aspx
            MyWebBrower.Navigate(e.Url);
            #endregion
        }
        protected void AnalysisWebInfo(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //  WbBlockNewUrl myDoc = sender as WbBlockNewUrl;
            myDoc = sender as WbBlockNewUrl;

            if (myDoc.Url.ToString().IndexOf("https://publish.caasdata.com/") >= 0 && isrun == ProcessStatus.初始化)
            {




            }
        }
        #endregion
        public clsAllnew()
        {
            //Initialize(null);
            //InitializeExtras();   
        }
        #region GeckoWebBrowser

        public List<clsWEBINFO> ReadGeckoWEN(List<string> newlisttime)
        {

            try
            {
                xuanzhezhanghao = new List<string>();
                xuanzhezhanghao = newlisttime;


                Xpcom.Initialize(xulrunnerPath);
                isrun = ProcessStatus.初始化;
                if (run_type == 1)
                {
                    InitAllCityData();
                    while (!isOneFinished)
                    {
                        //   GmyDoc.Refresh();
                        if (run_type == 1)
                        {
                            if (isrun == ProcessStatus.确认YES)
                                NewMethod();
                        }
                        else if (run_type == 2)
                        {
                            if (isrun == ProcessStatus.确认YES)
                                readqiyemingcheng();


                        }
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                else if (run_type == 2)//查询
                {

                    yinhangrun_index = 0;
                    foreach (clsDatabaseinfo item in RResult)
                    {
                        isOneFinished = false;
                        isrun = ProcessStatus.初始化;



                        Item_RResult = item;
                        InitAllCityData();
                        while (!isOneFinished)
                        {

                            if (run_type == 2)
                            {
                                if (isrun == ProcessStatus.确认YES)
                                    readqiyemingcheng();


                            }
                            System.Windows.Forms.Application.DoEvents();
                        }
                        item.xuyaoyanzhengdegongsiming = Item_RResult.xuyaoyanzhengdegongsiming;
                        yinhangrun_index++;
                        //bgWorker1.ReportProgress(0, "正在读取   :  " + item.zhanghao + "  " + yinhangrun_index.ToString() + "/" + RResult.Count.ToString());
                        tsStatusLabel1.Text = "正在读取   :  " + item.zhanghao + "  " + yinhangrun_index.ToString() + "/" + RResult.Count.ToString();

                        if (viewForm != null)
                        {

                            MyWebBrower = null;
                            viewForm.Close();

                        }
                    }
                    if (MyWebBrower != null)
                    {
                        if (MyWebBrower.IsBusy)
                        {
                            MyWebBrower.Stop();
                        }
                        MyWebBrower.Dispose();
                        MyWebBrower = null;
                    }
                }
                if (run_type == 3)
                {
                    isOneFinished = false;

                    isrun1 = EapprovalProcessStatus.初始化;
                    InitAllCityData();
                    while (!isOneFinished)
                    {
                        //   GmyDoc.Refresh();

                        if (run_type == 3)
                        {
                            if (isrun == ProcessStatus.确认YES || isrun == ProcessStatus.第一页面)
                                erici_fuzhi();


                        }
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                if (MyWebBrower != null)
                {
                    if (MyWebBrower.IsBusy)
                    {
                        MyWebBrower.Stop();
                    }
                    MyWebBrower.Dispose();
                    MyWebBrower = null;
                }
                if (viewForm != null)
                {

                    MyWebBrower = null;
                    viewForm.Close();

                }
                return null;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show("" + ex);
                return null;
                throw;
            }
        }
        private void InitAllCityData()
        {
            Browser = new GeckoWebBrowser();


            Browser.DocumentCompleted += new EventHandler<Gecko.Events.GeckoDocumentCompletedEventArgs>(Browser_DocumentCompleted_Init);

            GeckoPreferences.User["gfx.font_rendering.graphite.enabled"] = true;
            Browser.Dock = DockStyle.Fill;

            //显示用的窗体
            viewForm = new Form();
            //viewForm.Icon=
            viewForm.ClientSize = new System.Drawing.Size(550, 600);
            viewForm.StartPosition = FormStartPosition.CenterScreen;
            viewForm.Controls.Clear();
            viewForm.Controls.Add(Browser);
            viewForm.FormClosing += new FormClosingEventHandler(viewForm_FormClosing);
            if (run_type == 3)//查询
            {
                viewForm.Show();
                viewForm.MaximizeBox = true;
                viewForm.WindowState = FormWindowState.Maximized;
            }
            viewForm.Text = "MSTI";

            if (run_type == 2 || run_type == 3)
                testUrl = "https://om.qq.com/userAuth/index";//请完成企业银行账户验证
            else if (run_type == 1)
                testUrl = "https://publish.caasdata.com/";//火星


            Browser.Navigate(testUrl);
        }
        void Browser_DocumentCompleted_Init(object sender, EventArgs e)
        {


            GeckoWebBrowser br = sender as GeckoWebBrowser;
            if (br.Url.ToString() == "about:blank") { return; }


            GmyDoc = sender as GeckoWebBrowser;
            //GmyDoc.Document.Cookie.Remove(0, (GmyDoc.Document.Cookie.Count() - 1));

            #region 火星
            if (run_type == 1)
            {
                #region 退出账号

                GeckoHtmlElement extit = null;
                GeckoElementCollection extit1 = GmyDoc.Document.GetElementsByTagName("span");
                foreach (GeckoHtmlElement item in extit1)
                {
                    if (item.InnerHtml.Contains("退出"))
                    {
                        extit = item;

                    }
                }
                if (extit != null)
                {
                    //  extit.Click();

                }
                #endregion


                if (GmyDoc.Url.ToString().IndexOf("https://publish.caasdata.com/") >= 0 && isrun == ProcessStatus.初始化)
                {
                    GeckoElement userName = null;
                    GeckoHtmlElement submit = null;
                    GeckoElementCollection userNames = GmyDoc.Document.GetElementsByTagName("input");
                    foreach (GeckoHtmlElement item in userNames)
                    {
                        if (item.GetAttribute("value") == "default")
                            submit = item;

                    }
                    //myBrowser.Document.GetHtmlElementById("vcodeA").Click();
                    //     submit.Click();

                    isrun = ProcessStatus.结束页面;

                }
                if (isrun == ProcessStatus.结束页面)
                {


                    GeckoHtmlElement namevalue = null;
                    GeckoHtmlElement password = null;
                    GeckoHtmlElement submit = null;
                    GeckoElementCollection userNames = GmyDoc.Document.GetElementsByTagName("input");
                    GeckoElementCollection userNames1 = GmyDoc.Document.GetElementsByTagName("a");

                    foreach (GeckoHtmlElement item in userNames)
                    {
                        if (item.GetAttribute("placeholder") == "输入手机号")
                            namevalue = item;
                        if (item.GetAttribute("id") == "password")
                            password = item;
                    }
                    foreach (GeckoHtmlElement item in userNames1)
                    {
                        if (item.GetAttribute("id") == "ensure")
                            submit = item;

                    }
                    if (namevalue != null)
                    {
                        namevalue.SetAttribute("Value", "16619776280");
                        //   password.SetAttribute("Value", "123456");

                        //  submit.Click();
                    }

                    isrun = ProcessStatus.登录界面;
                }
                if (isrun == ProcessStatus.登录界面)
                {
                    GeckoHtmlElement namevalue = null;

                    GeckoElementCollection userNames = GmyDoc.Document.GetElementsByTagName("p");
                    foreach (GeckoHtmlElement item in userNames)
                    {
                        //选择账号
                        if (item.GetAttribute("title") == xuanzhezhanghao[0])
                            namevalue = item;
                        if (item.InnerHtml.Contains(xuanzhezhanghao[0]))
                            namevalue = item;
                    }


                    //上传视频
                    GeckoHtmlElement submit = null;
                    GeckoElement upid = GmyDoc.Document.GetElementById("uploadBtn");


                    if (namevalue != null)
                        namevalue.Click();

                    //if (upid != null)
                    //    br.Document.GetHtmlElementById("uploadBtn").Click();

                    //单独编目
                    GeckoHtmlElement dandu = null;
                    GeckoElementCollection span = GmyDoc.Document.GetElementsByTagName("span");
                    foreach (GeckoHtmlElement item in span)
                    {
                        //选择账号
                        if (item.InnerHtml == "(不同账号可以分别填写发布信息) ")
                            dandu = item;

                    }
                    if (dandu != null)
                        dandu.Click();


                    isrun = ProcessStatus.确认YES;
                }

                if (isrun == ProcessStatus.确认YES)
                {
                    NewMethod();


                }
            }
            #endregion
            #region 请完成企业银行账户验证
            if (run_type == 2 || run_type == 3)
            {
                if (GmyDoc.Url.ToString().IndexOf("https://om.qq.com/userAuth/index") >= 0 && isrun == ProcessStatus.初始化)
                {
                    GeckoHtmlElement fenlei = null;

                    GeckoElementCollection div = GmyDoc.Document.GetElementsByTagName("div");
                    foreach (GeckoHtmlElement item in div)
                    {
                        //名称
                        if (item.GetAttribute("btn-type") == "email")
                        {
                            fenlei = item;
                            //jianjie.SetAttribute("Value", "简介");
                            // jianjie.TextContent = "我的内容";
                        }
                    }
                    if (fenlei != null)
                        fenlei.Click();
                    isrun = ProcessStatus.结束页面;
                }
                if (isrun == ProcessStatus.结束页面)
                {

                    GeckoHtmlElement namevalue = null;
                    GeckoHtmlElement password = null;
                    GeckoHtmlElement submit = null;
                    GeckoElementCollection userNames = GmyDoc.Document.GetElementsByTagName("input");
                    GeckoElementCollection userNames1 = GmyDoc.Document.GetElementsByTagName("button");

                    foreach (GeckoHtmlElement item in userNames)
                    {
                        if (item.GetAttribute("placeholder") == "邮箱/手机号")
                            namevalue = item;
                        if (item.GetAttribute("name") == "password")
                            password = item;
                    }
                    foreach (GeckoHtmlElement item in userNames1)
                    {
                        if (item.GetAttribute("type") == "button" && item.InnerHtml == "登录")
                            submit = item;

                    }
                    if (namevalue != null && password != null && submit != null)
                    {
                        //namevalue.SetAttribute("Value", "tnbpsb653dq2@163.com");
                        //password.SetAttribute("Value", "aabb33");
                        namevalue.SetAttribute("Value", Item_RResult.zhanghao);
                        password.SetAttribute("Value", Item_RResult.mima);
                        submit.Click();
                    }

                    isrun = ProcessStatus.登录界面;

                }
                if (GmyDoc.Url.ToString().IndexOf("https://om.qq.com/user/comBlackProIntercept") >= 0 && isrun == ProcessStatus.登录界面)
                {
                    if (run_type == 2)
                        readqiyemingcheng();
                    if (run_type == 3)
                        erici_fuzhi();

                    isrun = ProcessStatus.确认YES;
                }

            }



            #endregion
        }
        private void readqiyemingcheng()
        {
            if (isrun == ProcessStatus.确认YES)
            {
                //名称
                GeckoHtmlElement mingcheng = null;
                GeckoHtmlElement yinhangzhanghao = null;
                GeckoElementCollection input = GmyDoc.Document.GetElementsByTagName("input");
                foreach (GeckoHtmlElement item in input)
                {
                    //名称
                    if (item.GetAttribute("placeholder") == "默认带入注册信息中的企业名称（不可修改）")
                        mingcheng = item;
                    if (item.GetAttribute("placeholder") == "请输入企业开户的银行账号")
                        yinhangzhanghao = item;
                }
                if (mingcheng != null)
                {
                    string qiyemingcheng = mingcheng.GetAttribute("value");
                    Item_RResult.xuyaoyanzhengdegongsiming = qiyemingcheng;
                    RResult[yinhangrun_index].xuyaoyanzhengdegongsiming = qiyemingcheng;

                    #region 退出账号

                    exitwb();
                    #endregion
                    isrun = ProcessStatus.第一页面;
                    isOneFinished = true;
                }
            }
        }

        private void exitwb()
        {
            GeckoHtmlElement extit = null;
            GeckoElementCollection extit1 = GmyDoc.Document.GetElementsByTagName("a");
            foreach (GeckoHtmlElement item in extit1)
            {
                if (item.GetAttribute("title") == "退出")
                {
                    extit = item;

                }
            }
            if (extit != null)
            {
                extit.Click();

            }
        }

        private void erici_fuzhi()
        {
            if (isrun == ProcessStatus.确认YES)
            {
                //名称
                GeckoHtmlElement mingcheng = null;
                GeckoHtmlElement yinhangzhanghao = null;
                GeckoElementCollection input = GmyDoc.Document.GetElementsByTagName("input");
                foreach (GeckoHtmlElement item in input)
                {
                    //名称
                    if (item.GetAttribute("placeholder") == "默认带入注册信息中的企业名称（不可修改）")
                        mingcheng = item;
                    if (item.GetAttribute("placeholder") == "请输入企业开户的银行账号")
                        yinhangzhanghao = item;
                }
                if (mingcheng != null)
                {
                    string qiyemingcheng = mingcheng.GetAttribute("value");

                }
                if (yinhangzhanghao != null)
                {
                    yinhangzhanghao.SetAttribute("Value", Item_RResult.yinhangka);

                }

                GeckoHtmlElement kaihuyinhang = null;
                GeckoHtmlElement shengfen = null;
                GeckoHtmlElement shiqu = null;
                GeckoHtmlElement zhihang = null;
                GeckoElementCollection span = GmyDoc.Document.GetElementsByTagName("span");
                GeckoElementCollection li = GmyDoc.Document.GetElementsByTagName("li");
                GeckoElementCollection div = GmyDoc.Document.GetElementsByTagName("div");
                foreach (GeckoHtmlElement item in span)
                {
                    //选择账号
                    if (item.InnerHtml == "请选择银行")
                        kaihuyinhang = item;
                    if (item.InnerHtml == "省份/直辖市")
                        shengfen = item;
                    if (item.InnerHtml == "市/区")
                        shiqu = item;
                    if (item.InnerHtml == "请选择支行")
                        zhihang = item;
                }
                kaihuyinhang = null;
                foreach (GeckoHtmlElement item in div)
                {
                    if (item.GetAttribute("class") == "chosen-container chosen-container-single")
                        kaihuyinhang = item;
                }

                if (kaihuyinhang != null && isrun1 == EapprovalProcessStatus.初始化)
                {
                    //kaihuyinhang.InnerHtml = Item_RResult.kaihuhang;
                    kaihuyinhang.Click();
                    kaihuyinhang.Click();
                    isrun1 = EapprovalProcessStatus.Current_Task;

                }
                kaihuyinhang = null;

                foreach (GeckoHtmlElement item in li)
                {
                    //选择账号
                    if (item.InnerHtml == "中国银行")
                        kaihuyinhang = item;
                }
                if (kaihuyinhang != null)
                {
                    kaihuyinhang.SetAttribute("class", "active-result result-selected");
                }

                if (kaihuyinhang != null)
                {
                    shengfen.InnerHtml = Item_RResult.kaihusheng;// "黑龙江省";
                }
                if (kaihuyinhang != null)
                {
                    shiqu.InnerHtml = Item_RResult.kaihushi; //"佳木斯市";
                }
                if (kaihuyinhang != null)
                {
                    zhihang.InnerHtml = Item_RResult.kaihuzhihang; //"中国农业银行股份有限公司黑龙江省佳木斯市中山支行";

                }
                //提交

                GeckoHtmlElement submit = null;
                GeckoElementCollection userNames1 = GmyDoc.Document.GetElementsByTagName("button");
                foreach (GeckoHtmlElement item in userNames1)
                {
                    if (item.GetAttribute("type") == "button" && item.InnerHtml == "提交")
                        submit = item;
                }
                if (submit != null && yinhangzhanghao != null)
                {
                    //  submit.Click();
                    vi();

                    isrun = ProcessStatus.第一页面;
                }

            }
            if (GmyDoc.Url.ToString().IndexOf("https://om.qq.com/user/comBlackProIntercept") >= 0 && isrun == ProcessStatus.第一页面)
            {

                GeckoHtmlElement namevalue = null;
                GeckoHtmlElement vi_meipaotong = null;

                GeckoElementCollection userNames = GmyDoc.Document.GetElementsByTagName("p");
                foreach (GeckoHtmlElement item in userNames)
                {
                    if (item.GetAttribute("class") == "info")
                        namevalue = item;
                    if (item.GetAttribute("class") == "help-block error" && item.InnerHtml == "请选择银行")
                        vi_meipaotong = item;
                }
                if (vi_meipaotong != null && isOneFinished_VI == true)
                {
                    //   vi();
                }
                if (namevalue != null)
                {
                    string jiegou = namevalue.InnerHtml;
                    Item_RResult.tijiao_jieguo = jiegou;
                    #region 退出账号

                    exitwb();
                    #endregion
                    isrun = ProcessStatus.第一页面;
                    isOneFinished = true;
                }
            }

        }

        private bool vi()
        {
            //遍历所有查找到的进程
            isOneFinished_VI = false;
            StopTime = DateTime.Now;
            bool istrue = false;
            while (!isOneFinished_VI)
            {
                Process[] pro = Process.GetProcesses();//获取已开启的所有进程

                bool iscontains = false;
                for (int ii = 0; ii < pro.Length; ii++)
                {
                    if (pro[ii].ProcessName.ToString().Contains("ST"))
                    {
                        iscontains = true;

                        DateTime rq2 = DateTime.Now;  //结束时间
                        TimeSpan ts = rq2 - StopTime;
                        int timeTotal = ts.Minutes;
                        if (timeTotal >= 2)
                        {
                            //bgWorker1.ReportProgress(0, "系统错误，正在执行推出，请稍后自行检查数据源错误或运行环境问题！");
                            pro[ii].Kill();//结束进程
                            isOneFinished_VI = true;
                            //Application.Exit();
                        }
                    }
                }
                if (iscontains == false)
                    isOneFinished_VI = true;
            }
            Thread.Sleep(1000);
            //  if (i == 0)
            user_winauto("");
            return isOneFinished_VI;

        }
        private static void user_winauto(string fajianren)
        {

            {
                string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ""), "");
                System.Diagnostics.Process.Start("ST.exe", ZFCEPath);
            }

        }
        //private void Notice_LoadCompleted(object sender, System.Windows.Navigation.NavigationEventArgs e)
        //{
        //    string script = "document.body.style.overflow ='hidden'";
        //    WebBrowser wb = (WebBrowser)sender;
        //    wb.InvokeScript("execScript", new Object[] { script, "JavaScript" });
        //}
        private void NewMethod()
        {

            //名称
            GeckoHtmlElement mingcheng = null;
            GeckoElementCollection input = GmyDoc.Document.GetElementsByTagName("input");
            foreach (GeckoHtmlElement item in input)
            {
                //名称
                if (item.GetAttribute("placeholder") == "请输入节目名称")
                    mingcheng = item;

            }
            if (mingcheng != null)
                mingcheng.SetAttribute("Value", "名称123");


            //分类
            GeckoHtmlElement fenlei = null;
            GeckoElementCollection select = GmyDoc.Document.GetElementsByTagName("select");
            foreach (GeckoHtmlElement item in select)
            {
                //名称
                if (item.GetAttribute("class") == "modal_form_group_sel")
                    fenlei = item;

            }
            GeckoElementCollection option = GmyDoc.Document.GetElementsByTagName("option");
            foreach (GeckoHtmlElement item in option)
            {
                //名称
                if (item.GetAttribute("value") == "体育趣闻")
                    fenlei = item;

            }

            if (fenlei != null)
            {
                //fenlei.SetAttribute("selecti", "1");
                //state - select.SetAttribute("value", "值");
                //fenlei.SetAttribute("Value", "综合体育");
                //fenlei.SetAttribute("selectedIndex", "2");
                //fenlei.SetAttribute("value", "2");
                fenlei.InnerHtml = "综合体育";
                fenlei.TextContent = "篮球";



                isrun = ProcessStatus.第二页面;
            }

            ////标签
            //GeckoElement biaoqian = GmyDoc.Document.GetElementById("input_center2");
            //if (biaoqian != null)
            //{
            //    biaoqian.SetAttribute("text", "标签");


            //}
            ////简介}

            //GeckoHtmlElement jianjie = null;
            //GeckoHtmlElement jianjiediv = null;
            //GeckoElementCollection textarea = GmyDoc.Document.GetElementsByTagName("textarea");
            //GeckoElementCollection div = GmyDoc.Document.GetElementsByTagName("div");
            //foreach (GeckoHtmlElement item in textarea)
            //{
            //    //名称
            //    if (item.GetAttribute("placeholder") == "请输入简介")
            //    {
            //        jianjie = item;
            //        //jianjie.SetAttribute("Value", "简介");
            //        // jianjie.TextContent = "我的内容";
            //    }
            //}

            //if (jianjie != null)
            //{
            //    //IHTMLWindow2 win = (IHTMLWindow2)GmyDoc.Document.Body;
            //    //string scriptLine = @"document.getElementById('login').focus()";
            //    //win.execScript(scriptLine, "Javascript");


            //    jianjie.SetAttribute("Text", "简介");
            //    // textarea[0].TextContent = "nihaoodd";

            //    // GmyDoc.Document.ChildNodes[0].OwnerDocument.GetElementById("tinymce").TextContent = "我的内容";

            //    //jianjiediv.SetAttribute("Value", "简介");
            //    //jianjie.TextContent = "我的内容";
            //    // isrun = ProcessStatus.第二页面;
            //}
            ////确认发布
            //GeckoHtmlElement fabu = null;
            //GeckoElementCollection button = GmyDoc.Document.GetElementsByTagName("button");
            //foreach (GeckoHtmlElement item in button)
            //{
            //    //名称
            //    if (item.InnerHtml == "确定发行")
            //        fabu = item;

            //}
            ////if (fabu != null)
            ////    fabu.Click();
        }


        static void InitializeExtras()
        {
            //Initialize the cookie manager
            CookieMan = Xpcom.GetService<nsICookieManager>("@mozilla.org/cookiemanager;1");
            CookieMan = Xpcom.QueryInterface<nsICookieManager>(CookieMan);
        }
        public static void Initialize(string binDirectory)
        {
            InitializeExtras();
        }
        public void DeleteAllCookies()
        {
            CookieMan.RemoveAll();
        }
        #endregion

        private void steap_tocover(object sender, EventArgs e)
        {
            SaveAs();
            //bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_015);
            // MessageBox.Show("Mytest");
        }

        #region AＰI

        //第一步保存
        //private void SaveAs(object sender, EventArgs e)
        private void SaveAs()
        {
            DateTime rq2 = DateTime.Now;  //结束时间               

            int a = rq2.Second - strFileName.Second;
            if (a >= 10 || rq2.Second < strFileName.Second)
            {
                isReadyForSearch = true;
                aTimer.Stop();
                //  isOneFinished = true;
                return;

            }

            List<IntPtr> arrHwnd_Sap_before = getSAPWindow();
            FileInfo Log4NetFile = new FileInfo("./Log4Net.Config");
            log4net.Config.XmlConfigurator.Configure(Log4NetFile);

            // log4net.ILog objLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");

            bool blFresh = false;
            RUNING = true;
            try
            {
                //得到File Download窗体对象
                IntPtr h1 = FindWindow("#32770", "File Download - Security Warning");
                IntPtr h2 = FindWindow("#32770", "文件下载 - 安全警告");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "File Download");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "文件下载");

                if (h2.ToInt32() > 0 || h1.ToInt32() > 0)
                {
                    //得到Save按钮对象
                    IntPtr duixiang;

                    if (h1.ToInt32() > 0)
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h1, 0, "Button", "&Save");
                    }
                    else
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h2, 0, "BUTTON", "&Save");
                    }
                    //objLogger.Fatal("BUTTON 保存(&S):" + duixiang.ToString());
                    //如果得到点击Save按钮
                    if (duixiang.ToInt32() > 0)
                    {
                        SendMessage(duixiang, 0xF5, 0, 0);
                        SendMessage(duixiang, 0xF5, 0, 0);
                        WinAPIuser32.SendMessage(duixiang, WM_LBUTTONUP, IntPtr.Zero, null);
                        WinAPIuser32.SendMessage(duixiang, WM_LBUTTONUP, IntPtr.Zero, null);
                        //设定File Download窗体是否点击变量为已点击
                        blFresh = true;
                    }
                    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "&Save");
                    if (hwnd2.ToInt32() <= 0)
                        hwnd2 = WinAPIuser32.FindWindow("#32770", "Save");
                    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());
                }

                RUNING = false;
                if (true)
                {

                    //得到Save As窗体对象
                    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "Save As");
                    if (hwnd2.ToInt32() <= 0)
                        hwnd2 = WinAPIuser32.FindWindow("#32770", "名前を付けて保存");

                    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());

                    if (hwnd.ToInt32() > 0 || hwnd2.ToInt32() > 0)
                    {
                        //得到其下的一系列子窗体对象

                        IntPtr htextbox;
                        IntPtr htextbox1;

                        if (hwnd.ToInt32() > 0)
                        {
                            htextbox = WinAPIuser32.FindWindowEx(hwnd, 0, "ComboBoxEx32", null);
                            htextbox1 = WinAPIuser32.FindWindowEx(hwnd, 0, "DUIViewWndClassName", null);
                        }
                        else
                        {
                            htextbox = WinAPIuser32.FindWindowEx(hwnd2, 0, "ComboBoxEx32", null);
                            htextbox1 = WinAPIuser32.FindWindowEx(hwnd2, 0, "DUIViewWndClassName", null);
                        }

                        //objLogger.Fatal("ComboBoxEx32" + htextbox.ToString() + htextbox1.ToString());

                        if (htextbox.ToInt32() > 0)
                        {
                            IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox, 0, "ComboBox", null);
                            if (htextbox4.ToInt32() > 0)
                            {
                                Thread.Sleep(1000);
                                //得到子窗体中输入保存路径文本框

                                IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                                string strPath = FileDownURL + FileName;
                                WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                                //得到Save按钮对象
                                IntPtr hbutton;
                                if (hwnd.ToInt32() > 0)
                                {
                                    hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                                }
                                else
                                {
                                    hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                                    if (hbutton.ToInt32() <= 0)
                                        hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                                }

                                //如果得到点击Save按钮
                                if (hbutton.ToInt32() > 0)
                                {
                                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                                    //停止每5秒点击Save As窗体Timer控件
                                    WebSiteStatus = true;
                                }
                            }
                        }
                        else if (htextbox1.ToInt32() > 0)
                        {
                            IntPtr htextbox2 = WinAPIuser32.FindWindowEx(htextbox1, 0, "DirectUIHWND", null);
                            if (htextbox2.ToInt32() > 0)
                            {
                                IntPtr htextbox3 = WinAPIuser32.FindWindowEx(htextbox2, 0, "FloatNotifySink", null);
                                if (htextbox3.ToInt32() > 0)
                                {
                                    IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox3, 0, "ComboBox", null);
                                    if (htextbox4.ToInt32() > 0)
                                    {
                                        Thread.Sleep(1000);
                                        //得到子窗体中输入保存路径文本框
                                        // FileName = @"C:";


                                        ///////////////////////////////////////////////////////////路径保存地址
                                        FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources\\PDF\\");
                                        string strPath = FileDownURL + FileName + publicPDFName + ".pdf";
                                        if (File.Exists(strPath))
                                        {
                                            File.Delete(strPath);

                                        }
                                        IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                                        WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                                        //得到Save按钮对象
                                        IntPtr hbutton;
                                        if (hwnd.ToInt32() > 0)
                                        {
                                            hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                                        }
                                        else
                                        {
                                            hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                                            if (hbutton.ToInt32() <= 0)
                                                hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                                        }
                                        //如果得到点击Save按钮
                                        if (hbutton.ToInt32() > 0)
                                        {
                                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                                            //停止每5秒点击Save As窗体Timer控件
                                            WebSiteStatus = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                RUNING = false;
                throw;
            }
            //    isReadyForSearch = false;
            //    WebSiteStatus = false;
            //    aTimer.Stop();


        }

        //第二部保存
        private void OnTimedEvent(object sender, EventArgs e)
        {
            Thread.Sleep(3000);
            //IntPtr hwnd = FindWindow("#32770", "另存为");
            IntPtr hwnd = FindWindow("#32770", "Save As");
            if (int.Parse(hwnd.ToString()) == 0)
            {
                hwnd = FindWindow("#32770", "Save");
            }
            if (int.Parse(hwnd.ToString()) > 0)
            {
                IntPtr hbutton = GetDlgItem(hwnd, 1);
                if (int.Parse(hbutton.ToString()) > 0)
                {
                    SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                    SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                    aTimer.Stop();
                }
            }

        }


        public void SaveAPIButton()
        {
            bool RUNING = false;

            //  List<IntPtr> arrHwnd_Sap_before = getSAPWindow();
            FileInfo Log4NetFile = new FileInfo("./Log4Net.Config");
            log4net.Config.XmlConfigurator.Configure(Log4NetFile);

            // log4net.ILog objLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");

            bool blFresh = false;
            RUNING = true;
            try
            {
                //得到File Download窗体对象
                IntPtr h1 = FindWindow("#32770", "File Download - Security Warning");
                IntPtr h2 = FindWindow("#32770", "文件下载 - 安全警告");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "ファイルのダウンロード");
                if (h2.ToInt32() <= 0)
                    h2 = FindWindow("#32770", "文件下载");

                if (h2.ToInt32() > 0 || h1.ToInt32() > 0)
                {
                    //得到Save按钮对象
                    IntPtr duixiang;

                    if (h1.ToInt32() > 0)
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h1, 0, "Button", "&Save");
                    }
                    else
                    {
                        duixiang = WinAPIuser32.FindWindowEx(h2, 0, "BUTTON", "保存(&S)");
                    }
                    //objLogger.Fatal("BUTTON 保存(&S):" + duixiang.ToString());
                    //如果得到点击Save按钮
                    if (duixiang.ToInt32() > 0)
                    {
                        SendMessage(duixiang, 0xF5, 0, 0);
                        SendMessage(duixiang, 0xF5, 0, 0);
                        WinAPIuser32.SendMessage(duixiang, WM_LBUTTONUP, IntPtr.Zero, null);

                        //设定File Download窗体是否点击变量为已点击
                        blFresh = true;
                    }
                    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "Save As");
                    if (hwnd2.ToInt32() <= 0)
                        hwnd2 = WinAPIuser32.FindWindow("#32770", "名前を付けて保存");

                    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());
                }
                RUNING = false;

                if (true)
                {
                    //得到Save As窗体对象
                    IntPtr hwnd = WinAPIuser32.FindWindow("#32770", "另存为");
                    IntPtr hwnd2 = WinAPIuser32.FindWindow("#32770", "Save As");
                    if (hwnd2.ToInt32() <= 0)
                        hwnd2 = WinAPIuser32.FindWindow("#32770", "名前を付けて保存");

                    //objLogger.Fatal("另存为" + hwnd.ToString() + hwnd2.ToString());

                    if (hwnd.ToInt32() > 0 || hwnd2.ToInt32() > 0)
                    {
                        //得到其下的一系列子窗体对象

                        IntPtr htextbox;
                        IntPtr htextbox1;

                        if (hwnd.ToInt32() > 0)
                        {
                            htextbox = WinAPIuser32.FindWindowEx(hwnd, 0, "ComboBoxEx32", null);
                            htextbox1 = WinAPIuser32.FindWindowEx(hwnd, 0, "DUIViewWndClassName", null);
                        }
                        else
                        {
                            htextbox = WinAPIuser32.FindWindowEx(hwnd2, 0, "ComboBoxEx32", null);
                            htextbox1 = WinAPIuser32.FindWindowEx(hwnd2, 0, "DUIViewWndClassName", null);
                        }

                        //objLogger.Fatal("ComboBoxEx32" + htextbox.ToString() + htextbox1.ToString());

                        if (htextbox.ToInt32() > 0)
                        {
                            IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox, 0, "ComboBox", null);
                            if (htextbox4.ToInt32() > 0)
                            {
                                Thread.Sleep(1000);
                                //得到子窗体中输入保存路径文本框

                                IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                                string strPath = FileDownURL + FileName;
                                WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                                //得到Save按钮对象
                                IntPtr hbutton;
                                if (hwnd.ToInt32() > 0)
                                {
                                    hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                                }
                                else
                                {
                                    hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                                    if (hbutton.ToInt32() <= 0)
                                        hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                                }

                                //如果得到点击Save按钮
                                if (hbutton.ToInt32() > 0)
                                {
                                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                                    WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                                    //停止每5秒点击Save As窗体Timer控件
                                    WebSiteStatus = true;
                                }
                            }
                        }
                        else if (htextbox1.ToInt32() > 0)
                        {
                            IntPtr htextbox2 = WinAPIuser32.FindWindowEx(htextbox1, 0, "DirectUIHWND", null);
                            if (htextbox2.ToInt32() > 0)
                            {
                                IntPtr htextbox3 = WinAPIuser32.FindWindowEx(htextbox2, 0, "FloatNotifySink", null);
                                if (htextbox3.ToInt32() > 0)
                                {
                                    IntPtr htextbox4 = WinAPIuser32.FindWindowEx(htextbox3, 0, "ComboBox", null);
                                    if (htextbox4.ToInt32() > 0)
                                    {
                                        Thread.Sleep(1000);
                                        //得到子窗体中输入保存路径文本框
                                        FileName = @"C:";
                                        IntPtr htextbox5 = WinAPIuser32.FindWindowEx(htextbox4, 0, "Edit", null);
                                        string strPath = FileDownURL + FileName;
                                        WinAPIuser32.SendMessage(htextbox5, WM_SETTEXT, IntPtr.Zero, strPath);

                                        //得到Save按钮对象
                                        IntPtr hbutton;
                                        if (hwnd.ToInt32() > 0)
                                        {
                                            hbutton = WinAPIuser32.FindWindowEx(hwnd, 0, "BUTTON", "保存(&S)");
                                        }
                                        else
                                        {
                                            hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "&Save");
                                            if (hbutton.ToInt32() <= 0)
                                                hbutton = WinAPIuser32.FindWindowEx(hwnd2, 0, "BUTTON", "保存(&S)");
                                        }
                                        //如果得到点击Save按钮
                                        if (hbutton.ToInt32() > 0)
                                        {
                                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONDOWN, IntPtr.Zero, null);
                                            WinAPIuser32.SendMessage(hbutton, WM_LBUTTONUP, IntPtr.Zero, null);
                                            //停止每5秒点击Save As窗体Timer控件
                                            WebSiteStatus = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RUNING = false;
                throw;
            }
        }

        #region Windows API Function

        private List<IntPtr> getSAPWindow()
        {
            IntPtr hwnd_sap = IntPtr.Zero;
            List<IntPtr> arrHwnd = new List<IntPtr>();

            while (true)
            {
                hwnd_sap = FindWindowEx(IntPtr.Zero, hwnd_sap, "SAP_FRONTEND_SESSION", null);
                if (hwnd_sap != IntPtr.Zero)
                {
                    arrHwnd.Add(hwnd_sap);
                }
                else
                {

                    break;
                }
            }

            return arrHwnd;
        }

        private void monitorSAP()
        {
            if (hwnd_main == IntPtr.Zero)
            {
                //MessageBox.Show("找不到SAP主窗口！");
                return;
            }

            if (IsWindowVisible(hwnd_main))
            {
                if (ScreenStatus == 0)
                {
                    hwnd_ReportTree1 = findReportTree1();
                    //MessageBox.Show(hwnd_ReportTree1.ToInt32().ToString());
                    selectReport(intCnt);

                    IntPtr btnExecute = findExecuteButton();
                    SendMessage(btnExecute, BM_CLICK, 0, 0);
                    ScreenStatus = 1;
                }
                //clickYes(btnExecute);
            }
        }

        private IntPtr findExecuteButton()
        {
            IntPtr children = FindWindowEx(hwnd_main, IntPtr.Zero, null, "");
            while (children != IntPtr.Zero)
            {
                children = FindWindowEx(hwnd_main, children, null, "");
                int nRet;
                StringBuilder ClassName = new StringBuilder(100);
                //Get the window class name
                nRet = GetClassName(children, ClassName, ClassName.Capacity);
                Regex r = new Regex("Afx:[(a-z)|(A-Z)|(0-9)]{8}:8:[0-9]{8}:00000000:00000000");
                if (nRet != 0 && r.Match(ClassName.ToString()).Success)
                {
                    IntPtr hwnd_level2 = FindWindowEx(children, IntPtr.Zero, "Button", null);
                    if (hwnd_level2 != IntPtr.Zero)
                    {
                        return hwnd_level2;
                    }
                }
            }
            return IntPtr.Zero;
        }

        private void clickYes(IntPtr hwnd_Control)
        {
            IntPtr hwnd_Button = FindWindowEx(hwnd_Control, new IntPtr(0), "Button", null);
            SendMessage(hwnd_Button, BM_CLICK, 0, 0);
        }

        private void selectReport(int intCount)
        {
            sendPageUp();
            sendTab(intCount);
        }

        private void sendPageUp()
        {
            for (int i = 0; i < 5; i++)
            {
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_CONTROL, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_PRIOR, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_PRIOR, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_CONTROL, 0);
            }
        }

        private void sendTab(int intCount)
        {
            //Tab
            for (int i = 0; i < intCount; i++)
            {
                SendMessage(hwnd_ReportTree1, WM_KEYDOWN, VK_TAB, 0);
                SendMessage(hwnd_ReportTree1, WM_KEYUP, VK_TAB, 0);
            }
        }

        private IntPtr findReportTree1()
        {
            IntPtr hwnd_level1, hwnd_level2, hwnd_level3;
            hwnd_level1 = FindWindowEx(hwnd_main, IntPtr.Zero, "Docking Container Class", null);
            //MessageBox.Show("Docking Container Class:" + hwnd_level1.ToInt32().ToString());
            if (hwnd_level1 != IntPtr.Zero)
            {
                hwnd_level2 = FindWindowEx(hwnd_level1, IntPtr.Zero, "Shell Window Class", "Control  Container");
                //MessageBox.Show("Control  Container:" + hwnd_level2.ToInt32().ToString());
                if (hwnd_level2 != IntPtr.Zero)
                {
                    hwnd_level3 = FindWindowEx(hwnd_level2, IntPtr.Zero, "AfxOleControl80", null);
                    if (hwnd_level3 == IntPtr.Zero)
                        hwnd_level3 = FindWindowEx(hwnd_level2, IntPtr.Zero, "AfxOleControl90", null);
                    //MessageBox.Show("AfxOleControl80:" + hwnd_level3.ToInt32().ToString());
                    if (hwnd_level3 != IntPtr.Zero)
                    {
                        //MessageBox.Show("SAPTreeList:" + FindWindowEx(hwnd_level3, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist").ToInt32().ToString());
                        return FindWindowEx(hwnd_level3, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist");
                        //return FindWindowEx(hwnd_level2, IntPtr.Zero, "SAPTreeList", "SAP's Advanced Treelist");
                    }
                }
            }
            return IntPtr.Zero;
        }

        private IntPtr findControlWindow()
        {
            string strCaption = "Execute Project Report: Initial Screen";

            IntPtr hwnd_Child = IntPtr.Zero;
            while (true)
            {
                hwnd_Child = FindWindowEx(IntPtr.Zero, hwnd_Child, "#32770", strCaption);
                if (GetParent(hwnd_Child) == hwnd_main || hwnd_Child == IntPtr.Zero)
                {
                    break;
                }
            }
            return hwnd_Child;
        }
        #endregion
        #endregion


        #region 读取Ex菜单l
        public List<clsDatabaseinfo> ReadfindngFile(string casetype)
        {

            List<clsDatabaseinfo> Result = new List<clsDatabaseinfo>();

            //string path = AppDomain.CurrentDomain.BaseDirectory + "Resources\\qiyeziliao.xlsx";
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(casetype, Type.Missing, true, Type.Missing,
                "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range rng;
            //rng = WS.get_Range(WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]);
            rng = WS.Range[WS.Cells[2, 1], WS.Cells[WS.UsedRange.Rows.Count, 30]];
            int rowCount = WS.UsedRange.Rows.Count - 1;
            object[,] o = new object[1, 1];
            o = (object[,])rng.Value2;
            clsCommHelp.CloseExcel(excelApp, analyWK);

            for (int i = 1; i <= rowCount; i++)
            {
                bgWorker1.ReportProgress(0, "正在导入   :  " + i.ToString() + "/" + rowCount.ToString());
                clsDatabaseinfo temp = new clsDatabaseinfo();

                #region 基础信息

                temp.zhanghao = "";
                if (o[i, 1] != null)
                    temp.zhanghao = o[i, 1].ToString().Trim();


                temp.mima = "";
                if (o[i, 2] != null)
                    temp.mima = o[i, 2].ToString().Trim();

                temp.zhuangtai = "";
                if (o[i, 3] != null)
                    temp.zhuangtai = o[i, 3].ToString().Trim();

                temp.xuyaoyanzhengdegongsiming = "";
                if (o[i, 4] != null)
                    temp.xuyaoyanzhengdegongsiming = o[i, 4].ToString().Trim();
                if (temp.zhanghao == "" || temp.zhanghao == null)
                    continue;

                temp.yonghuming = "";
                if (o[i, 5] != null)
                    temp.yonghuming = o[i, 5].ToString().Trim();
                temp.yinhangka = "";
                if (o[i, 6] != null)
                    temp.yinhangka = o[i, 6].ToString().Trim(); //clsCommHelp.objToDateTime(o[i, 6]);

                temp.kaihuhang = "";
                if (o[i, 7] != null)
                    temp.kaihuhang = o[i, 7].ToString().Trim();
                temp.kaihusheng = "";
                if (o[i, 8] != null)
                    temp.kaihusheng = o[i, 8].ToString().Trim();

                temp.kaihushi = "";
                if (o[i, 9] != null)
                    temp.kaihushi = o[i, 9].ToString().Trim();

                temp.kaihuzhihang = "";
                if (o[i, 10] != null)
                    temp.kaihuzhihang = o[i, 10].ToString().Trim();

                temp.tijiaozhuangtai = "";
                if (o[i, 11] != null)
                    temp.tijiaozhuangtai = o[i, 11].ToString().Trim();
                temp.tijiaoanjian = "";
                if (o[i, 12] != null)
                    temp.tijiaoanjian = o[i, 12].ToString().Trim();

                temp.yanzhengjine = "";
                if (o[i, 13] != null)
                    temp.yanzhengjine = o[i, 13].ToString().Trim();

                //temp.FSSCSHENHEJIEGUO = "";
                //if (o[i, 14] != null)
                //    temp.FSSCSHENHEJIEGUO = o[i, 14].ToString().Trim();


                //temp.SHENHERIQI = "";
                //if (o[i, 15] != null)
                //    temp.SHENHERIQI = o[i, 15].ToString().Trim();

                //temp.FSSCCHULIYIJIAN = "";
                //if (o[i, 16] != null)
                //    temp.FSSCCHULIYIJIAN = o[i, 16].ToString().Trim();

                //temp.ISSUE = "";
                //if (o[i, 17] != null)
                //    temp.ISSUE = o[i, 17].ToString().Trim();

                //temp.STATUS = "";
                //if (o[i, 18] != null)
                //    temp.STATUS = o[i, 18].ToString().Trim();

                //temp.VRNUMBER = "";
                //if (o[i, 19] != null)
                //    temp.VRNUMBER = o[i, 19].ToString().Trim();

                //temp.SAPDOCUMENT = "";
                //if (o[i, 20] != null)
                //    temp.SAPDOCUMENT = o[i, 20].ToString().Trim();

                //temp.PAIDDATE = "";
                //if (o[i, 21] != null)
                //    temp.PAIDDATE = clsCommHelp.objToDateTime1(o[i, 21]).Replace("/", "");

                //temp.BUCHONGZHILIAO_BARCODE = "";
                //if (o[i, 22] != null)
                //    temp.BUCHONGZHILIAO_BARCODE = o[i, 22].ToString().Trim();

                //temp.MPR_FUKUANSHENQINGHAO = "";
                //if (o[i, 22] != null)
                //    temp.MPR_FUKUANSHENQINGHAO = o[i, 22].ToString().Trim();


                #endregion

                Result.Add(temp);
            }
            return Result;

        }
        #endregion
    }
}
