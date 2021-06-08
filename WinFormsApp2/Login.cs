using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
namespace WinFormsApp2
{
    public partial class Login : Form
    {
        WebBrowser webBrowser = new WebBrowser();
        public string loging = "";
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetGetCookieEx(string pchURL, string pchCookieName, StringBuilder pchCookieData, ref System.UInt32 pcchCookieData, int dwFlags, IntPtr lpReserved);

        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int InternetSetCookieEx(string lpszURL, string lpszCookieName, string lpszCookieData, int dwFlags, IntPtr dwReserved);


        public Login()
        {
            InitializeComponent();
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void Login_Load(object sender, EventArgs e)
        {

            
            //webBrowser.LocationChanged;
            //webBrowser.LocationChanged += new EventHandler(Change);
            
            webBrowser.Dock = DockStyle.Fill;
            webBrowser.Width = this.Width;
            webBrowser.Height = this.Height;
            //webBrowser.Navigate(new Uri("https://www.google.com"));
            //webBrowser.Url = new Uri("https://www.baidu.com");
            webBrowser.Url = new Uri(loging);
            webBrowser.ScrollBarsEnabled = true;
            webBrowser.Visible = true;
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.StatusTextChanged += new EventHandler(Change);
            this.Controls.Add(webBrowser);
            
            //webView21.Site.Container = loging;
            
            //webView21.Site.Name = loging;

        }
        private static string GetCookies(string url)
        {
            uint datasize = 256;
            StringBuilder cookieData = new StringBuilder((int)datasize);
            if (!InternetGetCookieEx(url, null, cookieData, ref datasize, 0x2000, IntPtr.Zero))
            {
                if (datasize < 0)
                    return null;


                cookieData = new StringBuilder((int)datasize);
                if (!InternetGetCookieEx(url, null, cookieData, ref datasize, 0x00002000, IntPtr.Zero))
                    return null;
            }
            return cookieData.ToString();
        }
        public void Change(object sender, EventArgs e) {
            
            try
            {
                if (webBrowser.Document == null)
                {

                }
                else
                {

                    //Debug.WriteLine(webBrowser.Document.Domain);

                    //string Cookie = webBrowser.Document.Cookie.ToString();
                    string Cookie = GetCookies(webBrowser.Document.Url.AbsoluteUri);
                    Debug.WriteLine(Cookie);

                    //Debug.WriteLine(Cookie);
                    if (Cookie.IndexOf("rtFa") != -1) {
                        //ConfigAppSettings.GetValue("Cookie");
                        ConfigAppSettings.SetValue("Cookie", Cookie);
                        
                        this.Close();
                        
                    }
                    
                    Debug.WriteLine(Cookie.IndexOf("rtFa"));
                }

                
            }
            catch(System.NullReferenceException) { 
            }
            

            //Debug.WriteLine(webBrowser.Url);
        }
    }
}
