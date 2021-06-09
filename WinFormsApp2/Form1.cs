using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.VisualBasic;

namespace WinFormsApp2
{
   // public static string Cookie = "";
    
    public partial class Form1 : Form
    {
    
        public string cookie = "";
        public string mainsite = "%27%2Fsites%2Fbox%2FShared%20Documents%27";
        public JObject GlobalJs;
        public string mulu = "";
        public string root = "";
        public string mm = "https://boxvideo.sharepoint.cn/sites/box";
        public void reload() {

            cookie = ConfigAppSettings.GetValue("Cookie");
        }

        public Form1()
           
        {
            InitializeComponent();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

            

        }
        private string GetJson(string a1, string a2) {
            
            
            string url = "https://boxvideo.sharepoint.cn/";
            url = mm+"/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=" + a1 + "&RootFolder=" + a2 + "&TryNewExperienceSingle=TRUE";
            string post = "{\"parameters\":{\"__metadata\":{\"type\":\"SP.RenderListDataParameters\"},\"RenderOptions\":1446151,\"AllowMultipleValueFilterForTaxonomyFields\":true,\"ViewXml\":\"<RowLimit Paged=\\\"TRUE\\\">99999</RowLimit>\",\"AddRequiredFields\":true}}";
            HttpWebRequest request = null;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            request = WebRequest.Create(url) as HttpWebRequest;
            //HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);


            request.Method = "POST";
            request.KeepAlive = false;
            request.ContentType = "application/json;odata=verbose";

            request.Headers["Cookie"] = cookie;

            //request.ProtocolVersion = HttpVersion.Version11;

            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(post);
            request.ContentLength = bytes.Length;
            Stream stream = request.GetRequestStream();
            stream.Write(bytes, 0, bytes.Length);//写入参数
            stream.Close();
            string header = request.Accept;
            System.Diagnostics.Debug.WriteLine(header);
            try
            {
                //HttpWebResponse response = await request.GetResponseAsync() as HttpWebResponse;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//响应对象
                string Ret = "";
                StreamReader st = new StreamReader(response.GetResponseStream(), Encoding.ASCII);
                Ret = st.ReadToEnd();
                Debug.WriteLine(Ret);
                //MessageBox.Show(Ret);
                //Debugger.Launch();
                st.Close();
                response.Close();

                return Ret;

            }
            catch (WebException)
            {

                return "";

            }



        }
        private void ListNow(string a1,string a2,bool a3=false) {
            listView1.Clear();
            /**
            string url = "https://boxvideo.sharepoint.cn/";
            Uri CC = new Uri("https://boxvideo.sharepoint.cn");
            //string past = "%27%2Fsites%2Fbox%2FShared%20Documents%27";
            //%2Fsites%2Fbox%2FShared%20Documents%2FCOS
            //url = "https://boxvideo.sharepoint.cn/sites/box/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2Fsites%2Fbox%2FShared%20Documents%27&RootFolder=%2Fsites%2Fbox%2FShared%20Documents%2FCOS&View=d4467c2f-de2c-472e-9ba3-7c565accb7cf&TryNewExperienceSingle=TRUE";
            
            
            url = mm+"/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1="+a1+"&RootFolder="+a2+"&TryNewExperienceSingle=TRUE";
            string post = "{\"parameters\":{\"__metadata\":{\"type\":\"SP.RenderListDataParameters\"},\"RenderOptions\":1446151,\"AllowMultipleValueFilterForTaxonomyFields\":true,\"ViewXml\":\"<RowLimit Paged=\\\"TRUE\\\">99999</RowLimit>\",\"AddRequiredFields\":true}}";
            HttpWebRequest request = null;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            request = WebRequest.Create(url) as HttpWebRequest;
            //HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);


            request.Method = "POST";
            request.KeepAlive = false;
            request.ContentType = "application/json;odata=verbose";

            request.Headers["Cookie"] = cookie;
            
            //request.ProtocolVersion = HttpVersion.Version11;

            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(post);
            request.ContentLength = bytes.Length;
            Stream stream = request.GetRequestStream();
            stream.Write(bytes, 0, bytes.Length);//写入参数
            stream.Close();
            string header = request.Accept;
            System.Diagnostics.Debug.WriteLine(header);
            **/

                
            string Ret = GetJson(a1,a2);
            Debug.WriteLine(Ret);
            if (Ret != "") { 

            var objJSON = JsonConvert.DeserializeObject(Ret);

            JObject jo = (JObject)JsonConvert.DeserializeObject(objJSON.ToString());
            GlobalJs = jo;

            string bb = jo["ListData"]["Row"].ToString();
                //Debug.WriteLine(bb);
            JArray ja = JArray.Parse(bb);


            int count = ja.Count;
            for (int i = 0; i < count; i++)
            {
                listView1.Items.Add(ja[i]["FileLeafRef"].ToString(), i);

            }
            if (a3) {
                root = ja[0]["FileRef.urlencode"].ToString();
                root = root.Substring(0, root.LastIndexOf("%2F"));
                root = root.Substring(0, root.LastIndexOf("%2F"));

            }
            }

        }





        

        private void Form1_Load(object sender, EventArgs e)
        {
            reload();

            ListNow(mainsite, "",true);
            //ListNow("%27%2Fsites%2Fbox%2FShared%20Documents%27", "%2Fsites%2Fbox%2FShared%20Documents%2FCOS");
            init_aria();



        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (mulu != "") {
                string new_mulu = mulu.Substring(0, mulu.LastIndexOf("%2F"));
                //Debug.WriteLine(mulu);
                //Debug.WriteLine(new_mulu);
                if (new_mulu == root) {
                    MessageBox.Show("您已经到达根目录了", "提示");
                    return;
                }
                mulu = new_mulu;
                ListNow(mainsite, mulu);
            }
             //mulu
        }
        public void init_aria() {
            string url = "http://127.0.0.1:6802/jsonrpc?jsonrpc=2.0&method=aria2.tellActive&id=Admire&params=";
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "GET";
            request.KeepAlive = false;
            request.ContentType = "application/json";
            try
            {

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//响应对象
            }

            catch (WebException)
            {

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = @".\aria2c.exe";
                info.Arguments = " --enable-rpc --rpc-allow-origin-all --rpc-listen-port=6802";
                //info.WindowStyle = ProcessWindowStyle.Minimized;
                info.UseShellExecute = false;//是否重定向标准输入 
                info.RedirectStandardInput = false;//是否重定向标准转出 
                info.RedirectStandardOutput = false;//是否重定向错误 
                info.RedirectStandardError = false;//执行时是不是显示窗口 
                info.CreateNoWindow = true;//启动 
                info.WindowStyle = ProcessWindowStyle.Hidden;

                Process pro = Process.Start(info);
                
                //pro.WaitForExit();

            }

            //aria2c --enable-rpc --rpc-allow-origin-all



            //return true;

        
        }
        public void close_aria()
        {

            string url = "http://127.0.0.1:6802/jsonrpc?jsonrpc=2.0&method=aria2.shutdown&id=Admire&params=";
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "GET";
            request.KeepAlive = false;
            request.ContentType = "application/json";
            try
            {

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//响应对象
            }

            catch (WebException)
            {}

            }
        public string request_aria(string method,string param) {
            string url = "http://127.0.0.1:6802/jsonrpc?jsonrpc=2.0&method="+method+"&id=Admire&params="+param;
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "GET";
            request.KeepAlive = false;
            request.ContentType = "application/json";
            try
            {

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//响应对象
                string Ret = "";
                StreamReader st = new StreamReader(response.GetResponseStream(), Encoding.ASCII);
                Ret = st.ReadToEnd();
                st.Close();
                response.Close();
                return Ret;
            }

            catch (WebException)
            {
                return "";
            }
            

        }
         private void listView1_DoubleClick(object sender, EventArgs e)
        {
            //Debug.WriteLine(e.GetType);
            //return;
            //Debug.WriteLine(listView1.SelectedItems.Count.ToString());
            //listView1.SelectedItems.Count
            //listView1.SelectedItems.Count
            //Debug.WriteLine(listView1.Items.Count);
            try
            {
                //Debug.WriteLine(listView1.SelectedItems[0].Text);
                //Debug.WriteLine(listView1.Items.Count);
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    ListViewItem item = listView1.Items[i];
                    //Debug.WriteLine(item.SubItems.Count);
                    //Debug.WriteLine(item.SubItems[0].Text);
                    //Debug.WriteLine(listView1.SelectedItems[0].Text);


                    if (item.SubItems[0].Text == listView1.SelectedItems[0].Text)
                    {
                        //Debug.WriteLine(i);
                        //Debug.WriteLine();
                        string bb = GlobalJs["ListData"]["Row"].ToString();//[8]["FileRef"]["urlencode"]
                        JArray ja = JArray.Parse(bb);

                        Debug.WriteLine(ja[i]["FileRef.urlencode"].ToString());
                        if (ja[i]["ContentType"].ToString() != "文件夹")
                        {
                            MessageBox.Show("您选择的不是文件夹","提示");
                        }
                        else
                        {
                            mulu = ja[i]["FileRef.urlencode"].ToString();
                            ListNow(mainsite, mulu);
                        }
                        //Debug.WriteLine(bb);
                        //string cc = GlobalJs["ListData"]["Row"].ToString();
                    }


                }

            }
            catch
            {

            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            close_aria();

        }

        private void listView1_ContextMenuStripChanged(object sender, EventArgs e)
        {
            
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void contextMenuStrip1_Opening_1(object sender, CancelEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
        //Login.ActiveForm.Activate();
            Login a = new Login();
            a.loging = "https://boxvideo.sharepoint.cn/";
            a.ShowDialog();
            reload();
            ListNow(mainsite, "");


        }
        public static string UrlEncode(string str)
        {
            StringBuilder sb = new StringBuilder();
            byte[] byStr = System.Text.Encoding.UTF8.GetBytes(str); //默认是System.Text.Encoding.Default.GetBytes(str)
            for (int i = 0; i < byStr.Length; i++)
            {
                sb.Append(@"%" + Convert.ToString(byStr[i], 16));
            }

            return (sb.ToString());
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string str = Interaction.InputBox("输入分享链接", "提示", "", -1, -1);

            string url = str;
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "GET";
            request.KeepAlive = false;
            request.ContentType = "application/json";
            request.Headers.Add("Cookie",cookie);
            try
            {

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();//响应对象
                StreamReader st = new StreamReader(response.GetResponseStream(), Encoding.ASCII);
                string Ret = st.ReadToEnd();
                Debug.WriteLine(Ret);
                st.Close();
                response.Close();
                int start = Ret.IndexOf("webAbsoluteUrl");//14
                int last = Ret.IndexOf('"', start+18);
                Debug.WriteLine(start);
                Debug.WriteLine(last);
                Debug.WriteLine(Ret.Substring (start+17, last-start-17));

                mm = Ret.Substring(start + 17, last - start - 17);

                start = Ret.IndexOf("listUrl");//7+3
                last = Ret.IndexOf('"', start + 11);
                Debug.WriteLine(start);
                Debug.WriteLine(last);
                Debug.WriteLine(Ret.Substring(start + 10, last - start -10));
                mainsite = "%27"+UrlEncode(Ret.Substring(start + 10, last - start - 10))+ "%27";
                Debug.WriteLine(mainsite);
                //Base64Encoder Decode = new Base64Encoder();

                //byte[] bytes = System.Text.Encoding.UTF8.GetBytes(Ret.Substring(start + 23, last - start - 23));
                //mainsite = Decode.GetEncoded(bytes);
                mulu = "";
                ListNow(mainsite,mulu);

                
                //Debug.WriteLine(Ret);


            }

            catch (WebException)
            { }

            
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {

 
            
            //listView1.Width = Form1.Width;
            //listView1.Width = Form1.Width;
        }
        private void download_file(string filemu)
        {
            string Ret=GetJson(mainsite, filemu);
            var objJSON = JsonConvert.DeserializeObject(Ret);
            JObject jo = (JObject)JsonConvert.DeserializeObject(objJSON.ToString());
            string bb = jo["ListData"]["Row"].ToString();
            JArray ja = JArray.Parse(bb);
            for (int i = 0; i < ja.Count; i++) {

                if (ja[i]["ContentType"].ToString() == "文件夹")
                {
                    download_file(ja[i]["FileRef.urlencode"].ToString());
                }
                else {

                    string url = "https://boxvideo.sharepoint.cn/" + ja[i]["FileRef.urlencodeasurl"].ToString();

                    string Pa = "[[\"" + url + "\"],{\"out\":\"" + ja[i]["_EditMenuTableStart"].ToString() + "\",\"header\":[\"User-Agent:netdisk\",\"Referer: http://sharepoint.com\",\"Cookie: " + cookie + "\"]}]";
                    Base64Encoder Decode = new Base64Encoder();

                    byte[] bytes = System.Text.Encoding.UTF8.GetBytes(Pa);
                    string rr = Decode.GetEncoded(bytes);
                    //Decode.GetEncoded()
                    Debug.WriteLine(rr);

                    request_aria("aria2.addUri", rr);
                }
            
            }




        }
        private void 下载_Click(object sender, EventArgs e)
        {

            try
            {

                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    ListViewItem item = listView1.Items[i];

                    if (item.SubItems[0].Text == listView1.SelectedItems[0].Text)
                    {

                        string bb = GlobalJs["ListData"]["Row"].ToString();//[8]["FileRef"]["urlencode"]
                        JArray ja = JArray.Parse(bb);

                        Debug.WriteLine(ja[i]["FileRef.urlencode"].ToString());
                        if (ja[i]["ContentType"].ToString() == "文件夹")
                        {
                            //MessageBox.Show("您选择的是文件夹", "提示");文件夹循环下载
                            //MessageBox.Show("正在写中", "提示");
                            download_file(ja[i]["FileRef.urlencode"].ToString());

                        }
                        else
                        {
                            //mulu = ja[i]["FileRef.urlencode"].ToString();
                            //ListNow("%27%2Fsites%2Fbox%2FShared%20Documents%27", mulu);
                            
                            string url = "https://boxvideo.sharepoint.cn/" + ja[i]["FileRef.urlencodeasurl"].ToString();

                            string Pa= "[[\""+url+"\"],{\"out\":\""+ ja[i]["_EditMenuTableStart"].ToString() + "\",\"header\":[\"User-Agent:netdisk\",\"Referer: http://sharepoint.com\",\"Cookie: "+cookie+"\"]}]";
                            Base64Encoder Decode = new Base64Encoder();
                            
                            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(Pa);
                            string rr= Decode.GetEncoded(bytes);
                            //Decode.GetEncoded()
                            Debug.WriteLine(rr);

                            request_aria("aria2.addUri", rr);


                        }
                    }


                }

            }
            catch
            {}

        }
    }
}
