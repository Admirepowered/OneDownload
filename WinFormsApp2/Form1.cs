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

namespace WinFormsApp2
{
   // public static string Cookie = "";
    
    public partial class Form1 : Form
    {
        //public string cookie = "rtFa=nPvgjXhD+g8SbZf50BODSJTuycj8EbWa/xz8Hsu5ugQmOTQwRjNEQjYtREI2MS00QzIwLTk1OTgtOEJGNzAyRDE2M0I5o2wrlP6lr3PqLejCYJlLlruL4d9DSc+jLrxfn/lywR0JapxDn3rFuBfBGPbk3sOZAWSgMc3S1SPiciVxdSt7CY1G9vx9wyCwHVbDhVIR4GR5bS6Iu+11SnQYkfU6x+H5LD9qsSLVkagy9XjY+hxv1XjZ/npE2p37hcWlMhmTnZUdAmvuAGCF74UMjY7nbk+atQTliTfIc2E3b5sybPy00FLrI00qEtxbOCCkAgg2n/fQPrqRnB3J/JB7yeZZoqwvf6wix8hzcX7SvMnxTTx+vREQEiV0rEr6cmLqymIG17gA4Edxw31KxTPR8AQXfAnloWa/ehmhXs9hREGVqhkt0UUAAAA=; FedAuth=77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48U1A+VjEwLDBoLmZ8bWVtYmVyc2hpcHwxMDAzMzIzMGM1Y2Q3NTI5QGxpdmUuY29tLDAjLmZ8bWVtYmVyc2hpcHxhZG1pcmVAYm94dmlkZW8ucGFydG5lci5vbm1zY2hpbmEuY24sMTMyNjI0NjEyNTgwMDAwMDAwLDEzMjYyMDE0NDc4MDAwMDAwMCwxMzI2NzY0MjU1ODQyMzI2OTAsMTAxLjIwNy4xNy4xMjUsMyw5NDBmM2RiNi1kYjYxLTRjMjAtOTU5OC04YmY3MDJkMTYzYjksLDFlY2ZjYmEzLWE1NTctNDFkNy1hODIxLTA1N2NjNGI4Y2Q4Ziw1YzRkY2U5Zi1hMGI0LTAwMDAtMTIwNy1jNGVlZmJjYTIyNGEsNWM0ZGNlOWYtYTBiNC0wMDAwLTEyMDctYzRlZWZiY2EyMjRhLCwwLDEzMjY3MjE0MTU4Mjk4Mjc4OSwxMzI2NzQ2OTc1ODI5ODI3ODksLCwsMjY1MDQ2Nzc0Mzk5OTk5OTk5OSwxMzI2NzIxMDU1ODAwMDAwMDAsN2MyMzczNGQtZjNiMS00MDgzLWEwYzMtY2E0MGY5ZmQ5ZDQxLCwsLCwsYnlzRmNPS3NUcXZIVXlNK2pXL3Zha3FTMkN4Zm5vcEQvUzdXN3Y0eFlRcC9ZZXlkcFZ5SDlwaW51S1lVY0lRMEt1cDhoWVJ6a3ozNHYrK3ZGdVFQeUE4L3RPSG5FS3ZJMEo4NS81cDdKUURUYmRUN3pVVFN5OUpNb3Q0alg4RUllYjlGWERiTGtrVjV0WnFSbTAzYmlxVUZQYTFaYU5DNWdZd3RsT0w5VlV3U0xoeEpoNFFISHJMUGI2TkZEWmROMVo1U3ZVSE9nZ0pBOXczYnNuT05aMkpyWTJ6cjMwU0lhYzZLUG52VE9YZXlBcW4vT1ZiZW5jRmhOazJLdTkxRTJycTFhRmFJWnNGNzVQbTdwRC9HTlVwNFFMejhXRmJFSWNlUEg5am9xTW9MdXBOOWU3L0ZMSmhiRHhIMWRieHJtRDgyalpLUFVwN1hzOXVqS00xU1lnPT08L1NQPg==; SPWorkLoadAttribution=Url=https://boxvideo.sharepoint.cn/&AppTitle=RenderClientSideBasePage; ScaleCompatibilityDeviceId=08e50dfe-d4a9-45d4-9b72-d821da232edd";
        public string cookie = "";
        
        public JObject GlobalJs;
        public string mulu = "";
        public string root = "";

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
            url = "https://boxvideo.sharepoint.cn/sites/box/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=" + a1 + "&RootFolder=" + a2 + "&TryNewExperienceSingle=TRUE";
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
            
            
            url = "https://boxvideo.sharepoint.cn/sites/box/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1="+a1+"&RootFolder="+a2+"&TryNewExperienceSingle=TRUE";
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

            ListNow("%27%2Fsites%2Fbox%2FShared%20Documents%27", "",true);
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
                ListNow("%27%2Fsites%2Fbox%2FShared%20Documents%27", mulu);
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
                            ListNow("%27%2Fsites%2Fbox%2FShared%20Documents%27", mulu);
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
            ListNow("%27%2Fsites%2Fbox%2FShared%20Documents%27", "");


        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {

 
            
            //listView1.Width = Form1.Width;
            //listView1.Width = Form1.Width;
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
                            MessageBox.Show("正在写中", "提示");
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
