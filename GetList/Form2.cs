using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetList
{
    public partial class Form2 : Form
    {

        WebBrowser webBrowser1;
        public Form2()
        {
           
            InitializeComponent();

            //webBrowser1 = new WebBrowser();
        }

      
        private void Form2_Load(object sender, EventArgs e)
        {
            //webBrowser1 = new WebBrowser();

            //webBrowser1.AllowWebBrowserDrop = false;
            //webBrowser1.IsWebBrowserContextMenuEnabled = false;
            //webBrowser1.WebBrowserShortcutsEnabled = false;
            //webBrowser1.ObjectForScripting = this;
            // Uncomment the following line when you are finished debugging.


        }

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

          
            webBrowser1.DocumentCompleted -= WebBrowser1_DocumentCompleted;

            HtmlElement search = webBrowser1.Document.GetElementById("name1");
            if (search != null)
            {
                search.SetAttribute("value",textBox1.Text);
                HtmlElement searchButton = webBrowser1.Document.GetElementById("checkCompanyName_0");

                searchButton.InvokeMember("click");

                webBrowser1.DocumentCompleted += WebBrowser1_DocumentCompleted1;


                //if (webBrowser2.ReadyState == WebBrowserReadyState.Complete)
                //{
                //    HtmlElement resultTable = webBrowser2.Document.GetElementById("companyList");
                //}
                //



                //foreach (HtmlElement ele in search.Parent.Children)
                //{
                //    if (ele.TagName.ToLower() == "input" && ele.Name.ToLower() == "go")
                //    {
                //        ele.InvokeMember("click");
                //        break;
                //    }
                //}
            }
        }

        private void WebBrowser1_DocumentCompleted1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string value = string.Empty;
            HtmlElement resultTable = webBrowser1.Document.GetElementById("companyList");

            if(resultTable != null)
            {
                HtmlElementCollection input1 = resultTable.GetElementsByTagName("TD");

                foreach (HtmlElement input2 in input1)
                {
                    value += input2.InnerText + "\n";


                }
                label3.Text = value;
            }
            else
            {
                label3.Text = "Not Found";
            }
            

            
        }

        void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            webBrowser1 = new WebBrowser();
            webBrowser1.ScriptErrorsSuppressed = true;

            webBrowser1.DocumentCompleted += WebBrowser1_DocumentCompleted;

            //webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);
            webBrowser1.Navigate("http://www.mca.gov.in/mcafoportal/showCheckCompanyName.do");
        }
    }
}
