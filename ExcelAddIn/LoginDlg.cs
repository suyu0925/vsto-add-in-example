using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class LoginDlg : Form
    {
        public LoginDlg()
        {
            InitializeComponent();
            Load += InitWhenLoaded;
        }

        private async void InitWhenLoaded(object sender, EventArgs e)
        { 
            var env = await CoreWebView2Environment.CreateAsync(null, Config.UserDataFolder);
            await webView.EnsureCoreWebView2Async(env);

            webView.Source = new Uri("https://bing.com");
        }        
    }
}
