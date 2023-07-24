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
    public partial class TestForm : Form
    {
        public enum FormType {
            NonblockTask,
            BlockTask,
            ThrowError,
        }

        readonly FormType formType;

        public TestForm(FormType formType)
        {
            this.formType = formType;

            InitializeComponent();

            Load += InitWhenLoaded;
        }
        
        private async void InitWhenLoaded(object sender, EventArgs e)
        {
            switch (formType) {
                case FormType.BlockTask:
                    System.Threading.Thread.Sleep(60 * 1000);
                    break;
                case FormType.ThrowError:
                    throw new Exception("ThorwError from TestForm");
                case FormType.NonblockTask:
                    await Task.Delay(60 * 1000);
                    break;
            }
        }
    }
}
