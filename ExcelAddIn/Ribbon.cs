using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnLogin(Office.IRibbonControl control)
        {
            using (LoginDlg dlg = new LoginDlg())
            {
                dlg.StartPosition = FormStartPosition.CenterScreen;
                DialogResult result = dlg.ShowDialog();
            }
        }

        public bool IsControlVisible(Office.IRibbonControl control)
        {
            return true;
        }

        public void OnDelayDo(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.InvokeAsyncCallToExcel();
            // Thread t = new Thread(() =>
            // {
            //     MessageFilter.Register();
            //     Thread.Sleep(2000);
            //     try
            //     {
            //         Globals.ThisAddIn.Application.ActiveCell.Value2 =
            //         DateTime.Now.ToShortTimeString();
            //     }
            //     catch (Exception ex)
            //     {
            //         Debug.WriteLine(ex.ToString());
            //     }
            // });
            // t.SetApartmentState(ApartmentState.STA);
            // t.Start();
        }

        public void OnDoSomethingOnThread(Office.IRibbonControl control)
        {
            Thread t = new Thread(() =>
            {
                MessageFilter.Register();
                CallExcel();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        public async void OnDoSomething(Office.IRibbonControl control)
        {
            MessageFilter.Register();
            await Task.Run(new System.Action(CallExcel));
        }

        public void OnHelp(Office.IRibbonControl control)
        {
            MessageBox.Show("help");
        }
        #endregion

        private void CallExcel()
        {
            try
            {
                var currSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                int rowSize = 50;
                int colSize = 50;

                for (int i = 1; i <= rowSize; i++)
                    for (int j = 1; j <= colSize; j++)
                        ((Range)currSheet.Cells[i, j]).Value2 = "sample";
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        public class MessageFilter : IOleMessageFilter
        {
            //
            // Class containing the IOleMessageFilter
            // thread error-handling functions.

            // Start the filter.
            public static void Register()
            {
                IOleMessageFilter newFilter = new MessageFilter();
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(newFilter, out oldFilter);
                Debug.WriteLine($"MessageFilter.Register, oldFilter: {oldFilter}");
            }

            // Done with the filter, close it.
            public static void Revoke()
            {
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(null, out oldFilter);
                Debug.WriteLine($"MessageFilter.Revoke, oldFilter: {oldFilter}");
            }

            // Implement the IOleMessageFilter interface.
            [DllImport("Ole32.dll")]
            private static extern int
              CoRegisterMessageFilter(IOleMessageFilter newFilter, out
              IOleMessageFilter oldFilter);

            //
            // IOleMessageFilter functions.
            // Handle incoming thread requests.
            int IOleMessageFilter.HandleInComingCall(int dwCallType,
              System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
              lpInterfaceInfo)
            {
                //Return the flag SERVERCALL_ISHANDLED.
                return 0;
            }

            // Thread call was rejected, so try again.
            int IOleMessageFilter.RetryRejectedCall(System.IntPtr
              hTaskCallee, int dwTickCount, int dwRejectType)
            {
                Debug.WriteLine($"RetryRejectedCall {hTaskCallee} {dwTickCount} {dwRejectType}");
                if (dwRejectType == 2)
                // flag = SERVERCALL_RETRYLATER.
                {
                    // Retry the thread call immediately if return >=0 & <100.
                    // COM will wait for this many milliseconds and then retry the call.
                    return 200;
                }
                // Too busy; cancel call.
                return -1;
            }

            int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
              int dwTickCount, int dwPendingType)
            {
                //Return the flag PENDINGMSG_WAITDEFPROCESS.
                return 2;
            }
        }

        [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
        InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(
                int dwCallType,
                IntPtr hTaskCaller,
                int dwTickCount,
                IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(
                IntPtr hTaskCallee,
                int dwTickCount,
                int dwRejectType);

            [PreserveSig]
            int MessagePending(
                IntPtr hTaskCallee,
                int dwTickCount,
                int dwPendingType);
        }
        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}
