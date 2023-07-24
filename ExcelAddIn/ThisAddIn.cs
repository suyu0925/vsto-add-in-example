using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace ExcelAddIn
{
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

    public partial class ThisAddIn : IOleMessageFilter
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

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
            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER.
            {
                // Retry the thread call immediately if return >=0 & 
                // <100.
                return 99;
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

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int
          CoRegisterMessageFilter(IOleMessageFilter newFilter, out
          IOleMessageFilter oldFilter);

        private IOleMessageFilter oldMessageFilter;

        internal void InvokeAsyncCallToExcel()
        {
            Thread t = new Thread(this.RegisterFilter);
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        private void RegisterFilter()
        {
            int a = CoRegisterMessageFilter(this, out oldMessageFilter);
            Debug.WriteLine($"CoRegisterMessageFilter return {a}");
            Thread.Sleep(3000);
            CallExcel();
        }

        private void heavy()
        {
            var currSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            int rowSize = 50;
            int colSize = 50;

            for (int i = 1; i <= rowSize; i++)
                for (int j = 1; j <= colSize; j++)
                    ((Excel.Range)currSheet.Cells[i, j]).Value2 = "sample";
        }

        private void light()
        {
            this.Application.ActiveCell.Value2 = DateTime.Now.ToShortTimeString();
        }

        private void CallExcel()
        {
            try
            {
                heavy();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
