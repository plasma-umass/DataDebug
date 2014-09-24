using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace DataDebug
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /*
             * NOTE: DO NOT ADD ANY UI CODE HERE. ANYTHING THAT REQUIRES
             * USER INTERACTION WILL BREAK AUTOMATION CODE THAT USES
             * THE PLUGIN!!!
             */

            WorkbookOpen(this.Application.ActiveWorkbook);
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += WorkbookOpen;
            this.Application.WorkbookOpen += WorkbookOpen;
            this.Application.WorkbookActivate += WorkbookActivated;
            this.Application.WorkbookDeactivate += WorkbookDeactivated;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            /*
             * NOTE: DO NOT ADD ANY UI CODE HERE. ANYTHING THAT REQUIRES
             * USER INTERACTION WILL BREAK AUTOMATION CODE THAT USES
             * THE PLUGIN!!!
             */
        }

        private void WorkbookOpen(Excel.Workbook workbook)
        {
            MessageBox.Show("Workbook added.");
        }

        private void WorkbookActivated(Excel.Workbook workbook)
        {
            MessageBox.Show("Workbook activated.");
        }

        private void WorkbookDeactivated(Excel.Workbook workbook)
        {
            MessageBox.Show("Workbook deactivated.");
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