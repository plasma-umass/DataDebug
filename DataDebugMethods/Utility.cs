using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace DataDebugMethods
{
    class Utility
    {
        static Excel.Workbook OpenWorkbook(string filename, Excel.Application app)
        {
            // we need to disable all alerts, e.g., password prompts, etc.
            app.DisplayAlerts = false;

            // disable macros
            app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // This call is stupid.  See:
            // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
            app.Workbooks.Open(filename, 2, true, Missing.Value, "thisisnotapassword", Missing.Value, true, Missing.Value, Missing.Value, Missing.Value, false, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            return app.Workbooks[1];
        }
    }
}
