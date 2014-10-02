using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.FSharp.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataDebugMethods
{
    public class COMRef
    {
        private HashSet<COMRef> _inputs;
        private HashSet<COMRef> _outputs;
        private Excel.Workbook _wb;
        private Excel.Worksheet _ws;
        private Excel.Range _r;
        private bool _is_cell;
        private string _interned_unique_id;
        private int _width;
        private int _height;
        private string _workbook_name;
        private string _worksheet_name;
        private FSharpOption<string> _formula;
        private bool _do_not_perturb;

        public COMRef(string unique_id,
                      Excel.Workbook wb,
                      Excel.Worksheet ws,
                      Excel.Range r,
                      string workbook_name,
                      string worksheet_name,
                      FSharpOption<string> formula,
                      int width,
                      int height)
        {
            _wb = wb;
            _ws = ws;
            _r = r;
            _interned_unique_id = String.Intern(unique_id);
            _width = width;
            _height = height;
            _workbook_name = workbook_name;
            _worksheet_name = worksheet_name;
            _do_not_perturb = FSharpOption<string>.get_IsSome(_formula);
        }

        public Excel.Workbook Workbook
        {
            get { return _wb; }
        }

        public Excel.Worksheet Worksheet
        {
            get { return _ws; }
        }

        public Excel.Range Range
        {
            get { return _r; }
        }

        public bool IsFormula
        {
            get { return FSharpOption<string>.get_IsSome(_formula); }
        }

        public string Formula
        {
            get {
                if (FSharpOption<string>.get_IsNone(_formula))
                {
                    throw new Exception("Not a formula reference.");
                }
                else
                {
                    return _formula.Value;
                }
            }
        }

        public bool IsCell
        {
            get { return _width == 1 && _height == 1; ; }
        }

        public string UniqueID
        {
            get { return _interned_unique_id; }
        }

        public int Width
        {
            get { return _width;  }
        }

        public int Height
        {
            get { return _height; }
        }

        public string WorkbookName
        {
            get { return _workbook_name; }
        }

        public string WorksheetName
        {
            get { return _worksheet_name; }
        }

        public bool DoNotPerturb
        {
            get { return _do_not_perturb; }
            set { _do_not_perturb = value; }
        }

        public override int GetHashCode()
        {
            // equivalent strings always point
            // to the same address; thus references
            // are guaranteed to be good hash codes
            return _interned_unique_id.GetHashCode();
        }

        public HashSet<COMRef> getInputs()
        {
            return _inputs;
        }

        public HashSet<COMRef> getOutputs()
        {
            return _outputs;
        }
    }
}
