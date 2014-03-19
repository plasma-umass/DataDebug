using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataDebugMethods
{
    /// <summary>
    /// Data structure for representing nodes of the dependence graph (DAG) internally
    /// There are three types of nodes: normal cells, ranges, and charts
    /// The different types are distinguished by the name
    /// For normal cells, the name is just the address of the cell
    /// For ranges, the name is of the format <EndCell>:<EndCell>, such as "A1:A5"
    /// For chart nodes, the name begins with the string "Chart", followed by the name of the chart object from Excel, with the white spaces stripped
    /// </summary>
    public class TreeNode
    {
        Excel.Workbook _workbook;
        private HashSet<TreeNode> _inputs; // these are the TreeNodes that feed into the current cell
        private HashSet<TreeNode> _outputs;    //these are the TreeNodes that the current cell feeds into
        private string _name;    //The name of each node: for cells, it is its address as a string; for ranges, it is of the form <EndCell>:<EndCell>; for charts it is "Chart<Name of chart>"
        private string _worksheet_name;  //This keeps track of the worksheet where this cell/range/chart is located
        private Excel.Worksheet _worksheet; // A reference to the actual worksheet where this TreeNode is located
        private double _weight = 0.0;  //The weight of the node as computed by propagating values down the tree
        private bool _chart;
        private bool _is_formula = false; //this indicates whether this node is a formula
        private string _formula;
        private Excel.Range _COM;
        private bool _dont_perturb = false; // this flag indicates that this is an input range which contains non-perturbable elements, like function outputs.  We assume, by default, that all treenodes are perturbable.
        private int _height = 1;
        private int _width = 1;
        private AST.Address _addr;
        bool _is_a_cell = false;
        public TreeNode(Excel.Range com, Excel.Worksheet ws, Excel.Workbook wb)
        {
            _inputs = new HashSet<TreeNode>();
            _outputs = new HashSet<TreeNode>();
            _name = String.Intern(wb.FullName + ws.Name + com.Address);
            _height = com.Rows.Count;
            _width = com.Columns.Count;
            _is_a_cell = _height == 1 && _width == 1;
            _worksheet = ws;
            if (_worksheet == null)
            {
                _worksheet_name = "none";
            }
            else
            {
                _worksheet_name = ws.Name;
            }
            _workbook = wb;
            _weight = 0.0;
            _chart = false;
            _COM = com;
            
            // save parsed address of THIS cell;
            // used frequently in equality comparisons
            if (_is_a_cell)
            {
                _addr = AST.Address.AddressFromCOMObject(_COM, _workbook);
            }

            // node is formula iff the COM object is both a single cell and a formula
            if (_height == 1 && _width == 1 && com.HasFormula == true)
            {
                _is_formula = true;
                _formula = com.Formula;
            }
        }

        //public override bool Equals(object obj)
        //{
        //    return _addr == ((TreeNode)obj).GetAddress();
        //}

        public override bool Equals(object o)
        {
            return _name == ((TreeNode)o)._name;
        }

        public override int GetHashCode()
        {
            //return _addr.AddressAsInt32();
            return _name.GetHashCode();
        }

        public AST.Address GetAddress()
        {
            if (!_is_a_cell)
            {
                throw new Exception("Cannot get AST.Address for a TreeNode representing a range of cells.");
            }
            return _addr;
        }

        public int Columns() 
        { 
            return _width; 
        }
        
        public int Rows() 
        { 
            return _height; 
        }

        public void DontPerturb()
        {
            _dont_perturb = true;
        }

        public void Perturb()
        {
            _dont_perturb = false;
        }

        public bool GetDontPerturb()
        {
            return _dont_perturb;
        }

        public string toString()
        {
            string parents_string = "";
            foreach (TreeNode node in _inputs)
            {
                parents_string += node.getWorksheetName() + " " + node.getName() + ", ";
            }
            string children_string = "";
            foreach (TreeNode node in _outputs)
            {
                children_string += node.getName() + ", ";
            }
            return _name + Environment.NewLine + "Parents: " + parents_string + Environment.NewLine + "Children: " + children_string;
        }

        //Returns the name of the node
        public string getName()
        {
            return _name;
        }

        //Returns the workbook object of the node
        public Excel.Workbook getWorkbookObject()
        {
            return _workbook;
        }
        //Returns the weight of this node
        public double getWeight()
        {
            return _weight;
        }

        //Sets the weight of the node to the double passed as an argument
        public void setWeight(double w)
        {
            _weight = w;
        }

        // adds an input to a TreeNode's input list
        public void addInput(TreeNode node)
        {
            // never add self
            if (node == this)
            {
                throw new Exception(String.Format("Attempted to add {0} as an input to itself.", this._name));
            }
            // never re-add input
            if (_inputs.Contains(node))
            {
                return;
            }
            _inputs.Add(node);
        }

        // adds an output to a TreeNode's output list
        public void addOutput(TreeNode node)
        {
            // never add self
            if (node == this)
            {
                throw new Exception(String.Format("Attempted to add {0} as an output to itself.", this._name));
            }
            // never re-add output
            if (_outputs.Contains(node))
            {
                return;
            }
            _outputs.Add(node);
        }

        //Checks if this node has any children
        public bool hasOutputs()
        {
            if (_outputs.Count == 0)
                return false;
            else
                return true;
        }

        //Checks if this node has any parents
        public bool hasInputs()
        {
            if (_inputs.Count == 0)
                return false;
            else
                return true;
        }

        //By convention, we name ranges with the string ":" separating the end cells, such as "A1:A5"
        //If the name contains an underscore, and it is not a Chart node, then it is a Range node
        public bool isRange()
        {
            if (_name.Contains(":") && !isChart())
                return true;
            else
                return false;
        }

        public bool isChart()
        {
            return _chart;
        }

        public void setChart(bool value)
        {
            _chart = value;
        }

        //Retuns the List<TreeNode> of children of this node
        public HashSet<TreeNode> getOutputs()
        {
            return _outputs;
        }

        //Retuns the List<TreeNode> of parents of this node
        public HashSet<TreeNode> getInputs()
        {
            return _inputs;
        }

        //Returns the name of the worksheet that holds this cell/range/chart
        public string getWorksheetName()
        {
            return _worksheet_name;
        }

        // Returns a reference to the worksheet that contains this TreeNode
        public Excel.Worksheet getWorksheetObject()
        {
            return _worksheet;
        }
        //Sets the name of the worksheet that holds this cell/range/chart to the argument string s
        public void setWorksheet(string s)
        {
            _worksheet_name = s;
        }

        public void setIsFormula()
        {
            _is_formula = true;
        }

        public bool isFormula()
        {
            return _is_formula;
        }

        /**
         * This is a recursive method for propagating the weights down the nodes in the tree
         * All outputs have weight 1. Their n children have weight 1/n, and so forth. 
         * Modifies the weights set on the trees.
         * TODO: make this an instance method.
         */
        public static void propagateWeight(TreeNode node, double passed_down_weight)
        {
            if (!node.hasInputs())
            {
                return;
            }
            else
            {
                int denominator = 0;  //keeps track of how many objects we are dividing the influence by
                foreach (TreeNode parent in node.getInputs())
                {
                    if (parent.isRange() || parent.isChart())
                        denominator = denominator + parent.getInputs().Count;
                    else
                        denominator = denominator + 1;
                }
                foreach (TreeNode parent in node.getInputs())
                {
                    if (parent.isRange() || parent.isChart())
                    {
                        parent.setWeight(parent.getWeight() + passed_down_weight * parent.getInputs().Count / denominator);
                        propagateWeight(parent, passed_down_weight * parent.getInputs().Count / denominator);
                    }
                    else
                    {
                        parent.setWeight(parent.getWeight() + passed_down_weight / node.getInputs().Count);
                        propagateWeight(parent, passed_down_weight / node.getInputs().Count);
                    }
                }
            }
        }

        internal void setFormula(string formula)
        {
            _formula = formula;
        }

        public string getFormula()
        {
            return _formula;
        }

        public Excel.Range getCOMObject()
        {
            return _COM;
        }

        public string getCOMValueAsString()
        {
            return System.Convert.ToString(this.getCOMObject().Value2);
        }

        public bool isLeaf()
        {
            return _inputs.Count == 0;
        }
    }
}
