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
        private List<TreeNode> _parents;  //these are the TreeNodes that feed into the current cell
        private List<TreeNode> _children;    //these are the TreeNodes that the current cell feeds into
        private string _name;    //The name of each node: for cells, it is its address as a string; for ranges, it is of the form <EndCell>:<EndCell>; for charts it is "Chart<Name of chart>"
        private string _worksheet_name;  //This keeps track of the worksheet where this cell/range/chart is located
        private Excel.Worksheet _worksheet; // A reference to the actual worksheet where this TreeNode is located
        private double _weight = 0.0;  //The weight of the node as computed by propagating values down the tree
        private bool _chart;
        private bool _is_formula; //this indicates whether this node is a formula
        private System.Drawing.Color originalColor;
        private string _formula;
        private Excel.Range _COM;
        private bool _dont_perturb = false; // this flag indicates that this is an input range which contains non-perturbable elements, like function outputs.  We assume, by default, that all treenodes are perturbable.
        private int _height = 1;
        private int _width = 1;
        public TreeNode(Excel.Range com, Excel.Worksheet ws, Excel.Workbook wb)
        {
            _parents = new List<TreeNode>();
            _children = new List<TreeNode>();
            _name = com.Address;
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
            _height = com.Rows.Count;
            _width = com.Columns.Count;
            if (_height == 1 && _width == 1)
            {
                if (com.HasFormula == true)
                {
                    _is_formula = true;
                    _formula = com.Formula;
                }
                else
                {
                    _is_formula = false;
                }
            }
            else
            {
                _is_formula = false;
            }
            _dont_perturb = true;
        }

        public AST.Address GetAddress()
        {
            return AST.Address.AddressFromCOMObject(_COM,
                                                    new Microsoft.FSharp.Core.FSharpOption<string>(_worksheet_name),
                                                    new Microsoft.FSharp.Core.FSharpOption<string>(_workbook.Name),
                                                    new Microsoft.FSharp.Core.FSharpOption<string>(_workbook.FullName));
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
            foreach (TreeNode node in _parents)
            {
                parents_string += node.getWorksheet() + " " + node.getName() + ", ";
            }
            string children_string = "";
            foreach (TreeNode node in _children)
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

        //Adds a parent to the list of parent nodes; checks for duplicates before adding it
        public void addInput(TreeNode node)
        {
            //Make sure we are not adding a parent more than once
            bool parent_already_added = false;
            foreach (TreeNode n in _parents)
            {
                if (node.getName() == n.getName())
                    parent_already_added = true;
            }
            //If the parent is not on the list, add it
            if (!parent_already_added)
                _parents.Add(node);
        }

        //Adds a child to the list of child nodes; checks for duplicates before adding it
        public void addOutput(TreeNode node)
        {
            //Make sure we are not adding a child more than once
            bool child_already_added = false;
            foreach (TreeNode n in _children)
            {
                if (node.getName() == n.getName())
                    child_already_added = true;
            }
            //If the child is not on the list, add it
            if (!child_already_added)
                _children.Add(node);
        }

        //Checks if this node has any children
        public bool hasChildren()
        {
            if (_children.Count == 0)
                return false;
            else
                return true;
        }

        //Checks if this node has any parents
        public bool hasParents()
        {
            if (_parents.Count == 0)
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
        public List<TreeNode> getChildren()
        {
            return _children;
        }

        //Retuns the List<TreeNode> of parents of this node
        public List<TreeNode> getParents()
        {
            return _parents;
        }

        //Returns the name of the worksheet that holds this cell/range/chart
        public string getWorksheet()
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
            if (_is_formula == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /**
         * This is a recursive method for propagating the weights down the nodes in the tree
         * All outputs have weight 1. Their n children have weight 1/n, and so forth. 
         * Modifies the weights set on the trees.
         * TODO: make this an instance method.
         */
        public static void propagateWeight(TreeNode node, double passed_down_weight)
        {
            if (!node.hasParents())
            {
                return;
            }
            else
            {
                int denominator = 0;  //keeps track of how many objects we are dividing the influence by
                foreach (TreeNode parent in node.getParents())
                {
                    if (parent.isRange() || parent.isChart())
                        denominator = denominator + parent.getParents().Count;
                    else
                        denominator = denominator + 1;
                }
                foreach (TreeNode parent in node.getParents())
                {
                    if (parent.isRange() || parent.isChart())
                    {
                        parent.setWeight(parent.getWeight() + passed_down_weight * parent.getParents().Count / denominator);
                        propagateWeight(parent, passed_down_weight * parent.getParents().Count / denominator);
                    }
                    else
                    {
                        parent.setWeight(parent.getWeight() + passed_down_weight / node.getParents().Count);
                        propagateWeight(parent, passed_down_weight / node.getParents().Count);
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
    }
}
