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
    //public class TreeNode
    //{
    //    private COMRef _com;
    //    private HashSet<TreeNode> _inputs; 
    //    private HashSet<TreeNode> _outputs;
    //    private double _weight = 0.0;
    //    private bool _dont_perturb = false; // True if the treenode contains non-perturbable elements, like function outputs.
    //    private AST.Address _addr;
    //    bool _is_a_cell = false;
    //    public TreeNode(AST.Address address, COMRef cr, bool is_cell, bool is_formula)
    //    {
    //        _addr = address;
    //        _com = cr;
    //        _inputs = new HashSet<TreeNode>();
    //        _outputs = new HashSet<TreeNode>();
    //        _weight = 0.0;
    //    }

    //    public override bool Equals(object o)
    //    {
    //        return _com.UniqueID == ((TreeNode)o)._com.UniqueID;
    //    }

    //    public override int GetHashCode()
    //    {
    //        return _com.GetHashCode();
    //    }

    //    public AST.Address GetAddress()
    //    {
    //        if (!_is_a_cell)
    //        {
    //            throw new Exception("Cannot get AST.Address for a TreeNode representing a range of cells.");
    //        }
    //        return _addr;
    //    }

    //    public int Columns() 
    //    { 
    //        return _com.Width; 
    //    }
        
    //    public int Rows() 
    //    {
    //        return _com.Height;
    //    }

    //    public void SetDoNotPerturb()
    //    {
    //        _dont_perturb = true;
    //    }

    //    public void Perturb()
    //    {
    //        _dont_perturb = false;
    //    }

    //    public bool GetDontPerturb()
    //    {
    //        return _dont_perturb;
    //    }

    //    public string toString()
    //    {
    //        string parents_string = "";
    //        foreach (TreeNode node in _inputs)
    //        {
    //            parents_string += node.getWorksheetName() + " " + node.getName() + ", ";
    //        }
    //        string children_string = "";
    //        foreach (TreeNode node in _outputs)
    //        {
    //            children_string += node.getName() + ", ";
    //        }
    //        return _com.UniqueID + Environment.NewLine + "Parents: " + parents_string + Environment.NewLine + "Children: " + children_string;
    //    }

    //    //Returns the name of the node
    //    public string getName()
    //    {
    //        return _com.UniqueID;
    //    }

    //    //Returns the workbook object of the node
    //    public Excel.Workbook getWorkbookObject()
    //    {
    //        return _com.Workbook;
    //    }

    //    //Returns the weight of this node
    //    public double getWeight()
    //    {
    //        return _weight;
    //    }

    //    //Sets the weight of the node to the double passed as an argument
    //    public void setWeight(double w)
    //    {
    //        _weight = w;
    //    }

    //    // adds an input to a TreeNode's input list
    //    public void addInput(TreeNode node)
    //    {
    //        // never add self
    //        if (node == this)
    //        {
    //            throw new Exception(String.Format("Attempted to add {0} as an input to itself.", _com.UniqueID));
    //        }
    //        // never re-add input
    //        if (_inputs.Contains(node))
    //        {
    //            return;
    //        }
    //        _inputs.Add(node);
    //    }

    //    // adds an output to a TreeNode's output list
    //    public void addOutput(TreeNode node)
    //    {
    //        // never add self
    //        if (node == this)
    //        {
    //            throw new Exception(String.Format("Attempted to add {0} as an output to itself.", _com.UniqueID));
    //        }
    //        // never re-add output
    //        if (_outputs.Contains(node))
    //        {
    //            return;
    //        }
    //        _outputs.Add(node);
    //    }

    //    //Checks if this node has any children
    //    public bool hasOutputs()
    //    {
    //        if (_outputs.Count == 0)
    //            return false;
    //        else
    //            return true;
    //    }

    //    //Checks if this node has any parents
    //    public bool hasInputs()
    //    {
    //        if (_inputs.Count == 0)
    //            return false;
    //        else
    //            return true;
    //    }

    //    public bool isCell()
    //    {
    //        return _is_a_cell;
    //    }

    //    //Retuns the List<TreeNode> of children of this node
    //    public HashSet<TreeNode> getOutputs()
    //    {
    //        return _outputs;
    //    }

    //    //Retuns the List<TreeNode> of parents of this node
    //    public HashSet<TreeNode> getInputs()
    //    {
    //        return _inputs;
    //    }

    //    //Returns the name of the worksheet that holds this cell/range/chart
    //    public string getWorksheetName()
    //    {
    //        return _com.WorksheetName;
    //    }

    //    // Returns a reference to the worksheet that contains this TreeNode
    //    public Excel.Worksheet getWorksheetObject()
    //    {
    //        return _com.Worksheet;
    //    }

    //    public bool isFormula()
    //    {
    //        return _com.IsFormula;
    //    }

    //    public string ToDOT(HashSet<AST.Address> visited)
    //    {
    //        // base case 1: loop protection
    //        if (visited.Contains(this.GetAddress()))
    //        {
    //            return "";
    //        }

    //        // base case 2: an input
    //        if (!_com.IsFormula && _is_a_cell)
    //        {
    //            return "";
    //        }
    //        // recursive case: we're a node with an input
    //        String s = "";
    //        foreach (TreeNode t in this.getInputs())
    //        {
    //            // print
    //            s += t.GetAddress().A1Local() + " -> " + this.GetAddress().A1Local() + ";\n";
    //            // mark visit
    //            visited.Add(this.GetAddress());
    //            // recurse
    //            s += t.ToDOT(visited);
    //        }
    //        return s;
    //    }

    //    public bool ContainsLoop(Dictionary<TreeNode,TreeNode> visited, TreeNode from_tn)
    //    {
    //        // base case 1: loop check
    //        if (visited.ContainsKey(this))
    //        {
    //            return true;
    //        }
    //        // base case 2: an input
    //        if (!_com.IsFormula && _is_a_cell)
    //        {
    //            return false;
    //        }
    //        // recursive case
    //        bool OK = true;
    //        foreach (TreeNode t in this.getInputs())
    //        {
    //            if (OK)
    //            {
    //                // new dict to mark visit
    //                var visited2 = new Dictionary<TreeNode, TreeNode>(visited);
    //                // mark visit
    //                visited2.Add(this, from_tn);
    //                // recurse
    //                OK = OK && !t.ContainsLoop(visited2, this);
    //            }
    //        }
    //        return !OK;
    //    }

    //    public string getFormula()
    //    {
    //        return _com.Formula;
    //    }

    //    public Excel.Range getCOMObject()
    //    {
    //        return _com.Range;
    //    }

    //    public string getCOMValueAsString()
    //    {
    //        // null values become the empty string
    //        var s = System.Convert.ToString(this.getCOMObject().Value2);
    //        if (s == null)
    //        {
    //            return "";
    //        } else {
    //            return s;
    //        }
    //    }

    //    public bool isLeaf()
    //    {
    //        return _inputs.Count == 0;
    //    }
    }
}
