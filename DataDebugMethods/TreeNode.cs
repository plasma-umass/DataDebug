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
    /// For ranges, the name is of the format <EndCell>_to_<EndCell>, such as "A1_to_A5"
    /// For chart nodes, the name begins with the string "Chart", followed by the name of the chart object from Excel, with the white spaces stripped
    /// </summary>
    public class TreeNode
    {
        Excel.Workbook _workbook;
        private List<TreeNode> _parents;  //these are the TreeNodes that feed into the current cell
        private List<TreeNode> _children;    //these are the TreeNodes that the current cell feeds into
        private string _name;    //The name of each node: for cells, it is its address as a string; for ranges, it is of the form <EndCell>_to_<EndCell>; for charts it is "Chart<Name of chart>"
        private string _worksheet_name;  //This keeps track of the worksheet where this cell/range/chart is located
        private Excel.Worksheet _worksheet; // A reference to the actual worksheet where this TreeNode is located
        private double _weight;  //The weight of the node as computed by propagating values down the tree
        private bool _chart;
        private bool _is_formula; //this indicates whether this node is a formula
        private System.Drawing.Color originalColor;
        //private int originalColor;  //For using ColorIndex property instead of Color property
        private int colorBit = 0; 
        //Constructor method -- the string argument n is used as the name of the node; the string argument ws is used as the worksheet of the node
        public TreeNode(string n, Excel.Worksheet ws, Excel.Workbook wb)
        {
            _parents = new List<TreeNode>();
            _children = new List<TreeNode>();
            _name = n;
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
            _is_formula = false;
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
            return _name + "\nParents: " + parents_string + "\nChildren: " + children_string;
        }

        //Method for displaying a string representation of the node in GraphViz format
        public string toGVString(double max_weight)
        {
            string parents_string = "";
            foreach (TreeNode parent in _parents)
            {
                //parents_string += "\n" + parent.getWorksheet().Replace(" ", "") + "_" + parent.getName().Replace(" ", "") + "_weight_" + parent.getWeight() + "->" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "_weight_" + weight;
                parents_string += "\n" + parent.getWorksheet().Replace(" ", "") + "_" + parent.getName().Replace(" ", "") + "->" + _worksheet_name.Replace(" ", "") + "_" + _name.Replace(" ", "");
            }
            //string children_string = "";
            //foreach (TreeNode child in children)
            //{
            //    children_string += "\n" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "_weight:" + weight + "->" + child.getWorksheet().Replace(" ", "") + "_" + child.getName().Replace(" ", "") + "_weight:" + weight;
            //}
            //string weight_string = "\n" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "->iuc" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + " [style=dotted, arrowhead=odot, arrowsize=1] ; \niuc" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + " [shape=plaintext,label=\"Weight=" + weight + "\"]; \n{rank=same; " + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + ";iuc" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "}";

            //return ("\n" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "[shape = ellipse, fillcolor = \"0.000 " + (weight / max_weight) + " 0.878\", style = \"filled\"]"
            //return ("\n" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "_weight_" + weight + "[shape = ellipse]"
            //return ("\n" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "[label=\"\", shape = ellipse]"
            return ("\n" + _worksheet_name.Replace(" ", "") + "_" + _name.Replace(" ", "") + "[shape = ellipse]"
                //+ weight_string 
                + parents_string).Replace("$", "");
            //fillcolor = \"green\"   \"0.000 " + weight + " 0.878\"
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
        public void addParent(TreeNode node)
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
        public void addChild(TreeNode node)
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

        //By convention, we name ranges with the string "_to_" separating the end cells, such as "A1_to_A5"
        //If the name contains an underscore, and it is not a Chart node, then it is a Range node
        public bool isRange()
        {
            if (_name.Contains("_") && !isChart())
                return true;
            else
                return false;
        }

        //By convention, we add the string "Chart" to the beginning of the name of every Chart node
        public bool isChart()
        {
            return _chart;
            //if (name.Contains("Chart"))
            //    return true;
            //else
            //    return false;
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

        //Returns the original color of the cell as a System.Drawing.Color
        public System.Drawing.Color getOriginalColor()
        //public int getOriginalColor() //For using ColorIndex property instead of Color property
        {
            return originalColor;
        }

        //Sets the value of the original color of the cell as a System.Drawing.Color
        public void setOriginalColor(System.Drawing.Color color)
        //public void setOriginalColor(int color)   //For using ColorIndex property instead of Color property
        {
            //We only want to set the original color once
            if (colorBit == 0)
            {
                colorBit = 1; 
                originalColor = color;
            }
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

        /**
         * This is a recursive method for propagating the weights up the nodes in the tree.
         * It is used for weighting the outputs in the computation tree -- outputs with a lot of 
         * inputs have higher weight than ones that have fewer inputs. 
         * All inputs have weight 1 and their weights get passed up in the tree and accumulated at the outputs.
         * Modifies output_cells, reachable_grid, and reachable_impacts_grid
         */
        public static void propagateWeightUp(TreeNode node, double weight_passed_up, TreeNode originalNode, List<TreeNode> output_cells, bool[][][][] reachable_grid, List<double[]>[] reachable_impacts_grid)
        {
            if (!node.hasChildren())
            {
                int originalNode_row = originalNode.getWorksheetObject().Cells.get_Range(originalNode.getName()).Row - 1;
                int originalNode_col = originalNode.getWorksheetObject().Cells.get_Range(originalNode.getName()).Column - 1;
                //Mark that this output (node) is reachable from originalNode
                //Find node in output_cells
                for (int i = 0; i < output_cells.Count; i++)
                {
                    if (output_cells[i].getName().Equals(node.getName()) && output_cells[i].getWorksheet().Equals(node.getWorksheet()))
                    {
                        reachable_grid[originalNode.getWorksheetObject().Index - 1][originalNode_row][originalNode_col][i] = true;
                        reachable_impacts_grid[i].Add(new double[4] { (double)originalNode.getWorksheetObject().Index - 1, (double)originalNode_row, (double)originalNode_col, 0.0 });
                        //MessageBox.Show("Output " + i + " is reachable from " + originalNode.getWorksheetObject().Name + " " + originalNode.getWorksheetObject().Cells.get_Range(originalNode.getName()).get_Address().Replace("$",""));
                        break;
                    }
                }
                return;
            }
            else
            {
                foreach (TreeNode child in node.getChildren())
                {
                    child.setWeight(child.getWeight() + weight_passed_up);
                    propagateWeightUp(child, 1.0, originalNode, output_cells, reachable_grid, reachable_impacts_grid);
                }
            }
        }
    }
}
