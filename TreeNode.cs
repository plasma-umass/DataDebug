using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataDebug
{
    /// <summary>
    /// Data structure for representing nodes of the dependence graph (DAG) internally
    /// There are three types of nodes: normal cells, ranges, and charts
    /// The different types are distinguished by the name
    /// For normal cells, the name is just the address of the cell
    /// For ranges, the name is of the format <EndCell>_to_<EndCell>, such as "A1_to_A5"
    /// For chart nodes, the name begins with the string "Chart", followed by the name of the chart object from Excel, with the white spaces stripped
    /// </summary>
    class TreeNode
    {
        private List<TreeNode> parents;  //these are the TreeNodes that feed into the current cell
        private List<TreeNode> children;    //these are the TreeNodes that the current cell feeds into
        private string name;    //The name of each node: for cells, it is its address as a string; for ranges, it is of the form <EndCell>_to_<EndCell>; for charts it is "Chart<Name of chart>"
        private string worksheet;  //This keeps track of the worksheet where this cell/range/chart is located
        private double weight;  //The weight of the node as computed by propagating values down the tree
        private bool chart;
        private bool is_formula; //this indicates whether this node is a formula
        private System.Drawing.Color originalColor;
        //private int originalColor;  //For using ColorIndex property instead of Color property
        private int colorBit = 0; 
        //Constructor method -- the string argument n is used as the name of the node; the string argument ws is used as the worksheet of the node
        public TreeNode(string n, string ws)
        {
            parents = new List<TreeNode>();
            children = new List<TreeNode>();
            name = n;
            worksheet = ws;
            weight = 0.0;
            chart = false;
            is_formula = false;
        }

        public string toString()
        {
            string parents_string = "";
            foreach (TreeNode node in parents)
            {
                parents_string += node.getWorksheet() + " " + node.getName() + ", ";
            }
            string children_string = "";
            foreach (TreeNode node in children)
            {
                children_string += node.getName() + ", ";
            }
            return name + "\nParents: " + parents_string + "\nChildren: " + children_string;
        }

        //Method for displaying a string representation of the node in GraphViz format
        public string toGVString(double max_weight)
        {
            string parents_string = "";
            foreach (TreeNode parent in parents)
            {
                //parents_string += "\n" + parent.getWorksheet().Replace(" ", "") + "_" + parent.getName().Replace(" ", "") + "_weight_" + parent.getWeight() + "->" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "_weight_" + weight;
                parents_string += "\n" + parent.getWorksheet().Replace(" ", "") + "_" + parent.getName().Replace(" ", "") + "->" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "");
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
            return ("\n" + worksheet.Replace(" ", "") + "_" + name.Replace(" ", "") + "[shape = ellipse]"
                //+ weight_string 
                + parents_string).Replace("$", "");
            //fillcolor = \"green\"   \"0.000 " + weight + " 0.878\"
        }

        //Returns the name of the node
        public string getName()
        {
            return name;
        }

        //Returns the weight of this node
        public double getWeight()
        {
            return weight;
        }

        //Sets the weight of the node to the double passed as an argument
        public void setWeight(double w)
        {
            weight = w;
        }

        //Adds a parent to the list of parent nodes; checks for duplicates before adding it
        public void addParent(TreeNode node)
        {
            //Make sure we are not adding a parent more than once
            bool parent_already_added = false;
            foreach (TreeNode n in parents)
            {
                if (node.getName() == n.getName())
                    parent_already_added = true;
            }
            //If the parent is not on the list, add it
            if (!parent_already_added)
                parents.Add(node);
        }

        //Adds a child to the list of child nodes; checks for duplicates before adding it
        public void addChild(TreeNode node)
        {
            //Make sure we are not adding a child more than once
            bool child_already_added = false;
            foreach (TreeNode n in children)
            {
                if (node.getName() == n.getName())
                    child_already_added = true;
            }
            //If the child is not on the list, add it
            if (!child_already_added)
                children.Add(node);
        }

        //Checks if this node has any children
        public bool hasChildren()
        {
            if (children.Count == 0)
                return false;
            else
                return true;
        }

        //Checks if this node has any parents
        public bool hasParents()
        {
            if (parents.Count == 0)
                return false;
            else
                return true;
        }

        //By convention, we name ranges with the string "_to_" separating the end cells, such as "A1_to_A5"
        //If the name contains an underscore, and it is not a Chart node, then it is a Range node
        public bool isRange()
        {
            if (name.Contains("_") && !isChart())
                return true;
            else
                return false;
        }

        //By convention, we add the string "Chart" to the beginning of the name of every Chart node
        public bool isChart()
        {
            return chart;
            //if (name.Contains("Chart"))
            //    return true;
            //else
            //    return false;
        }

        public void setChart(bool value)
        {
            chart = value;
        }

        //Retuns the List<TreeNode> of children of this node
        public List<TreeNode> getChildren()
        {
            return children;
        }

        //Retuns the List<TreeNode> of parents of this node
        public List<TreeNode> getParents()
        {
            return parents;
        }

        //Returns the name of the worksheet that holds this cell/range/chart
        public string getWorksheet()
        {
            return worksheet;
        }

        public Microsoft.Office.Interop.Excel.Worksheet getWorksheetObject()
        {
            //Find worksheet of the TreeNode
            Microsoft.Office.Interop.Excel.Worksheet nodeWorksheet = null; //This will be the worksheet where the node n is located
            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in Globals.ThisAddIn.Application.Worksheets)
            {
                if (ws.Name == worksheet)
                {
                    nodeWorksheet = ws;
                    break;
                }
            }
            return nodeWorksheet;
        }
        //Sets the name of the worksheet that holds this cell/range/chart to the argument string s
        public void setWorksheet(string s)
        {
            worksheet = s;
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
            is_formula = true;
        }

        public bool isFormula()
        {
            if (is_formula == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
