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
        private double weight;  //The weight of the node as computed by propagating values down the tree

        //Constructor method -- the string argument is used as the name of the node
        public TreeNode(string n)
        {
            parents = new List<TreeNode>();
            children = new List<TreeNode>();
            name = n;
            weight = 0.0;
        }

        public string toString()
        {
            string parents_string = "";
            foreach (TreeNode node in parents)
            {
                parents_string += node.getName() + ", ";
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
            foreach (TreeNode node in parents)
            {
                parents_string += "\n" + node.getName() + "->" + name;
            }
            string children_string = "";
            foreach (TreeNode node in children)
            {
                children_string += "\n" + name + "->" + node.getName();
            }
            string weight_string = "\n" + name + "->iuc" + name + " [style=dotted, arrowhead=odot, arrowsize=1] ; \niuc" + name + " [shape=plaintext,label=\"Weight=" + weight + "\"]; \n{rank=same; " + name + ";iuc" + name + "}";

            return ("\n" + name + "[shape = ellipse, fillcolor = \"0.000 " + (weight / max_weight) + " 0.878\", style = \"filled\"]" + weight_string + parents_string + children_string).Replace("$", "");
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
            if (name.Contains("Chart"))
                return true;
            else
                return false;
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
    }
}
