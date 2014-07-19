using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ErrorDict = System.Collections.Generic.Dictionary<AST.Address, double>;

namespace UserSimulation
{
    [Serializable]
    public class UserResults
    {
        public List<AST.Address> true_positives = new List<AST.Address>();
        public List<AST.Address> false_positives = new List<AST.Address>();
        public HashSet<AST.Address> false_negatives = new HashSet<AST.Address>();
        //Keeps track of the largest errors we observe during the simulation for each output
        public ErrorDict max_errors = new ErrorDict();
        public List<double> current_total_error = new List<double>();
        public List<double> PrecRel_at_k = new List<double>();
    }
}
