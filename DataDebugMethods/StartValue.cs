using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataDebugMethods
{
    //This class is used for storing the starting values of cells during the fuzzing procedure.
    //Cells may contain numbers or strings, and we would like to be able to handle both. Thus we create a list of StartValue objects,
    //which can store either a number or a string.
    public class StartValue
    {
        private string string_value;  //The string content of an output cell
        private double double_value;  //The numeric content of an output cell

        //Generic constructor method
        public StartValue()
        {
            string_value = null;
            double_value = 0.0;
        }

        //Constructor which takes a string as a parameter, initializing string_value to the string parameter
        public StartValue(string s)
        {
            string_value = s;
            double_value = 0.0;
        }

        //Constructor which takes a double as a parameter, initializing double_value to the double parameter
        public StartValue(double d)
        {
            string_value = null;
            double_value = d;
        }

        //Constructor which takes a decimal as a parameter, initializing double_value to the double parameter
        public StartValue(decimal d)
        {
            string_value = null;
            double_value = (double)d;
        }

        //Getter method for the string_value field
        public string get_string()
        {
            return string_value;
        }

        //Getter method for the double_value field
        public double get_double()
        {
            return double_value;
        }

        //Setter method for the string_valeu field
        public void set_string(string s)
        {
            string_value = s;
        }

        //Setter method for the double_value field
        public void set_double(double d)
        {
            double_value = d;
        }
    }
}
