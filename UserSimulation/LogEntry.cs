using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UserSimulation
{
    [Serializable]
    public class LogEntry
    {
        readonly AnalysisType _procedure;
        readonly string _filename;
        readonly AST.Address _address;
        readonly string _original_value;
        readonly string _erroneous_value;
        readonly double _output_error_magnitude;
        readonly double _num_input_error_magnitude;
        readonly double _str_input_error_magnitude;
        readonly bool _was_flagged;
        readonly bool _was_error;
        readonly double _significance;
        readonly double _threshold;
        public LogEntry(AnalysisType procedure,
                        string filename,
                        AST.Address address,
                        string original_value,
                        string erroneous_value,
                        double output_error_magnitude,
                        double num_input_error_magnitude,
                        double str_input_error_magnitude,
                        bool was_flagged,
                        bool was_error,
                        double significance,
                        double threshold)
        {
            _filename = filename;
            _procedure = procedure;
            _address = address;
            _original_value = original_value;
            _erroneous_value = erroneous_value;
            _output_error_magnitude = output_error_magnitude;
            _num_input_error_magnitude = num_input_error_magnitude;
            _str_input_error_magnitude = str_input_error_magnitude;
            _was_flagged = was_flagged;
            _was_error = was_error;
            _significance = significance;
            _threshold = threshold;
        }

        public static String Headers()
        {
            return "filename, " + // 0
                   "procedure, " + // 1
                   "significance, " + // 2
                   "threshold, " + // 3
                   "address, " + // 4
                   "original_value, " + // 5
                   "erroneous_value," + // 6
                   "total_relative_error, " + // 7
                   "num_input_err_mag, " + // 8
                   "str_input_err_mag, " + // 9
                   "was_flagged, " + // 10
                   "was_error" + // 11
                   Environment.NewLine; // 12
        }

        public void WriteLog(String logfile)
        {
            if (!System.IO.File.Exists(logfile))
            {
                System.IO.File.AppendAllText(logfile, Headers());
            }
            System.IO.File.AppendAllText(logfile, String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11}{12}",
                                                        _filename, // 0
                                                        _procedure, // 1
                                                        _significance, // 2
                                                        _threshold, // 3
                                                        _address.A1Local(), // 4
                                                        _original_value, // 5
                                                        _erroneous_value, // 6
                                                        _output_error_magnitude,// 7
                                                        _num_input_error_magnitude, // 8
                                                        _str_input_error_magnitude, // 9
                                                        _was_flagged, // 10
                                                        _was_error, // 11
                                                        Environment.NewLine // 12
                                                        ));
        }
    }
}
