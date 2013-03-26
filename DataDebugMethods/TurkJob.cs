using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace DataDebugMethods
{
    [Serializable]
    public class TurkJob
    {
        Regex r = new Regex("\"", RegexOptions.Compiled);

        private int _job_id;
        private string[] _cells;
        private string[] _addrs;

        public void SetJobId(int job_id) { _job_id = job_id; }
        public void SetCells(string[] cells) { _cells = cells; }
        public void SetAddrs(string[] addrs) { _addrs = addrs; }
        public string ToCSVLine()
        {
            return _job_id + "," + String.Join(",", _cells.Select(str => '"' + r.Replace(str, "\\\"") + '"')) + "\n";
        }
        public string GetValueAt(int index) { return _cells[index]; }
        public string GetAddrAt(int index) { return _addrs[index];  }
        public static void SerializeArray(string filename, TurkJob[] tjs)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None);
            formatter.Serialize(stream, tjs);
            stream.Close();
        }
        public static TurkJob[] DeserializeArray(string filename)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
            TurkJob[] obj = (TurkJob[])formatter.Deserialize(stream);
            stream.Close();
            return obj;
        }
    }
}
