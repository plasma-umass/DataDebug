using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Drawing;

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
        public string ToCSVHeaderLine(string wbname)
        {
            return "wbname,job_id," + String.Join(",", Enumerable.Range(0, _cells.Length).Select(num => "cell" + num));
        }
        public string ToCSVLine(string wbname)
        {
            return wbname + "," + _job_id + "," + String.Join(",", _cells.Select(str => '"' + r.Replace(str, "\\\"") + '"'));
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
        private Bitmap[] ToImages()
        {
            var half = _cells.Length / 2;
            var str1 = String.Join(",", _cells.Take(half));
            var str2 = String.Join(",", _cells.Skip(half).Take(half));
            var output = new Bitmap[2];
            output[0] = DataDebugMethods.Utility.CreateBitmapImage(str1, 11);
            output[1] = DataDebugMethods.Utility.CreateBitmapImage(str2, 11);
            return output;
        }
        public void WriteAsImages(string path, string basename)
        {
            var bitmaps = this.ToImages();
            for (var i = 0; i < bitmaps.Length; i++)
            {
                var b = bitmaps[i];
                var filename = Path.Combine(path, basename + "_" + _job_id + "_" + (i + 1) + ".png");
                b.Save(filename);
            }
        }
    }
}
