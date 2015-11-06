using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;

namespace UserSimulation
{
    [Serializable]
    public class ErrorDB
    {
        public List<Error> Errors { get; set; }

        public ErrorDB()
        {
            Errors = new List<Error>();
        }

        public void AddError(int r, int c, string worksheet, string workbook, string path, string value)
        {
            var e = new Error();
            e.row = r;
            e.col = c;
            e.worksheet = worksheet;
            e.workbook = workbook;
            e.path = path;
            e.value = value;
            Errors.Add(e);
        }

        public void Serialize(string filename)
        {
            XmlSerializer x = new System.Xml.Serialization.XmlSerializer(this.GetType());
            using (StreamWriter sw = File.CreateText(filename))
            {
                x.Serialize(sw, this);
            }
        }

        public static ErrorDB Deserialize(string filename)
        {
            ErrorDB obj;
            XmlSerializer x = new System.Xml.Serialization.XmlSerializer(typeof(ErrorDB));
            using (StreamReader sr = new StreamReader(filename))
            {
                obj = (ErrorDB)x.Deserialize(sr);
            }
            return obj;
        }
    }

    [Serializable]
    public struct Error
    {
        public int row;
        public int col;
        public string path;
        public string workbook;
        public string worksheet;
        public string value;
        public AST.Address GetAddress()
        {
            return AST.Address.fromR1C1(row, col, worksheet, workbook, path);
        }
    }
}
