using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Classification = UserSimulation.Classification;
using OptChar = Microsoft.FSharp.Core.FSharpOption<char>;

namespace CheckCellTests
{
    [TestClass]
    public class ClassificationTests
    {
        [TestMethod]
        public void TestSerialize()
        {
            var classification = new Classification();
            var s = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory());
            var v = System.IO.Directory.GetParent(s.FullName).FullName;
            System.IO.Directory.CreateDirectory(v + "\\GeneratedFiles");
            var full_path = v + "\\GeneratedFiles\\testfile_foo.bin";
            classification.Serialize(full_path);
            
            var t = System.IO.File.OpenRead(full_path);
            t.Close();
        }

        [TestMethod]
        public void TestDeserialize()
        {
            var classification = new Classification();
            //set typo dictionary to explicit one
            Dictionary<Tuple<OptChar, string>, int> typo_dict = new Dictionary<Tuple<OptChar, string>, int>();
            var key = new Tuple<OptChar, string>(OptChar.Some('t'), "y");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('t'), "t");
            typo_dict.Add(key, 0);

            key = new Tuple<OptChar, string>(OptChar.Some('T'), "TT");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('e'), "e");
            typo_dict.Add(key, 1);

            key = new Tuple<OptChar, string>(OptChar.Some('s'), "s");
            typo_dict.Add(key, 1);
            classification.SetTypoDict(typo_dict);
            var s = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory());
            var v = System.IO.Directory.GetParent(s.FullName).FullName;
            System.IO.Directory.CreateDirectory(v + "\\GeneratedFiles");
            var full_path = v + "\\GeneratedFiles\\testfile.bin";
            classification.Serialize(full_path);

            Classification c2 = Classification.Deserialize(full_path);
            var typo_dict_2 = c2.GetTypoDict();
            Assert.AreEqual(typo_dict_2.Count, typo_dict.Count);
        }
    }
}
