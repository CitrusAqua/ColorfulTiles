using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColorfulTiles
{
    class TextListConverter
    {
        public static List<string> ReadFile(string fileName)
        {
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            List<string> list = new List<string>();
            StreamReader sr = new StreamReader(fs);
            sr.BaseStream.Seek(0, SeekOrigin.Begin);
            string tmp = sr.ReadLine();
            while (tmp != null)
            {
                list.Add(tmp);
                tmp = sr.ReadLine();
            }
            sr.Close();
            fs.Close();
            return list;
        }

        public static void WriteFile(List<string> list, string txtFile)
        {
            FileStream fs = new FileStream(txtFile, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.Flush();
            sw.BaseStream.Seek(0, SeekOrigin.Begin);
            for (int i = 0; i < list.Count; i++) sw.WriteLine(list[i]);
            sw.Flush();
            sw.Close();
            fs.Close();
        }

        public static void AppendFile(List<string> list, string txtFile)
        {
            FileStream fs = new FileStream(txtFile, FileMode.Append, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.Flush();
            sw.BaseStream.Seek(0, SeekOrigin.Begin);
            for (int i = 0; i < list.Count; i++) sw.WriteLine(list[i]);
            sw.Flush();
            sw.Close();
            fs.Close();
        }
    }
}
