using System;
using System.IO;

namespace GenXmlTool
{
    class Program
    {

        public static string _xlsxPath = "";
        private static string _cxmlPath = "";
        private static string _ccsPath = "";
        private static string _sxmlPath = "";
        static void Main(string[] args)
        {
            DateTime start = DateTime.Now;
            string cfgTxt = Environment.CurrentDirectory + "./path.txt";
            string path = File.ReadAllText(cfgTxt);
            ParsePath(path.Split('\n'));

            DirectoryInfo dir = new DirectoryInfo(_xlsxPath);

            FileInfo[] allXlsx = dir.GetFiles("*.xlsx");
            new NPOIRunner().Run(_xlsxPath, _cxmlPath, _ccsPath, _sxmlPath);

            Console.WriteLine("处理完成, 花费时间{0}, 按任意键退出...", (DateTime.Now - start));
            Console.ReadKey(true);
        }

        private static void ParsePath(string[] paths)
        {
            string tmppath = "";
            string[] p;
            for (int i = 0; i < paths.Length; i++)
            {
                tmppath = paths[i];
                if (string.IsNullOrEmpty(tmppath))
                    continue;
                if (!tmppath.Contains("|"))
                    continue;
                p = tmppath.Split('|');
                if (p.Length != 2)
                    continue;
                switch (p[0])
                {
                    case "xlsxpath":
                        _xlsxPath = Environment.CurrentDirectory + "/" + p[1].Replace("\r", "");
                        break;
                    case "sxmlpath":
                        _sxmlPath = Environment.CurrentDirectory + "/" + p[1].Replace("\r", "");
                        if (!Directory.Exists(_sxmlPath))
                            Directory.CreateDirectory(_sxmlPath);
                        break;
                    case "cxmlpath":
                        _cxmlPath = Environment.CurrentDirectory + "/" + p[1].Replace("\r", "");
                        if (!Directory.Exists(_cxmlPath))
                            Directory.CreateDirectory(_cxmlPath);
                        break;
                    case "ccspath":
                        _ccsPath = Environment.CurrentDirectory + "/" + p[1].Replace("\r", "");
                        if (!Directory.Exists(_ccsPath))
                            Directory.CreateDirectory(_ccsPath);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
