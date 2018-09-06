using System;
using System.Data;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace GenXmlTool
{
    public class Runner
    {
        public void Run(string xlsxPath, string cxmlPath, string ccsPath, string sxmlPath)
        {
            DirectoryInfo dir = new DirectoryInfo(xlsxPath);
            Console.WriteLine("xlsx目录:" + dir.FullName);

            FileInfo[] allXlsx = dir.GetFiles("*.xlsx");
            if (allXlsx == null || allXlsx.Length == 0)
            {
                Console.WriteLine("未找到需要生成的Excel文件！！！");
                return;
            }

            foreach (FileInfo file in allXlsx)
                Excel2Xml(file, cxmlPath, sxmlPath);
        }

        private StringBuilder GenSB()
        {
            //创建一个StringBuilder存储数据
            StringBuilder sb = new StringBuilder();
            //创建Xml文件头
            sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.Append("\r\n");
            //创建根节点
            sb.Append("<Config>");
            sb.Append("\r\n");
            return sb;
        }

        private void GenXml(string xmlFiles, StringBuilder sb)
        {
            ////写入文件
            using (FileStream fileStream = new FileStream(xmlFiles, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.GetEncoding("utf-8")))
                {
                    textWriter.Write(sb.ToString());
                }
            }
        }

        private void Excel2Xml(FileInfo xlsxInfo, string cxmlPath, string sxmlPath)
        {
            Console.WriteLine("开始处理:" + xlsxInfo.Name);
            DateTime startT = DateTime.Now;
            Console.WriteLine("start time:" + startT);

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            Excel.Workbook wb = excelApp.Workbooks.Open(xlsxInfo.FullName, Missing.Value, true, Missing.Value, Missing.Value);

            Excel.Worksheet ws = null;
            try
            {
                ws = wb.Worksheets.Item[1];
                //FileStream fstream = File.Open(xlsxInfo.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
                //IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fstream);
                //DataSet excelDataSet = reader.AsDataSet();
                //if (excelDataSet.Tables.Count < 1)
                //{
                //    Console.WriteLine("Excel:{0},数据为空!", xlsxInfo.Name);
                //    return;
                //}

                int i = 0;
                //DataTable table = excelDataSet.Tables[0];
                //if (table.Rows.Count < 3)
                //{
                //    Console.WriteLine("Excel:{0},数据表名{1}数据格式不对!", xlsxInfo.Name, table.TableName);
                //    return;
                //}

                StringBuilder cSb = GenSB();
                StringBuilder sSb = GenSB();

                int rows = ws.UsedRange.Cells.Rows.Count;//table.Rows.Count;
                int cols = ws.UsedRange.Cells.Columns.Count;//table.Columns.Count;
                //for (int i = 0; i)
                //    Excel.Range rng = ws.Cells[5, 2];
                //Console.WriteLine(rng.Value2);
                string fieldValue;
                string flag;
                string fieldKey;
                //for (int col = 2; col <= cols; col++)
                //{
                //    flag = ws.Cells[2, col].Value2.ToString();//g[1, j] //.Rows[1][j].ToString();
                //    Console.WriteLine(flag);
                //}
                for (i = 5; i <= rows; i++)
                {
                    //if (string.IsNullOrEmpty(table.Rows[i][keyColIndex].ToString()))
                    //    break; //空数据行了
                    cSb.Append("     <item");
                    sSb.Append("     <item");
                    for (int j = 2; j <= cols; j++)
                    {
                        flag = ws.Cells[2, j].Value2.ToString();//g[1, j] //.Rows[1][j].ToString();
                        fieldKey = ws.Cells[1, j].Value2.ToString();//table.Rows[0][j].ToString();
                        if (ws.Cells[i, j].Value2 != null)
                            fieldValue = System.Security.SecurityElement.Escape(ws.Cells[i, j].Value2.ToString());//table.Rows[i][j].ToString());
                        else
                            fieldValue = "0";
                        if (flag.Contains("c|s"))
                        {
                            sSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                            cSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                        }
                        else if (flag.Contains("c"))
                        {
                            cSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                        }
                        else if (flag.Contains("s"))
                        {
                            sSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                        }
                    }
                    cSb.Append(" />");
                    cSb.Append("\n");
                    sSb.Append(" />");
                    sSb.Append("\n");
                }
                cSb.Append("</Config>");
                sSb.Append("</Config>");

                string xmlFiles = cxmlPath + "/" + ws.Name + ".xml";
                GenXml(xmlFiles, cSb);

                xmlFiles = sxmlPath + "/" + ws.Name + ".xml";
                GenXml(xmlFiles, sSb);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ex:" + ex);
            }
            wb.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excelApp);

            DateTime endT = DateTime.Now;
            Console.WriteLine("end time:" + endT);
            string log = string.Format("[{0}生成完成，消耗时间{1}毫秒]", xlsxInfo.Name, (endT - startT).ToString());
            Console.WriteLine(log);
            //System.Diagnostics.Debug.WriteLine(log);
            //Debuger.LogWarning(log);
        }
    }
}
