using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace GenXmlTool
{
    public class NPOIRunner
    {
        public void Run(string xlsxPath, string cxmlPath, string ccsPath, string sxmlPath)
        {
            DirectoryInfo dir = new DirectoryInfo(xlsxPath);

            FileInfo[] allXlsx = dir.GetFiles("*.xlsx");
            if (allXlsx == null || allXlsx.Length == 0)
            {
                Console.WriteLine("未找到需要生成的Excel文件！！！");
                return;
            }

            foreach (FileInfo file in allXlsx)
                Excel2Xml(file, cxmlPath, sxmlPath, ccsPath);
        }

        private void Excel2Xml(FileInfo xlsxInfo, string cxmlPath, string sxmlPath, string ccsPath)
        {
            Console.WriteLine("开始处理:" + xlsxInfo.Name);
            try
            {
                IWorkbook wk = null;
                string p = xlsxInfo.FullName;//@"E:\GenXmlTools\GenXmlTool\GenXmlTool\bin\xlsx\skill.xlsx";
                FileStream fs = new FileStream(p, FileMode.Open, FileAccess.Read);

                if (p.IndexOf(".xlsx") > 0) // 2007版本  
                {
                    wk = new XSSFWorkbook(fs);  //xlsx数据读入workbook  
                }
                else if (p.IndexOf(".xls") > 0) // 2003版本  
                {
                    wk = new HSSFWorkbook(fs);  //xls数据读入workbook  
                }
                ISheet sheet = wk.GetSheetAt(0);


                int i = 0;
                StringBuilder cSb = GenSBHead();
                StringBuilder sSb = GenSBHead();

                string fieldValue;
                string flag;
                string fieldKey;
                IRow keyRow = sheet.GetRow(0);
                IRow flagRow = sheet.GetRow(2);
                IRow typeRow = sheet.GetRow(1);
                IRow row;
                int cols = flagRow.LastCellNum;
                XMLClientCSDef clientCSDef = new XMLClientCSDef();
                clientCSDef.mXmlCSName = sheet.SheetName + "Config";
                clientCSDef.mKeyType = typeRow.GetCell(1).ToString();
                clientCSDef.mKeyName = keyRow.GetCell(1).ToString();

                for (i = 5; i <= sheet.LastRowNum; i++)
                {
                    cSb.Append("     <item");
                    sSb.Append("     <item");
                    row = sheet.GetRow(i);
                    for (int j = 1; j < cols; j++)
                    {
                        flag = flagRow.GetCell(j).ToString();
                        fieldKey = keyRow.GetCell(j).ToString();
                        if (row.GetCell(j) != null)
                        {
                            fieldValue = System.Security.SecurityElement.Escape(row.GetCell(j).ToString());
                        }
                        else
                        {
                            if (typeRow.GetCell(j).ToString() == "string")
                                fieldValue = "";
                            else
                                fieldValue = "0";
                        }
                        if (flag.Contains("c|s"))
                        {
                            sSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                            cSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                            clientCSDef.AddAttr(typeRow.GetCell(j).ToString(), fieldKey);
                        }
                        else if (flag.Contains("c"))
                        {
                            cSb.Append(" " + fieldKey + "=\"" + fieldValue + "\"");
                            clientCSDef.AddAttr(typeRow.GetCell(j).ToString(), fieldKey);
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

                string xmlFiles = cxmlPath + "/" + sheet.SheetName + "Config.xml";
                GenStringBuilderData(xmlFiles, cSb);

                xmlFiles = sxmlPath + "/" + sheet.SheetName + ".xml";
                GenStringBuilderData(xmlFiles, sSb);

                string csPath = ccsPath + "/" + clientCSDef.mXmlCSName + ".cs";
                GenClientCSFile(clientCSDef, csPath);
                fs.Close();
                wk.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ex:" + ex);
            }
        }

        private StringBuilder GenSBHead()
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

        private void GenStringBuilderData(string files, StringBuilder sb)
        {
            ////写入文件
            using (FileStream fileStream = new FileStream(files, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.GetEncoding("utf-8")))
                {
                    textWriter.Write(sb.ToString());
                }
            }
        }

        private void GenClientCSFile(XMLClientCSDef def, string csPath)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("// Auto Generated Code\r\n");
            sb.Append("// Author roy\r\n");
            sb.AppendLine();
            sb.Append("using System.Collections.Generic;\r\n");
            sb.Append("using System.Xml;\r\n");
            sb.AppendLine();

            sb.AppendFormat("public class {0}\r\n{{", def.mXmlCSName);
            sb.AppendLine();
            bool blLanaguage = def.mXmlCSName.Contains("Language");
            string attriCodeStr = blLanaguage ? "\tpublic {0} {1} {{ get; set; }}" : "\tpublic {0} {1};";
            Dictionary<string, XMLClientCSAttrDef>.ValueCollection valColl = def.mDictAttrs.Values;
            foreach(XMLClientCSAttrDef attrDef in valColl) 
            {
                sb.AppendFormat(attriCodeStr, attrDef.mAttrType, attrDef.mAttrCode);
                sb.AppendLine();
            }
            sb.AppendLine();

            sb.AppendFormat("\tpublic static readonly string urlKey = \"{0}\";\r\n", def.mXmlCSName);
            sb.AppendFormat("\tstatic Dictionary<{0},{1}> AllDatas;\r\n", def.mKeyType, def.mXmlCSName);
            sb.AppendLine();

            sb.Append("\tpublic static void Parse(XmlNode node)\r\n");
            sb.Append("\t{\r\n");
            sb.AppendFormat("\t\tAllDatas = new Dictionary<{0},{1}>();\r\n", def.mKeyType, def.mXmlCSName);
            sb.Append("\t\tif (node != null)\r\n");
            sb.Append("\t\t{\r\n");
            sb.Append("\t\t\tXmlNodeList nodeList = node.ChildNodes;\r\n");
            sb.Append("\t\t\tif (nodeList != null && nodeList.Count > 0)\r\n");
            sb.Append("\t\t\t{\r\n");
            sb.Append("\t\t\t\tforeach (XmlElement el in nodeList)\r\n");
            sb.Append("\t\t\t\t{\r\n");

            sb.AppendFormat("\t\t\t\t\t{0} config = new {1}();\r\n", def.mXmlCSName, def.mXmlCSName);
            sb.AppendLine();

            foreach (XMLClientCSAttrDef attrDef in valColl)
            {
                if (blLanaguage)
                {
                    if (attrDef.mAttrType == "string")
                        sb.AppendFormat("\t\t\t\t\tconfig.{0} = el.GetAttribute (\"{1}\");\r\n", attrDef.mAttrCode, attrDef.mAttrCode);
                    else if (attrDef.mAttrType == "int")
                        sb.AppendFormat("\t\t\t\t\tconfig.{0} = int.Parse(el.GetAttribute (\"{1}\"));\r\n", attrDef.mAttrCode, attrDef.mAttrCode);
                    else if (attrDef.mAttrType == "float")
                        sb.AppendFormat("\t\t\t\t\tconfig.{0} = float.Parse(el.GetAttribute (\"{0}\"));\r\n", attrDef.mAttrCode, attrDef.mAttrCode);
                }
                else
                {

                    if (attrDef.mAttrType == "string")
                        sb.AppendFormat("\t\t\t\t\tconfig.{0} = el.GetAttribute (\"{1}\");\r\n", attrDef.mAttrCode, attrDef.mAttrCode);
                    else if (attrDef.mAttrType == "int")
                        sb.AppendFormat("\t\t\t\t\tint.TryParse(el.GetAttribute (\"{0}\"), out config.{1});\r\n", attrDef.mAttrCode, attrDef.mAttrCode);
                    else if (attrDef.mAttrType == "float")
                        sb.AppendFormat("\t\t\t\t\tfloat.TryParse(el.GetAttribute (\"{0}\"), out config.{1});\r\n", attrDef.mAttrCode, attrDef.mAttrCode);
                }
                sb.AppendLine();
            }

            sb.AppendFormat("\t\t\t\t\tAllDatas.Add(config.{0}, config);\r\n", def.mKeyName);

            sb.Append("\t\t\t\t}\r\n");
            sb.Append("\t\t\t}\r\n");
            sb.Append("\t\t}\r\n");
            sb.Append("\t}\r\n");
            sb.AppendLine();

            sb.AppendFormat("\tpublic static {0} Get({1} key)\r\n", def.mXmlCSName, def.mKeyType);
            sb.AppendFormat("\t{{\r\n");
            sb.Append("\t\tif (AllDatas != null && AllDatas.ContainsKey(key))\r\n");
            sb.Append("\t\t\treturn AllDatas[key];\r\n");
            sb.Append("\t\treturn null;\r\n");
            sb.Append("\t}\r\n");

            sb.AppendLine();
            sb.AppendFormat("\tpublic static Dictionary<{0},{1}> Get()\r\n", def.mKeyType, def.mXmlCSName);
            sb.AppendFormat("\t{{\r\n");
            sb.Append("\t\treturn AllDatas;\r\n");
            sb.Append("\t}\r\n");

            sb.Append("}");
            sb.AppendLine();

            GenStringBuilderData(csPath, sb);
        }
    }

    public class XMLClientCSDef
    {
        public string mXmlCSName;
        public string mKeyType;
        public string mKeyName;
        //public List<XMLClientCSAttrDef> mListAttrs;
        public Dictionary<string, XMLClientCSAttrDef> mDictAttrs;

        public XMLClientCSDef()
        {
            mDictAttrs = new Dictionary<string, XMLClientCSAttrDef>();
        }

        public void AddAttr(string type, string code)
        {
            if (mDictAttrs.ContainsKey(code))
                return;
            XMLClientCSAttrDef adef = new XMLClientCSAttrDef();
            adef.mAttrType = type;
            adef.mAttrCode = code;
            mDictAttrs.Add(adef.mAttrCode, adef);
        }
    }

    public class XMLClientCSAttrDef
    {
        public string mAttrType;
        public string mAttrCode;
    }
}
