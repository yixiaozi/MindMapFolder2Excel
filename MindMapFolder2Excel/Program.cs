using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Spire.Xls;

namespace MindMapFolder2Excel
{
    class Program
    {
        private static string[] noFiles = new string[] { };
        public static List<node> nodes = new List<node>();
        static void Main(string[] args)
        {
            //思维导图文件夹路径
            string mindMapFolderPath = @"E:\test";
            //Excel文件路径
            string excelFilePath = @"E:\test\test.xlsx";
            //设置args第一个参数为思维导图文件夹路径，第二个参数为Excel文件路径
            noFiles = "".Split(';');

            if (args.Length == 2)
            {
                mindMapFolderPath = args[0];
                excelFilePath = args[1];
            }
            //更新nodes
            GetAllNode(new DirectoryInfo(mindMapFolderPath));
            GetAllFiles(new DirectoryInfo(mindMapFolderPath));

            //使用Spire.XLS将nodes写入Excel
            Workbook workbook = new Workbook();
            //读取excelFilePath
            workbook.LoadFromFile(excelFilePath, ExcelVersion.Version2010);
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "思维导图";
            sheet.Range["A1"].Text = "思维导图";
            sheet.Range["A2"].Text = "文件名";
            sheet.Range["B2"].Text = "节点名称";
            sheet.Range["C2"].Text = "节点ID";
            sheet.Range["D2"].Text = "父节点ID";
            sheet.Range["E2"].Text = "父节点路径";
            sheet.Range["F2"].Text = "创建时间";
            sheet.Range["G2"].Text = "修改时间";
            sheet.Range["H2"].Text = "思维导图路径";
            sheet.Range["I2"].Text = "是否已删除";
            //获取C列所有的值到列表中
            List<string> list = new List<string>();
            for (int i = 1; i <= sheet.Rows.Length; i++)
            {
                list.Add(sheet.Range["C" + i].Text);
            }
            foreach (node n in nodes)
            {
                //如果C列包含ID则跳过
                if (list.Contains(n.IDinXML))
                {
                    //如果节点名称不一致则更新节点名称
                    if (sheet.Range["B" + (list.IndexOf(n.IDinXML) + 1)].Text != n.Text)
                    {
                        sheet.Range["B" + (list.IndexOf(n.IDinXML) + 1)].Text = n.Text;
                    }
                    continue;
                }
                //获取行数
                int rowCount = sheet.Rows.Length;
                int i=rowCount+1;
                //在最后一行插入
                sheet.InsertRow(i);
                sheet.Range["A" + i].Text = n.mindmapName;
                try
                {
                    sheet.Range["B" + i].Text = n.Text;
                }
                catch (Exception)
                {
                    sheet.Range["B" + i].Text = n.Text.Substring(30*1024);
                }
                sheet.Range["C" + i].Text = n.IDinXML;
                sheet.Range["D" + i].Text = n.ParentID;
                sheet.Range["E" + i].Text = n.ParentNodePath;
                sheet.Range["F" + i].Text = n.Time.ToString();
                sheet.Range["G" + i].Text = n.editDateTime.ToString();
                sheet.Range["H" + i].Text = n.mindmapPath;
                sheet.Range["I" + i].Text = "否";

            }
            //遍历所有行，如果C列不在nodes中将第8列设置为是
            for (int i = 3; i <= sheet.Rows.Length; i++)
            {
                if (!nodes.Exists(x => x.IDinXML == sheet.Range["C" + i].Text))
                {
                    sheet.Range["I" + i].Text = "是";
                }
            }
            //设置宽度自适应
            // sheet.AutoFitColumn(1, 8);


            //保存
            workbook.SaveToFile(excelFilePath, ExcelVersion.Version2010);
            
        }
        public static void GetAllFiles(DirectoryInfo dir)
        {
            node diritem = new node();
            diritem.mindmapName = dir.Name;
            diritem.IDinXML = GetMD5(dir.FullName);
            diritem.ParentID = GetMD5(dir.Parent.FullName);
            diritem.ParentNodePath = dir.Parent.FullName;
            diritem.Time = dir.CreationTime;
            diritem.editDateTime = dir.LastWriteTime;
            diritem.mindmapPath = dir.FullName;
            diritem.Text = dir.Name;
            nodes.Add(diritem);
            foreach (FileInfo file in dir.GetFiles())
            {
                //添加到nodes，节点名称是文件名，节点ID是文件名的MD5值，父节点ID是所在文件夹名的MD5值，父节点路径是文件夹路径，创建时间是文件创建时间，修改时间是文件修改时间，思维导图路径是文件路径
                if (file.Extension == ".mm")
                {
                    node item=new node();
                    item.mindmapName = file.Name;
                    item.IDinXML = GetMD5(file.FullName);
                    item.ParentID = GetMD5(dir.FullName);
                    item.ParentNodePath = dir.FullName;
                    item.Time = file.CreationTime;
                    item.editDateTime = file.LastWriteTime;
                    item.mindmapPath = file.FullName;
                    item.Text = file.Name;
                    nodes.Add(item);
                }
            }
            foreach (DirectoryInfo d in dir.GetDirectories())
            {
                GetAllFiles(d);
            }
        }
        public static string GetMD5(string sDataIn)
        {
            System.Security.Cryptography.MD5CryptoServiceProvider md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
            byte[] bytValue, bytHash;
            bytValue = System.Text.Encoding.UTF8.GetBytes(sDataIn);
            bytHash = md5.ComputeHash(bytValue);
            md5.Clear();
            string sTemp = "";
            for (int i = 0; i < bytHash.Length; i++)
            {
                sTemp += bytHash[i].ToString("X").PadLeft(2, '0');
            }
            return sTemp.ToLower();
        }

        public static string GetAttribute(XmlNode node, string name, int resultLenght = 0)
        {
            string resultdefault = "";
            for (int i = 0; i < resultLenght; i++)
            {
                resultdefault += " ";
            }
            try
            {
                if (node == null || node.Attributes == null || (name != "TEXT" && node.Attributes[name] == null))
                {
                    return resultdefault;
                }
                else if (node == null || node.Attributes == null || (name == "TEXT" && node.Attributes[name] == null))
                {
                    try
                    {
                        if (node.FirstChild.Name == "richcontent")
                        {
                            return new HtmlToString().StripHTML(node.FirstChild.InnerText);
                        }
                    }
                    catch (Exception)
                    {
                        return "未找到richcontent";
                    }
                }
                string result = "";
                result = node.Attributes[name].Value;
                result = FormatTimeLenght(result, resultLenght);
                return result;
            }
            catch (Exception)
            {
                return resultdefault;
            }
        }
        public static string FormatTimeLenght(string result, int resultLenght)
        {
            for (int i = result.Length; i < resultLenght; i++)
            {
                result = " " + result;
            }
            if (resultLenght != 0 && result.Trim() == "0")
            {
                result = result.Replace("0", " ");
            }
            return result;
        }
        public static string GetFatherNodeName(XmlNode node)
        {
            try
            {
                string s = "";
                while (node.ParentNode != null)
                {
                    try
                    {
                        //去掉根节点
                        if (node.ParentNode.ParentNode == null || node.ParentNode.ParentNode.Name == "map")
                        {
                            break;
                        }
                        s = (node.ParentNode.Attributes["TEXT"] != null ? node.ParentNode.Attributes["TEXT"].Value : "") + (s != "" ? ">" : "") + s;
                        node = node.ParentNode;
                    }
                    catch (Exception)
                    {
                        break;
                    }
                }
                return s;
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static void GetAllNode(DirectoryInfo path)
        {
            foreach (FileInfo file in path.GetFiles("*.mm", SearchOption.AllDirectories))
            {
                if (!noFiles.Contains(file.Name) && file.Name[0] != '~')
                {
                    try
                    {
                        System.Xml.XmlDocument x = new XmlDocument();
                        x.Load(file.FullName);
                        string fileName = file.Name.Substring(0, file.Name.Length - 3);
                        List<string> contents = new List<string>();
                        foreach (XmlNode node in x.GetElementsByTagName("node"))
                        {
                            try
                            {
                                if (node.Attributes["TEXT"] == null || node.Attributes["ID"] == null)
                                {
                                    continue;
                                }
                                if (node.Attributes["TEXT"].Value != "")
                                {
                                    //if (node.Attributes["TEXT"].Value.Length <= 4 && node.Attributes["TEXT"].Value.All(char.IsDigit))
                                    //{
                                    //    continue;
                                    //}
                                    string father = GetFatherNodeName(node);
                                    if (father.Contains("Folder|"))
                                    {
                                        continue;
                                    }
                                    if (!contents.Contains(node.Attributes["TEXT"].Value))
                                    {
                                        DateTime CREATEDdt = DateTime.MinValue;
                                        DateTime MODIFIEDdt = DateTime.MinValue;
                                        string CREATED = GetAttribute(node, "CREATED");
                                        string MODIFIED = GetAttribute(node, "MODIFIED");
                                        if (CREATED != "")
                                        {
                                            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
                                            long unixTimeStampCREATED = Convert.ToInt64(CREATED);
                                            CREATEDdt = startTime.AddMilliseconds(unixTimeStampCREATED);
                                            long unixTimeStampMODIFIED = Convert.ToInt64(MODIFIED);
                                            MODIFIEDdt = startTime.AddMilliseconds(unixTimeStampMODIFIED);
                                        }
                                        if (node.ParentNode.Attributes["ID"] == null)
                                        {
                                            nodes.Add(new node
                                            {
                                                Text = node.Attributes["TEXT"].Value,
                                                mindmapName = fileName,
                                                mindmapPath = file.FullName,
                                                editDateTime = MODIFIEDdt,
                                                Time = CREATEDdt,
                                                IDinXML = node.Attributes["ID"].Value,
                                                ParentNodePath = father,
                                                ParentID = GetMD5(file.FullName)
                                            });
                                        }
                                        else
                                        {
                                            nodes.Add(new node
                                            {
                                                Text = node.Attributes["TEXT"].Value,
                                                mindmapName = fileName,
                                                mindmapPath = file.FullName,
                                                editDateTime = MODIFIEDdt,
                                                Time = CREATEDdt,
                                                IDinXML = node.Attributes["ID"].Value,
                                                ParentNodePath = father,
                                                ParentID = node.ParentNode.Attributes["ID"].Value
                                            });
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        foreach (XmlNode node in x.GetElementsByTagName("richcontent"))
                        {
                            try
                            {
                                if (node.Attributes["TEXT"] == null)
                                {
                                    continue;
                                }
                                DateTime CREATEDdt = DateTime.Now;
                                DateTime MODIFIEDdt = DateTime.Now;
                                string CREATED = GetAttribute(node, "CREATED");
                                string MODIFIED = GetAttribute(node, "CREATED");
                                long unixTimeStampCREATED = Convert.ToInt64(CREATED);
                                long unixTimeStampMODIFIED = Convert.ToInt64(MODIFIED);
                                System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
                                CREATEDdt = startTime.AddMilliseconds(unixTimeStampCREATED);
                                MODIFIEDdt = startTime.AddMilliseconds(unixTimeStampMODIFIED);
                                nodes.Add(new node
                                {
                                    Text = node.InnerText,
                                    mindmapName = fileName,
                                    mindmapPath = file.FullName,
                                    editDateTime = MODIFIEDdt,
                                    Time = CREATEDdt,
                                    IDinXML = node.Attributes["ID"].Value,
                                    ParentID=node.ParentNode.Attributes["ID"].Value
                                });
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }
    }
    public class node
    {
        public string mindmapName { get; set; }
        public string mindmapPath { get; set; }
        public string Text { get; set; }
        public DateTime Time { get; set; }
        public DateTime editDateTime { get; set; }
        public string IDinXML { get; set; }
        public string ParentNodePath { get; set; }
        public string ParentID { get; set; }
    }
    public class HtmlToString
    {
        //https://www.codeproject.com/Articles/11902/Convert-HTML-to-Plain-Text-2
        public string StripHTML(string source)
        {
            try
            {
                string result;

                // Remove HTML Development formatting
                // Replace line breaks with space
                // because browsers inserts space
                result = source.Replace("\r", " ");
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\n", " ");
                // Remove step-formatting
                result = result.Replace("\t", string.Empty);
                // Remove repeating spaces because browsers ignore them
                result = System.Text.RegularExpressions.Regex.Replace(result,
                                                                      @"( )+", " ");

                // Remove the header (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*head([^>])*>", "<head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*head( )*>)", "</head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<head>).*(</head>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all scripts (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*script([^>])*>", "<script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*script( )*>)", "</script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                //result = System.Text.RegularExpressions.Regex.Replace(result,
                //         @"(<script>)([^(<script>\.</script>)])*(</script>)",
                //         string.Empty,
                //         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<script>).*(</script>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all styles (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*style([^>])*>", "<style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*style( )*>)", "</style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<style>).*(</style>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert tabs in spaces of <td> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*td([^>])*>", "\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line breaks in places of <BR> and <LI> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*br( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*li( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line paragraphs (double line breaks) in place
                // if <P>, <DIV> and <TR> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*div([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*tr([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*p([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Remove remaining tags like <a>, links, images,
                // comments etc - anything that's enclosed inside < >
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<[^>]*>", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // replace special characters:
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @" ", " ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&bull;", " * ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lsaquo;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&rsaquo;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&trade;", "(tm)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&frasl;", "/",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lt;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&gt;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&copy;", "(c)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&reg;", "(r)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove all others. More can be added, see
                // http://hotwired.lycos.com/webmonkey/reference/special_characters/
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&(.{2,6});", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // for testing
                //System.Text.RegularExpressions.Regex.Replace(result,
                //       this.txtRegex.Text,string.Empty,
                //       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // make line breaking consistent
                result = result.Replace("\n", "\r");

                // Remove extra line breaks and tabs:
                // replace over 2 breaks with 2 and over 4 tabs with 4.
                // Prepare first to remove any whitespaces in between
                // the escaped characters and remove redundant tabs in between line breaks
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\t)", "\t\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\r)", "\t\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\t)", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove redundant tabs
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove multiple tabs following a line break with just one tab
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Initial replacement target string for line breaks
                string breaks = "\r\r\r";
                // Initial replacement target string for tabs
                string tabs = "\t\t\t\t\t";
                for (int index = 0; index < result.Length; index++)
                {
                    result = result.Replace(breaks, "\r\r");
                    result = result.Replace(tabs, "\t\t\t\t");
                    breaks = breaks + "\r";
                    tabs = tabs + "\t";
                }

                // That's it.
                return result;
            }
            catch
            {
                return source;
            }
        }
    }
}
