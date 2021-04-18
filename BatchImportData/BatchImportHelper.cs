using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

namespace BatchImportData
{
    public class BatchImportHelper
    {

        #region 英文模版
        /// <summary>
        /// 导出第一种案例
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<ParagraphKeyModel> ImportCase1(string filePath)
        {
            // 结果集
            List<ParagraphKeyModel> newDataList = new List<ParagraphKeyModel>();
            // 错误文件个数
            int errorcount = 0;
            // 案例1文件夹路径
            DirectoryInfo theFolder = new DirectoryInfo(filePath);
            FileInfo[] dirInfo = theFolder.GetFiles();
            // 遍历每份文件
            foreach (FileInfo file in dirInfo)
            {
                // 只处理Word文件
                if (!file.Extension.Contains(".doc") || !file.Extension.Contains(".docx"))
                {
                    continue;
                }

                try
                {
                    bool isSuccess; // 每步骤执行状态
                    string message; // 执行信息

                    LogFileHelper.WriteTextLog("开始处理文件：" + file.Name, file.Name);
                    Document doc = new Document(file.FullName);

                    #region Step 1: 读取表格编号
                    LogFileHelper.WriteTextLog("开始读取报告编号！", file.Name);

                    string StudyNo = string.Empty; // 表格编号
                    isSuccess = false;
                    Section sc = doc.FirstSection; //查找第一个section
                    if (sc != null)
                    {
                        NodeCollection parcList = sc.Body.GetChildNodes(NodeType.Paragraph, true);
                        if (parcList != null && parcList.Count > 0)
                        {
                            foreach (Paragraph item in parcList)
                            {
                                string data = Utils.GetReplaceMethod(item.GetText()).ToLower();
                                if (data == null)
                                    continue;
                                if (data.Contains("study no"))
                                {
                                    StudyNo = data.ToUpper()[9..];
                                }
                                if (data.Contains("report number"))
                                {
                                    if (item.NextParagraph() == null)
                                    {
                                        var nextNode = parcList.ElementAt(parcList.IndexOf(item) + 1);
                                        var no = Utils.GetReplaceMethod(nextNode.GetText());
                                        StudyNo = no;
                                    }
                                    else
                                    {
                                        var no = Utils.GetReplaceMethod(item.NextParagraph().GetText());
                                        StudyNo = no;
                                    }
                                }
                            }
                            if (StudyNo != string.Empty) isSuccess = true;
                        }
                    }
                    if (isSuccess)
                        LogFileHelper.WriteTextLog("成功读取报告编号【" + StudyNo + "】", file.Name);
                    else
                        LogFileHelper.WriteTextLog("报告编号读取失败！", file.Name);
                    #endregion

                    #region Step 2: 读取化合物共有特性
                    LogFileHelper.WriteTextLog("开始读取化合物共有特性！", file.Name);
                    isSuccess = false;
                    message = string.Empty;
                    string BodyContact = string.Empty;
                    string Duration = string.Empty;
                    string Population = string.Empty;
                    string deviceTableSearchKey = "Device Classification";
                    string typeOfBodyContact = "type of body contact";
                    string contactDuration = "contact duration";
                    string targetPopulation = "target population";
                    Paragraph DeviceTable = doc.FindParagraphByTitleName(deviceTableSearchKey);
                    if (DeviceTable != null)
                    {
                        Table tb = DeviceTable.NextTable(); //获取该段落后面的一个表格
                        if (tb != null)
                        {
                            foreach (Row row in tb.Rows)
                            {
                                foreach (Cell cell in row.Cells)
                                {
                                    string content = Utils.GetReplaceMethod(cell.GetText()).ToLower();
                                    if (content.Contains(typeOfBodyContact))
                                    {
                                        BodyContact = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    }
                                    if (content.Contains(contactDuration))
                                    {
                                        Duration = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    }
                                    if (content.Contains(targetPopulation))
                                    {
                                        Population = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    }
                                }
                                if (BodyContact != string.Empty && Duration != string.Empty && Population != string.Empty)
                                    isSuccess = true;
                                else
                                    message = "未查找到化合物特性！";
                            }
                        }
                        else
                            message = "未查找到化合物特性表格！";
                    }
                    else
                        message = "未查找到包含关键字【" + deviceTableSearchKey + "】的段落！";

                    if (isSuccess)
                        LogFileHelper.WriteTextLog("成功读取化合物特性！", file.Name);
                    else
                        LogFileHelper.WriteTextLog(message, file.Name);
                    #endregion

                    #region Step 3: 查找所有化合物表格
                    LogFileHelper.WriteTextLog("开始扫描所有化合物表格！", file.Name);
                    isSuccess = false;
                    message = string.Empty;
                    List<ParagraphKeyModel> listParas = new List<ParagraphKeyModel>(); // 表格数据的临时集合
                    string chemicalTableSearchKey_1 = "risk assessment of";
                    string chemicalTableSearchKey_2 = "tolerable intake and allowable limit levels";
                    string chemicalTableSearchKey_3 = "extractable chemicals by";

                    NodeCollection nodeList = doc.GetChildNodes(NodeType.Table, true);
                    LogFileHelper.WriteTextLog("全文共读取表格数：" + nodeList.Count(), file.Name);
                    foreach (Node item in nodeList)
                    {
                        Table table = (Table)item;
                        DataTable dt = new DataTable();
                        try
                        {
                            string body = Utils.GetReplaceMethod(item.GetText()).ToLower();
                            if (body.Contains(chemicalTableSearchKey_1) || body.Contains(chemicalTableSearchKey_2) || body.Contains(chemicalTableSearchKey_3))
                            {
                                char[] parp = new char[] { ' ' };
                                var words = body.Split(parp);
                                if (!words.Contains("table"))
                                    continue;
                                bool isFirstType;
                                // 情况1：表头本身是一个表格
                                if (table.Rows.Count() == 1)
                                {
                                    table = item.NextTable();
                                    isFirstType = true;
                                }
                                // 情况2：表头属于表格的一行,且为第一行第一列
                                else
                                {
                                    string title = Utils.GetReplaceMethod(table.FirstRow.GetText()).ToLower();
                                    if (title.Contains(chemicalTableSearchKey_1) || title.Contains(chemicalTableSearchKey_2) || title.Contains(chemicalTableSearchKey_3))
                                        isFirstType = false;
                                    else
                                        continue;
                                }

                                if (table != null)
                                {
                                    foreach (Row row in table.Rows)
                                    {
                                        var rowText = Utils.GetReplaceMethod(row.GetText());
                                        DataRow dataRow = dt.NewRow();
                                        int rowIndex = table.Rows.IndexOf(row);
                                        if (rowIndex == 0 && !isFirstType) // 如果标题属于表格的一行，则跳过第一行
                                            continue;

                                        if ((rowIndex == 0 && isFirstType) || (rowIndex == 1 && !isFirstType)) // 列标题
                                        {
                                            foreach (Cell cell in row.Cells)
                                            {
                                                string name = Utils.GetReplaceMethod(cell.GetText()).Replace(" ", "").Trim().ToLower();
                                                dt.Columns.Add(name);
                                            }
                                        }
                                        else
                                        {
                                            for (int i = 0; i < row.Cells.Count; i++)
                                            {
                                                Cell cell = row.Cells[i];
                                                string cellStr = cell.GetText().ToLower();
                                                string val = Utils.GetReplaceMethod(cell.GetText());
                                                // 判断分组
                                                bool isBold = cell.FirstParagraph.GetChildNodes(NodeType.Run, true).Cast<Run>().All(r => r.Font.Bold);
                                                if ((cellStr.Contains("na") || cellStr.Contains("n/a")) && isBold)
                                                    val += "-group";
                                                dataRow[i] = val;
                                            }
                                            dt.Rows.Add(dataRow);
                                        }
                                    }
                                }
                            }
                            // 表名为段落，直接读取表格判断
                            else if (body.Contains("cas") && (body.Contains("chemical name") || body.Contains("metal")) && body.Contains("ti"))
                            {
                                foreach (Row row in table.Rows)
                                {
                                    DataRow dataRow = dt.NewRow();
                                    int rowIndex = table.Rows.IndexOf(row);
                                    if (rowIndex == 0) // 列标题
                                    {
                                        foreach (Cell cell in row.Cells)
                                        {
                                            string name = Utils.GetReplaceMethod(cell.GetText()).Replace(" ", "").Trim().ToLower();
                                            dt.Columns.Add(name);
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 0; i < row.Cells.Count; i++)
                                        {
                                            Cell cell = row.Cells[i];
                                            string cellStr = cell.GetText().ToLower();
                                            string val = Utils.GetReplaceMethod(cell.GetText());
                                            // 判断分组
                                            bool isBold = cell.FirstParagraph.GetChildNodes(NodeType.Run, true).Cast<Run>().All(r => r.Font.Bold);
                                            if ((cellStr.Contains("na") || cellStr.Contains("n/a")) && isBold)
                                                val += "-group";
                                            dataRow[i] = val;
                                        }
                                        dt.Rows.Add(dataRow);
                                    }
                                }
                            }
                            // 从表格临时数据集中提取数据
                            if (dt != null)
                            {
                                foreach (DataRow dr in dt.Rows)
                                {
                                    ParagraphKeyModel model = new ParagraphKeyModel();
                                    model.BodyContact = BodyContact;
                                    model.Duration = Duration;
                                    model.Population = Population;

                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    {
                                        string key = dt.Columns[i].ColumnName.ToLower();
                                        string value = dr[key].ToString().Trim();
                                        // 分组判断
                                        if (value.Contains("-group"))
                                            model.GroupName = value;
                                        else
                                            model.GroupName = string.Empty;

                                        if (key.Contains("cas"))
                                            model.CASRN = value;
                                        if (key.Contains("chemicalname") || key.Contains("metal") || key.Contains("element"))
                                            model.ChemicalName = value;
                                        if (key.Contains("ti"))
                                            model.TiValue = value;
                                    }
                                    // 避免添加编号和名称重复的化合物
                                    var isExists = listParas.Where(t => !t.CASRN.Contains("-group") && t.CASRN == model.CASRN && t.ChemicalName == model.ChemicalName).Count();
                                    if (isExists == 0) listParas.Add(model);
                                }
                            }
                        }
                        catch(Exception e)
                        {
                            LogFileHelper.WriteTextLog("表格格式不支持，执行错误：" + e.Message.ToString(), file.Name);
                            continue; // 跳过当前表格
                        }
                    }
                    // 数据过滤
                    listParas = listParas.Where(p => !string.IsNullOrEmpty(p.CASRN) && !string.IsNullOrEmpty(p.ChemicalName)).ToList();
                    if (listParas.Count() > 0)
                        LogFileHelper.WriteTextLog("化合物读取成功！", file.Name);
                    #endregion

                    #region Step 4: 获取段落描述
                    LogFileHelper.WriteTextLog("开始读取化合物描述！", file.Name);
                    isSuccess = false;
                    message = string.Empty;
                    // 获取描述开始段落
                    List<string> paragraphStartKeys = new List<string>()
                    {
                        "REVIEW OF RELEVANT",
                        "material review and toxicological assessment",
                        "evaluation of extractables",
                        "risk assessment of extractable chemicals"
                    };
                    Paragraph paragraph = new Paragraph(doc);
                    foreach (var key in paragraphStartKeys)
                    {
                        paragraph = doc.FindParagraphByTitleName(key);
                        if (paragraph != null) break;
                    }
                    if (paragraph == null)
                    {
                        LogFileHelper.WriteTextLog("未获取描述开始段落，执行下一文件！", file.Name);
                        continue;
                    }
                    else
                    {
                        var paragraphText = Utils.GetReplaceMethod(paragraph.GetText());
                        LogFileHelper.WriteTextLog("读取描述开始段落：" + paragraphText, file.Name);
                    }

                    // 获取描述结束段落
                    List<string> paragraphEndKeys = new List<string>()
                    {
                        "DISCUSSION AND CONCLUSION", "discussion", "references"
                    };
                    Paragraph endParagrah = new Paragraph(doc);
                    foreach(var key in paragraphEndKeys)
                    {
                        endParagrah = doc.FindParagraphByTitleName(key);
                        if (endParagrah != null) break;
                    }
                    if (endParagrah == null)
                    {
                        LogFileHelper.WriteTextLog("未获取描述结束段落，执行下一文件！", file.Name);
                        continue;
                    }
                    else
                    {
                        var endParagrahText = Utils.GetReplaceMethod(endParagrah.GetText());
                        LogFileHelper.WriteTextLog("读取描述结束段落：" + endParagrahText, file.Name);
                    }
                    // 开始配对

                    string datakey = string.Empty;
                    var dataList = new List<ParagraphKeyModel>();
                    while (paragraph.GetText() != endParagrah.GetText())
                    {
                        string text = Utils.GetReplaceMethod(paragraph.GetText());
                        
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string RunHtml = string.Empty;
                        var runList = paragraph.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
                        if(runList.Count > 0 && runList.All(r => r.Font.Bold))
                        //if (paragraph.ParagraphFormat.Style.Font.Bold)  //判断是否粗体
                        {
                            bool flag = true;
                            foreach(var chemical in listParas)
                            {
                                if(text.Contains(chemical.CASRN) && text.Contains(chemical.ChemicalName) && text.Contains("CAS"))
                                {
                                    flag = false; break;
                                } // 遍历结束后，若flag==true，则该化合物未出现在表格中
                            }
                            if (flag)
                            {
                                model.DicKey = datakey;
                                model.DicValue = RunHtml;
                            }
                            else
                            {
                                model.DicKey = text;
                                model.DicValue = text;
                                datakey = text;
                            }
                        }
                        else
                        {
                            string runTitle = string.Empty;
                            NodeCollection nodelList = paragraph.GetChildNodes(NodeType.Run, true); //遍历每个段落下所有的run
                            foreach (Run item in nodelList)
                            {
                                if (item.Font.Bold) //如果run里面是加粗的
                                {
                                    runTitle += Utils.GetReplaceMethod(item.GetText());
                                }
                                else
                                {
                                    RunHtml += Utils.GetReplaceMethod(item.GetText());
                                }
                            }
                            if (!string.IsNullOrEmpty(runTitle))
                            {
                                datakey = runTitle;
                            }
                            model.DicKey = datakey;     //重复的key 
                            model.DicValue = RunHtml;   //把所有的run 值 拼接起来
                        }
                        model.Paragraph = paragraph; //保存段落的引用
                        dataList.Add(model);
                        // 获取下一段落
                        paragraph = paragraph.NextParagraph();
                        if (paragraph == null) break;
                    }

                    //dataList = dataList.Where(t => !string.IsNullOrEmpty(t.DicKey) && t.DicValue != "" && !t.DicValue.Contains("presented in Section")).ToList();
                    dataList = dataList.Where(t => !string.IsNullOrEmpty(t.DicKey) && t.DicValue != "").ToList();
                    dataList = dataList.Where(t => t.DicKey != t.DicValue).ToList();
                    if(dataList.Count() > 0)
                        LogFileHelper.WriteTextLog("成功读取化合物描述！", file.Name);
                    #endregion

                    #region Step 5: 获取报告引文 废弃
                    //LogFileHelper.WriteTextLog("开始读取报告引文！", file.Name);
                    //Paragraph referenceTitle = doc.FindParagraphByTitleName("references");
                    ////Paragraph referenceTitle = doc.FindReferenceTitle();
                    //if (referenceTitle != null)
                    //{
                    //    Paragraph nextReference = referenceTitle.NextParagraph();
                    //    string references = string.Empty;
                    //    while (nextReference != null)
                    //    {
                    //        if (nextReference.ParagraphFormat.OutlineLevel != OutlineLevel.BodyText
                    //            || nextReference.ParagraphFormat.Style.Font.Bold) break;

                    //        var linkRunList = nextReference.GetChildNodes(NodeType.Run, true);
                    //        string tempStr = string.Empty;
                    //        foreach(var run in linkRunList)
                    //        {
                    //            if (!run.GetText().Contains("HYPERLINK"))
                    //                tempStr += run.GetText();
                    //        }
                    //        references += Utils.GetReplaceMethod(tempStr) + "\n";
                    //        nextReference = nextReference.NextParagraph();
                    //    }
                    //    if(references != string.Empty)
                    //        LogFileHelper.WriteTextLog("成功读取报告引文！", file.Name);
                    //    else
                    //        LogFileHelper.WriteTextLog("未获取报告引文！", file.Name);
                    //    if (listParas.Count() > 0)
                    //    {
                    //        listParas[0].References = references;
                    //    }
                    //}
                    //else
                    //    LogFileHelper.WriteTextLog("未获取报告引文标题段落！", file.Name);
                    #endregion

                    #region Step 6: 整理已获取的数据
                    foreach (var item in listParas)
                    {
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string dicvalue = string.Empty;

                        // 化合物分组
                        if (item.CASRN != null)
                        {
                            if (!item.CASRN.Contains("-group"))
                            {
                                model.GroupName = item.GroupName;
                            }
                            else
                            {
                                item.CASRN = item.CASRN.Substring(0, item.CASRN.Length - 6);
                                model.GroupName = item.ChemicalName;
                            }   
                        }
                        // 基本信息
                        model.CASRN = item.CASRN;
                        model.ChemicalName = item.ChemicalName;
                        model.BodyContact = BodyContact;
                        model.TiValue = item.TiValue;
                        model.Duration = Duration;
                        model.Population = Population;
                        model.StudyNo = StudyNo;
                        model.FileName = file.Name;
                        model.FilePath = file.FullName;
                        model.References = item.References;
                        model.Language = 1;
                        // 结论描述
                        model.DicKey = item.ChemicalName + ", CASRN " + item.CASRN;
                        List<Paragraph> paragraphs = new List<Paragraph>();
                        foreach (var node in dataList)
                        {
                            var text = node.DicKey.Replace(" ", "").ToLower();
                            if (text.Contains("cas"))        //判断新集合里面是否包含 casrn值
                            {
                                string casrn = item.CASRN.Replace(" ", "").ToLower();
                                string chemicalName = item.ChemicalName.Replace(" ", "").ToLower();
                                if(casrn.Contains("notgiven") || casrn.Contains("n/a") || casrn.Contains("na"))
                                {
                                    if (text.Contains(chemicalName))
                                    { 
                                        dicvalue += node.DicValue;
                                        paragraphs.Add(node.Paragraph);
                                    }
                                }
                                else
                                {
                                    if (text.Contains(casrn) && text.Contains(chemicalName))
                                    {
                                        dicvalue += node.DicValue;
                                        paragraphs.Add(node.Paragraph);
                                    }
                                    else if (text.Contains(casrn))
                                    {
                                        dicvalue += node.DicValue;
                                        paragraphs.Add(node.Paragraph);
                                    }
                                }
                            }
                        }
                        if (dicvalue.Length > 50)
                        {
                            model.StarIndex = Utils.Left(dicvalue, 20);      //取段落中前20个字符
                            model.EndIndex = Utils.Right(dicvalue, 20);     //取段落中后20个字符
                        }
                        else
                        {
                            model.StarIndex = dicvalue;
                            model.EndIndex = dicvalue;
                        }
                        model.DicValue = dicvalue;
                        string folderName = Path.Combine(Utils.GetSettings("FileDir:SplitFolder"), model.FileName);
                        if (!Directory.Exists(folderName))
                        {
                            Directory.CreateDirectory(folderName);
                        }
                        model.SubFilePath = Path.Combine(folderName, Guid.NewGuid().ToString() + ".docx");
                        WordUtil.Split(paragraphs, model.SubFilePath);
                        var citations = WordUtil.ExtractCitation(paragraphs, doc);
                        if (citations != null)
                        {
                            model.References = string.Join("\r\f", citations);
                        }
                        WordUtil.GenerateCitationFile(model.References);
                        newDataList.Add(model);            //拼接到新的集合中，打印出来
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    errorcount++;
                    LogFileHelper.WriteTextLog(ex.Message, file.Name);
                    continue;
                }
            }
            if (errorcount > 0)
                Console.WriteLine($"{errorcount} files failed to execute");
            return newDataList;
        }

        /// <summary>
        /// 导入第二种案例
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<ParagraphKeyModel> ImportCase2(string filePath)
        {
            List<ParagraphKeyModel> newDataList = new List<ParagraphKeyModel>();
            int errcount = 0;
            DirectoryInfo theFolder = new DirectoryInfo(filePath);
            FileInfo[] dirInfo = theFolder.GetFiles();
            foreach (FileInfo file in dirInfo)
            {
                if (!file.Extension.Contains(".doc") || !file.Extension.Contains(".docx"))
                {
                    continue;
                }

                try
                {
                    Document doc = new Document(file.FullName);
                    string STUDYNO = string.Empty;
                    string BodyContact = string.Empty;
                    string Duration = string.Empty;
                    string Population = string.Empty;

                    //查找第一个section
                    Section sc = doc.FirstSection;
                    if (sc != null)
                    {
                        NodeCollection parcList = sc.Body.GetChildNodes(NodeType.Paragraph, true);
                        if (parcList != null && parcList.Count > 0)
                        {
                            foreach (Paragraph item in parcList)
                            {
                                string data = Utils.GetReplaceMethod(item.GetText());
                                if (data != null && data.ToLower().Contains("study no"))
                                {
                                    STUDYNO = data.ToUpper()[9..];
                                }
                                if (data != null && data.ToLower().Contains("report number"))
                                {
                                    if (item.NextParagraph() == null)
                                    {
                                        var nextNode = parcList.ElementAt(parcList.IndexOf(item) + 1);
                                        var no = Utils.GetReplaceMethod(nextNode.GetText());
                                        STUDYNO = no;
                                    }
                                    else
                                    {
                                        var no = Utils.GetReplaceMethod(item.NextParagraph().GetText());
                                        //STUDYNO = data.ToUpper() + " " + no;
                                        STUDYNO = no;
                                    }
                                }
                            }
                        }
                    }


                    NodeCollection nodeList = doc.GetChildNodes(NodeType.Table, true);
                    List<ParagraphKeyModel> listParas = new List<ParagraphKeyModel>();
                    foreach (Node item in nodeList)
                    {
                        string body = Utils.GetReplaceMethod(item.GetText());
                        //Regex reg = new Regex("PDE and Specification Limit for Acetophenone");
                        Regex reg = new Regex("PDE and Specification Limit for");
                        if (reg.IsMatch(body))
                        {
                            DataTable dt = new DataTable();
                            dt.Columns.Add("CASRN");
                            dt.Columns.Add("ChemicalName");
                            dt.Columns.Add("Adults");
                            Table table = item as Table;
                            if (table == null)
                            {
                                break;
                            }
                            foreach (Row row in table.Rows)
                            {
                                DataRow dataRow = dt.NewRow();
                                int rowIndex = table.Rows.IndexOf(row);
                                if (rowIndex > 2 && row.Cells.Count>3)
                                {
                                    string casrn= Utils.GetReplaceMethod(row.Cells[0].GetText());
                                    string chemical= Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    string adults= Utils.GetReplaceMethod(row.Cells[2].GetText());
                                    if(!string.IsNullOrEmpty(casrn) && !string.IsNullOrEmpty(chemical) && !string.IsNullOrEmpty(adults))
                                    {
                                        dataRow["CASRN"] = casrn;
                                        dataRow["ChemicalName"] = chemical;
                                        dataRow["Adults"] = adults;
                                        dt.Rows.Add(dataRow);
                                    }
                                }
                                
                            }
                            foreach (DataRow dr in dt.Rows)
                            {
                                ParagraphKeyModel model = new ParagraphKeyModel();
                                model.BodyContact = BodyContact;
                                model.Duration = Duration;
                                model.Population = Population;
                                model.CASRN = dr["CASRN"].ToString();
                                model.ChemicalName = dr["ChemicalName"].ToString();
                                //model.TiValue= dr["Adults"].ToString();
                                model.PDE = dr["Adults"].ToString();
                                //如果遇到相同的，就不添加
                                var isExists = listParas.Where(t => t.CASRN == model.CASRN && t.ChemicalName == model.ChemicalName).Count();
                                if(isExists == 0)
                                    listParas.Add(model);
                            }
                        }
                    }
                    string datakey = string.Empty;
                    var dataList = new List<ParagraphKeyModel>();
                    Paragraph paragraph = doc.FindParagraphByTitleName("REVIEW OF RELEVANT");
                    if (paragraph == null) paragraph = doc.FindParagraphByTitleName("material review and toxicological assessment");
                    if (paragraph == null) paragraph = doc.FindParagraphByTitleName("evaluation of extractables");
                    if (paragraph == null) continue;

                    Paragraph endParagrah = doc.FindParagraphByTitleName("DISCUSSION AND CONCLUSION");
                    if (endParagrah == null) endParagrah = doc.FindParagraphByTitleName("REFERENCES");
                    if (endParagrah == null) continue;

                    while (paragraph.GetText() != endParagrah.GetText())
                    {
                        string text = Utils.GetReplaceMethod(paragraph.GetText());
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string RunHtml = string.Empty;
                       
                        if (paragraph.ParagraphFormat.Style.Font.Bold)  //判断是否粗体
                        {
                            model.DicKey = text;
                            model.DicValue = text;
                            datakey = text;
                        }
                        else
                        {
                            string runTitle = string.Empty;
                            NodeCollection nodelList = paragraph.GetChildNodes(NodeType.Run, true); //遍历每个段落下所有的run
                            foreach (Run item in nodelList)
                            {
                                if (item.Font.Bold) //如果run里面是加粗的
                                {
                                    runTitle += Utils.GetReplaceMethod(item.GetText());
                                }
                                else
                                {
                                    RunHtml += Utils.GetReplaceMethod(item.GetText());
                                }
                            }
                            if (!string.IsNullOrEmpty(runTitle))
                                datakey = runTitle;

                            model.DicKey = datakey;      //重复的key 
                            model.DicValue = RunHtml;   //把所有的run 值 拼接起来
                        }
                        dataList.Add(model);
                        paragraph = paragraph.NextParagraph();
                        if (paragraph == null)
                            break;
                    }

                    dataList = dataList.Where(t => !string.IsNullOrEmpty(t.DicKey)).ToList();

                    #region 获取报告引文
                    Paragraph referenceTitle = doc.FindParagraphByTitleName("references");
                    if (referenceTitle != null)
                    {
                        Paragraph nextReference = referenceTitle.NextParagraph();
                        string references = string.Empty;
                        while (true)
                        {
                            if (nextReference == null || nextReference.ParagraphFormat.OutlineLevel == OutlineLevel.Level1)
                                break;

                            //references += Utils.GetReplaceMethod(nextReference.GetText());
                            references += nextReference.GetText();
                            nextReference = nextReference.NextParagraph();
                        }
                        //Console.WriteLine(references);
                        if (listParas.Count() > 0)
                        {
                            listParas[0].References = references;
                        }
                    }
                    #endregion
                    foreach (var item in listParas)
                    {
                        // string title = item.CASRN.Replace("/", "").Trim();
                        string title = (item.ChemicalName + ", CASRN " + item.CASRN).Replace(" ", "").ToLower();
                        string title1 = (item.ChemicalName + ", CAS# " + item.CASRN).Replace(" ", "").ToLower();

                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string dicvalue = string.Empty;
                        string casrn = item.CASRN;
                        model.DicKey = item.ChemicalName + ", CASRN " + item.CASRN; ;
                        model.BodyContact = BodyContact;
                        //model.TiValue = item.TiValue;
                        model.PDE = item.PDE;
                        model.Duration = Duration;
                        model.StudyNo = STUDYNO;
                        model.Population = Population;
                        model.Language = 1;
                        if (!title.ToUpper().Contains("NA"))
                            model.GroupName = string.Empty;
                        else
                            model.GroupName = item.ChemicalName;

                        foreach (var node in dataList)
                        {
                            var text = node.DicKey.Replace(" ", "").ToLower();
                            //if (key.Contains(title) || key.Contains(title1))        //判断新集合里面是否包含 casrn值
                            //{
                            //    dicvalue += node.DicValue;
                            //}

                            var keys = new List<string>()
                            {
                                item.ChemicalName.Replace(" ", "").ToLower(),
                                item.CASRN.Replace(" ", "").ToLower(),
                                ",cas"
                            };
                            if (Utils.IsContains(text, keys))        //判断新集合里面是否包含 casrn值
                            {
                                dicvalue += node.DicValue;
                            }
                        }
                        if (dicvalue.Length > 50)
                        {
                            model.StarIndex = Utils.Left(dicvalue, 20);      //取段落中前20个字符
                            model.EndIndex = Utils.Right(dicvalue, 20);     //取段落中后20个字符
                        }
                        else
                        {
                            model.StarIndex = dicvalue;
                            model.EndIndex = dicvalue;
                        }
                        model.FileName = file.Name;
                        model.TiValue = item.TiValue;
                        model.CASRN = item.CASRN;
                        model.ChemicalName = item.ChemicalName;
                        model.FilePath = file.FullName;
                        model.DicValue = dicvalue;
                        model.References = item.References;
                        
                        newDataList.Add(model);            //拼接到新的集合中，打印出来
                    }
                }
                catch (Exception ex)
                {
                    errcount++;
                    LogFileHelper.WriteTextLog(ex.Message, file.Name);
                    continue;
                }


            }

            if (errcount > 0)
                Console.WriteLine($"{errcount} files failed to execute");
            return newDataList;
        }

        #endregion


        #region 中文模版
        /// <summary>
        /// 第一种案例
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<ParagraphKeyModel> ImportCase3(string filePath)
        {
            List<ParagraphKeyModel> newDataList = new List<ParagraphKeyModel>();
            int errcount = 0;
            DirectoryInfo theFolder = new DirectoryInfo(filePath);
            FileInfo[] dirInfo = theFolder.GetFiles();
            foreach (FileInfo file in dirInfo)
            {
                if (!file.Extension.Contains(".doc") || !file.Extension.Contains(".docx"))
                {
                    continue;
                }


                try
                {
                    Document doc = new Document(file.FullName);
                    string BodyContact = string.Empty;
                    string Duration = string.Empty;
                    string Population = string.Empty;

                    #region 查找标题

                    Paragraph DeviceTable = doc.FindParagraphByTitleName("供试品分类");
                    if (DeviceTable != null)
                    {
                        Table tb = DeviceTable.NextTable(); //获取该段落后面的一个表格
                        if (tb != null)
                        {
                            foreach (Row row in tb.Rows)
                            {
                                foreach (Cell cell in row.Cells)
                                {
                                    string content = Utils.GetReplaceMethod(cell.GetText());
                                    if (content.Contains("接触类型"))
                                    {
                                        BodyContact = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    }
                                    if (content.Contains("接触时间"))
                                    {
                                        Duration = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    }
                                    if (content.Contains("目标人群"))
                                    {
                                        Population = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    }

                                }
                            }
                        }
                    }
                    string STUDYNO = string.Empty;
                    //查找第一个section
                    Section sc = doc.FirstSection;
                    if (sc != null)
                    {
                        NodeCollection parcList = sc.Body.GetChildNodes(NodeType.Paragraph, true);
                        if (parcList != null && parcList.Count > 0)
                        {
                            foreach (Paragraph item in parcList)
                            {
                                string data = Utils.GetReplaceMethod(item.GetText());
                                if (data != null && data.ToLower().Contains("试验编号"))
                                {
                                    STUDYNO = data.ToUpper()[4..];
                                }
                                if (data != null && data.ToLower().Contains("报告编号"))
                                {
                                    STUDYNO = data.ToUpper()[4..];
                                }
                            }
                        }
                    }
                    #endregion
                    NodeCollection nodeList = doc.GetChildNodes(NodeType.Table, true);
                    List<ParagraphKeyModel> listParas = new List<ParagraphKeyModel>();
                    foreach (Node item in nodeList)
                    {
                        string body = Utils.GetReplaceMethod(item.GetText());
                        if (body.Contains("可提取物的评估"))
                        {
                            DataTable dt = new DataTable();

                            Table table = item as Table;
                            if (table != null)
                            {
                                foreach (Row row in table.Rows)
                                {
                                    DataRow dataRow = dt.NewRow();
                                    int rowIndex = table.Rows.IndexOf(row);
                                    if (rowIndex == 1)
                                    {
                                        foreach (Cell cell in row.Cells)
                                        {
                                            string val = Utils.GetReplaceMethod(cell.GetText()).Trim();
                                            if (!string.IsNullOrEmpty(val))
                                            {
                                                dt.Columns.Add(val);
                                            }
                                        }
                                    }
                                    if (rowIndex > 1)
                                    {
                                        for (int i = 0; i < row.Cells.Count; i++)
                                        {
                                            string val = Utils.GetReplaceMethod(row.Cells[i].GetText()).Trim();
                                            dataRow[i] = val;
                                        }
                                        dt.Rows.Add(dataRow);
                                    }
                                }
                                foreach (DataRow dr in dt.Rows)
                                {
                                    ParagraphKeyModel model = new ParagraphKeyModel();
                                    model.BodyContact = BodyContact;
                                    model.Duration = Duration;
                                    model.Population = Population;
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    {
                                        string key = dt.Columns[i].ColumnName;
                                        string value = dr[key].ToString().Replace("/", "").Trim();
                                        if (!value.ToUpper().Contains("NA"))
                                            model.GroupName = string.Empty;
                                        else
                                            model.GroupName = value;

                                        if (key.Contains("CASRN"))
                                            model.CASRN = value;
                                        if (key.Contains("化合物名称"))
                                            model.ChemicalName = value;
                                        if (key.Contains("可耐受摄入量"))
                                            model.TiValue = value;
                                    }
                                    var isExists = listParas.Where(t => t.CASRN == model.CASRN && t.ChemicalName == model.ChemicalName).Count();
                                    if (isExists == 0)
                                        listParas.Add(model);
                                }
                            }
                        }
                    }

                    string datakey = string.Empty;
                    var dataList = new List<ParagraphKeyModel>();
                    Paragraph paragraph = doc.FindParagraphByTitleName("可提取化合物的相关毒理学信息综述及");

                    Paragraph endParagrah = doc.FindParagraphByTitleName("讨论和结论");
                    if (endParagrah == null)
                        endParagrah = doc.FindParagraphByTitleName("参考文献");

                    if (endParagrah == null)
                        continue;

                    while (paragraph.GetText() != endParagrah.GetText())
                    {
                        string text = Utils.GetReplaceMethod(paragraph.GetText());
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string RunHtml = string.Empty;
                        if (paragraph.ParagraphFormat.Style.Font.Bold)  //判断是否粗体
                        {
                            model.DicKey = text;
                            model.DicValue = text;
                            datakey = text;
                        }
                        else
                        {
                            string runTitle = string.Empty;
                            NodeCollection nodelList = paragraph.GetChildNodes(NodeType.Run, true); //遍历每个段落下所有的run
                            foreach (Run item in nodelList)
                            {
                                if (item.Font.Bold) //如果run里面是加粗的
                                {
                                    runTitle += Utils.GetReplaceMethod(item.GetText());
                                }
                                else
                                {
                                    RunHtml += Utils.GetReplaceMethod(item.GetText());
                                }
                            }
                            if (!string.IsNullOrEmpty(runTitle))
                            {
                                datakey = runTitle;
                            }
                            model.DicKey = datakey;      //重复的key 
                            model.DicValue = RunHtml;   //把所有的run 值 拼接起来
                        }
                        dataList.Add(model);
                        paragraph = paragraph.NextParagraph();

                    }
                    dataList = dataList.Where(t => !string.IsNullOrEmpty(t.DicKey)).ToList();
                    listParas = listParas.Where(t => !string.IsNullOrEmpty(t.CASRN)).ToList();

                    #region 获取报告引文
                    Paragraph referenceTitle = doc.FindParagraphByTitleName("参考文献");
                    if (referenceTitle != null)
                    {
                        Paragraph nextReference = referenceTitle.NextParagraph();
                        string references = string.Empty;
                        while (true)
                        {
                            if (nextReference == null || nextReference.ParagraphFormat.OutlineLevel == OutlineLevel.Level1)
                                break;

                            //references += Utils.GetReplaceMethod(nextReference.GetText());
                            references += nextReference.GetText();
                            nextReference = nextReference.NextParagraph();
                        }
                        //Console.WriteLine(references);
                        if (listParas.Count() > 0)
                        {
                            listParas[0].References = references;
                        }
                    }
                    #endregion
                    foreach (var item in listParas)
                    {
                        string title = item.CASRN.Trim();
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string dicvalue = string.Empty;
                        string casrn = item.CASRN;
                        model.DicKey = item.ChemicalName + ", CASRN " + item.CASRN;
                        model.BodyContact = BodyContact;
                        model.TiValue = item.TiValue;
                        model.Duration = Duration;
                        model.StudyNo = STUDYNO;
                        model.Language = 0;
                        model.Population = Population;
                        if (!title.ToUpper().Contains("NA"))
                            model.GroupName = string.Empty;
                        else
                            model.GroupName = item.ChemicalName;
                        foreach (var node in dataList)
                        {
                            if (node.DicKey.Contains(title))        //判断新集合里面是否包含 casrn值
                            {
                                dicvalue += node.DicValue;
                            }
                        }
                        if (dicvalue.Length > 50)
                        {
                            model.StarIndex = Utils.Left(dicvalue, 20);      //取段落中前20个字符
                            model.EndIndex = Utils.Right(dicvalue, 20);     //取段落中后20个字符
                        }
                        else
                        {
                            model.StarIndex = dicvalue;
                            model.EndIndex = dicvalue;
                        }
                        model.FileName = file.Name;
                        model.TiValue = item.TiValue;
                        model.CASRN = item.CASRN;
                        model.ChemicalName = item.ChemicalName;
                        model.FilePath = file.FullName;
                        model.DicValue = dicvalue;
                        model.References = item.References;
                        newDataList.Add(model);            //拼接到新的集合中，打印出来
                    }

                }
                catch (Exception ex)
                {
                    errcount++;
                    LogFileHelper.WriteTextLog(ex.Message, file.Name);
                    continue;
                }
            }
            if(errcount>0)
                Console.WriteLine($"{errcount} files failed to execute");
            return newDataList;
        }
        /// <summary>
        /// 第二种案例
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<ParagraphKeyModel> ImportCase4(string filePath)
        {
            List<ParagraphKeyModel> newDataList = new List<ParagraphKeyModel>();
            int errcount = 0;
            DirectoryInfo theFolder = new DirectoryInfo(filePath);
            FileInfo[] dirInfo = theFolder.GetFiles();
            foreach (FileInfo file in dirInfo)
            {
                if (!file.Extension.Contains(".doc") || !file.Extension.Contains(".docx"))
                {
                    continue;
                }

                try
                {
                    Document doc = new Document(file.FullName);
                    string BodyContact = string.Empty;
                    string Duration = string.Empty;
                    string Population = string.Empty;

                    string STUDYNO = string.Empty;
                    //查找第一个section
                    Section sc = doc.FirstSection;
                    if (sc != null)
                    {
                        NodeCollection parcList = sc.Body.GetChildNodes(NodeType.Paragraph, true);
                        if (parcList != null && parcList.Count > 0)
                        {
                            foreach (Paragraph item in parcList)
                            {
                                string data = Utils.GetReplaceMethod(item.GetText());
                                if (data != null && data.ToLower().Contains("试验编号"))
                                {
                                    STUDYNO = data.ToUpper()[4..];
                                }
                                if (data != null && data.ToLower().Contains("报告编号"))
                                {
                                    STUDYNO = data.ToUpper()[4..];
                                }
                            }
                        }
                    }

                    NodeCollection nodeList = doc.GetChildNodes(NodeType.Table, true);
                    List<ParagraphKeyModel> listParas = new List<ParagraphKeyModel>();
                    foreach (Node item in nodeList)
                    {
                        string body = Utils.GetReplaceMethod(item.GetText());
                        Regex reg = new Regex("PDE及化学分析限度");
                        if (reg.IsMatch(body))
                        {
                            DataTable dt = new DataTable();
                            dt.Columns.Add("CASRN");
                            dt.Columns.Add("化学品名称");
                            dt.Columns.Add("成人");
                            Table table = item as Table;
                            if (table == null)
                            {
                                break;
                            }
                            foreach (Row row in table.Rows)
                            {
                                DataRow dataRow = dt.NewRow();
                                int rowIndex = table.Rows.IndexOf(row);
                                if (rowIndex > 2 && row.Cells.Count > 3)
                                {
                                    string casrn = Utils.GetReplaceMethod(row.Cells[0].GetText());
                                    string chemical = Utils.GetReplaceMethod(row.Cells[1].GetText());
                                    string adults = Utils.GetReplaceMethod(row.Cells[2].GetText());
                                    if (!string.IsNullOrEmpty(casrn) && !string.IsNullOrEmpty(chemical) && !string.IsNullOrEmpty(adults))
                                    {
                                        dataRow["CASRN"] = casrn;
                                        dataRow["化学品名称"] = chemical;
                                        dataRow["成人"] = adults;
                                        dt.Rows.Add(dataRow);
                                    }
                                }

                            }
                            foreach (DataRow dr in dt.Rows)
                            {
                                ParagraphKeyModel model = new ParagraphKeyModel();
                                model.BodyContact = BodyContact;
                                model.Duration = Duration;
                                model.Population = Population;
                                model.CASRN = dr["CASRN"].ToString();
                                model.ChemicalName = dr["化学品名称"].ToString();
                                //model.TiValue = dr["成人"].ToString();
                                model.PDE = dr["成人"].ToString();
                                var isExists = listParas.Where(t => t.CASRN == model.CASRN && t.ChemicalName == model.ChemicalName).Count();
                                if(isExists==0)
                                    listParas.Add(model);
                            }
                        }
                    }
                    string datakey = string.Empty;
                    var dataList = new List<ParagraphKeyModel>();
                    Paragraph paragraph = doc.FindParagraphByTitleName("毒理学信息相关综述");

                    Paragraph endParagrah = doc.FindParagraphByTitleName("讨论和结论");
                    if (endParagrah == null)
                        endParagrah = doc.FindParagraphByTitleName("参考文献");

                    if (endParagrah == null)
                        continue;

                    while (paragraph.GetText() != endParagrah.GetText())
                    {
                        string text = Utils.GetReplaceMethod(paragraph.GetText());
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string RunHtml = string.Empty;

                        if (paragraph.ParagraphFormat.Style.Font.Bold)  //判断是否粗体
                        {
                            model.DicKey = text;
                            model.DicValue = text;
                            datakey = text;
                        }
                        else
                        {
                            string runTitle = string.Empty;
                            NodeCollection nodelList = paragraph.GetChildNodes(NodeType.Run, true); //遍历每个段落下所有的run
                            foreach (Run item in nodelList)
                            {
                                if (item.Font.Bold) //如果run里面是加粗的
                                {
                                    runTitle += Utils.GetReplaceMethod(item.GetText());
                                }
                                else
                                {
                                    RunHtml += Utils.GetReplaceMethod(item.GetText());
                                }
                            }
                            if (!string.IsNullOrEmpty(runTitle))
                                datakey = runTitle;

                            model.DicKey = datakey;      //重复的key 
                            model.DicValue = RunHtml;   //把所有的run 值 拼接起来
                        }
                        dataList.Add(model);
                        paragraph = paragraph.NextParagraph();

                    }

                    dataList = dataList.Where(t => !string.IsNullOrEmpty(t.DicKey)).ToList();

                    #region 获取报告引文
                    Paragraph referenceTitle = doc.FindParagraphByTitleName("参考文献");
                    if (referenceTitle != null)
                    {
                        Paragraph nextReference = referenceTitle.NextParagraph();
                        string references = string.Empty;
                        while (true)
                        {
                            if (nextReference == null || nextReference.ParagraphFormat.OutlineLevel == OutlineLevel.Level1)
                                break;

                            //references += Utils.GetReplaceMethod(nextReference.GetText());
                            references += nextReference.GetText();
                            nextReference = nextReference.NextParagraph();
                        }
                        //Console.WriteLine(references);
                        if (listParas.Count() > 0)
                        {
                            listParas[0].References = references;
                        }
                    }
                    #endregion
                    foreach (var item in listParas)
                    {
                        string title =item.CASRN==null?"": item.CASRN.Replace("/", "").Trim();
                        ParagraphKeyModel model = new ParagraphKeyModel();
                        string dicvalue = string.Empty;
                        string casrn = item.CASRN;
                        model.DicKey = item.ChemicalName + ", CASRN " + item.CASRN;
                        model.BodyContact = BodyContact;
                        //model.TiValue = item.TiValue;
                        model.PDE = item.PDE;
                        model.Duration = Duration;
                        model.Population = Population;
                        model.StudyNo = STUDYNO;
                        model.Language = 0;
                        if (!title.ToUpper().Contains("NA"))
                            model.GroupName = item.GroupName;
                        else
                            model.GroupName = item.ChemicalName;

                        foreach (var node in dataList)
                        {
                            if (node.DicKey.Contains(title))        //判断新集合里面是否包含 casrn值
                            {
                                dicvalue += node.DicValue;
                            }
                        }
                        if (dicvalue.Length > 50)
                        {
                            model.StarIndex = Utils.Left(dicvalue, 20);      //取段落中前20个字符
                            model.EndIndex = Utils.Right(dicvalue, 20);     //取段落中后20个字符
                        }
                        else
                        {
                            model.StarIndex = dicvalue;
                            model.EndIndex = dicvalue;
                        }
                        model.FileName = file.Name;
                        //model.TiValue = item.TiValue;
                        model.PDE = item.PDE;
                        model.CASRN = item.CASRN;
                        model.ChemicalName = item.ChemicalName;
                        model.FilePath = file.FullName;
                        model.DicValue = dicvalue;
                        model.References = item.References;
                        newDataList.Add(model);            //拼接到新的集合中，打印出来
                    }

                }
                catch (Exception ex)
                {
                    errcount++;
                    LogFileHelper.WriteTextLog(ex.Message, file.Name);
                    continue;
                }


            }

            if (errcount > 0)
                Console.WriteLine($"{errcount} files failed to execute");
            return newDataList;
        }

        #endregion
    }

}
