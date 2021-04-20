using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Threading;
using Aspose.Cells;
using Aspose.Words;
using System.Data.SqlClient;
using System.Linq;

namespace BatchImportData
{
    class Program
    {
        public const string SplitFileName = "split"; 

        static void Main(string[] args)
        {
            Console.Clear();
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Green;
            //Console.WriteLine("开始读取文件夹中的文件...........");
            Console.WriteLine("Begin reading files from the folder...........");

            //注册序列号
            string license = "Aspose.Total.NET.lic";
            Aspose.Words.License license_word = new Aspose.Words.License();
            Aspose.Cells.License license_cells = new Aspose.Cells.License();
            license_word.SetLicense(license);
            license_cells.SetLicense(license);

            Import();
            Clear();
            TaskCase1();
            TaskCase2();
            TaskCase3();
            TaskCase4();

            //Console.WriteLine("程序运行结束...........");
            Console.WriteLine("End of program...........");
            Console.ReadKey();
        }

        private static void Import()
        {
            var allCase = GetAllCase().ToList();
            if (allCase.Any(p => p.GetFiles("*.xls").Length > 0))
            {
                Console.WriteLine("检测到存在历史文件,是否要执行导入到数据库操作? (Y/N)");
                string result = Console.ReadLine().Trim();
                if (!result.Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("放弃导入数据库操作");
                    return;
                }
                Console.WriteLine("确认执行导入操作? (Y/N)");
                result = Console.ReadLine().Trim();
                if (!result.Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("放弃导入数据库操作");
                    return;
                }
                Console.WriteLine("开始导入...");
                ImportExcel(allCase[0], 0);
                ImportExcel(allCase[1], 0);
                ImportExcel(allCase[2], 1);
                ImportExcel(allCase[3], 1);
                Console.WriteLine("导入完成,按Enter退出程序");
                Console.ReadLine();
                Environment.Exit(0);
            }
        }

        private static void ImportExcel(DirectoryInfo directoryInfo, int language)
        {
            Console.WriteLine("正在处理{0}", directoryInfo.Name);
            var files = directoryInfo.GetFiles("*.xls");
            if (files.Length == 0)
            {
                Console.WriteLine("{0}, 没有检测到excel文件", directoryInfo.Name);
                return;
            }
            var file = files[0];

            Workbook workbook = new Workbook(file.FullName);
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;
            var firstCell = cells.Find("FileName", null);
            int rowIndex = firstCell.Row + 1;
            int colIndex = firstCell.Column;
            int totalRowCount = cells.Rows.Count - firstCell.Row - 1;
            Console.WriteLine("{0}, 预计共{1}条数据", directoryInfo.Name, totalRowCount);
            SqlConnection conn = new SqlConnection(Utils.GetSettings("ConnectionStr"));
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("导入失败,按Enter退出程序");
                Console.ReadLine();
                Environment.Exit(0);
            }
            var tran = conn.BeginTransaction();
            try
            {
                string preFilePath = null;
                byte[] sourceWord = null;

                for (; !IsNullOrBlank(cells.GetRow(rowIndex)); rowIndex++)
                {
                    Console.WriteLine("{0}, 预计共{1}条数据,当前第{2}条...", directoryInfo.Name, totalRowCount, rowIndex - firstCell.Row);
                    string filePath = cells[rowIndex, colIndex].StringValue;
                    string subFilePath = cells[rowIndex, colIndex + 1].StringValue;
                    string studyNo = cells[rowIndex, colIndex + 2].StringValue;
                    string groupName = cells[rowIndex, colIndex + 3].StringValue;
                    string cASRN = cells[rowIndex, colIndex + 4].StringValue;
                    string chemicalName = cells[rowIndex, colIndex + 5].StringValue;
                    string tiValue = cells[rowIndex, colIndex + 6].StringValue;
                    string pDE = cells[rowIndex, colIndex + 7].StringValue;
                    string bodyContact = cells[rowIndex, colIndex + 8].StringValue;
                    string duration = cells[rowIndex, colIndex + 9].StringValue;
                    string population = cells[rowIndex, colIndex + 10].StringValue;
                    string remark = cells[rowIndex, colIndex + 11].StringValue;
                    string references = cells[rowIndex, colIndex + 12].StringValue;

                    if (preFilePath != filePath)
                    {
                        sourceWord = new Document(Path.Combine(directoryInfo.FullName,filePath)).ToByte();
                    }
                    Document conclusionWord = new Document(Path.Combine(directoryInfo.FullName, subFilePath));

                    // 主表
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.Transaction = tran;
                    cmd.CommandText = @"
IF NOT EXISTS(SELECT Id FROM StudyReports WHERE StudyNo = @StudyNo AND [Language] = @Language)
BEGIN
    INSERT INTO StudyReports([IsDeleted],[CreationTime],[StudyNo],[TypeBodyContact],[Duration],[Population],[Dates],[Language],[Word])
    VALUES(0,GETDATE(),@StudyNo,@TypeBodyContact, @Duration, @Population, GETDATE(), @Language, @Word)
    SELECT SCOPE_IDENTITY()
END
ELSE
    SELECT Id FROM StudyReports WHERE StudyNo = @StudyNo AND [Language] = @Language
";
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@StudyNo", studyNo));
                    cmd.Parameters.Add(new SqlParameter("@Language", language));
                    cmd.Parameters.Add(new SqlParameter("@TypeBodyContact", bodyContact));
                    cmd.Parameters.Add(new SqlParameter("@Duration", duration));
                    cmd.Parameters.Add(new SqlParameter("@Population", population));
                    var wordPara = new SqlParameter("@Word", SqlDbType.Image);
                    wordPara.Value = sourceWord;
                    cmd.Parameters.Add(wordPara);
                    decimal pid = Convert.ToDecimal(cmd.ExecuteScalar());

                    // 子表
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter($"SELECT Id FROM [dbo].[Chemicals] WHERE IsDeleted = 0 AND Casrn = @Casrn AND ChemicalName = @ChemicalName AND StudyReportId =@StudyReportId", conn);
                    sqlDataAdapter.SelectCommand.Transaction = tran;
                    sqlDataAdapter.SelectCommand.Parameters.Add(new SqlParameter("Casrn", cASRN));
                    sqlDataAdapter.SelectCommand.Parameters.Add(new SqlParameter("ChemicalName", chemicalName));
                    sqlDataAdapter.SelectCommand.Parameters.Add(new SqlParameter("StudyReportId", pid));
                    DataTable dt = new DataTable();
                    if (dt.Rows.Count > 0)
                    {
                        continue;
                    }

                    cmd.CommandText = @"INSERT INTO [dbo].[Chemicals]([IsDeleted],[CreationTime],
[Casrn],[TiValue],[Details],[GroupName],[State],[PDE],[ChemicalName],[StudyReportId],[Word])
VALUES(0,GETDATE(),@Casrn, @TiValue,@Details, @GroupName, 0, @PDE, @ChemicalName, @PID, @Word);
SELECT SCOPE_IDENTITY()
";
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@Casrn", cASRN));
                    cmd.Parameters.Add(new SqlParameter("@TiValue", tiValue));
                    cmd.Parameters.Add(new SqlParameter("@Details", remark));
                    cmd.Parameters.Add(new SqlParameter("@GroupName", groupName));
                    cmd.Parameters.Add(new SqlParameter("@PDE", pDE));
                    cmd.Parameters.Add(new SqlParameter("@ChemicalName", chemicalName));
                    cmd.Parameters.Add(new SqlParameter("@PID", pid));
                    wordPara = new SqlParameter("@Word", SqlDbType.Image);
                    cmd.Parameters.Add(wordPara);
                    wordPara.Value = conclusionWord.ToByte();

                    decimal cid = Convert.ToDecimal(cmd.ExecuteScalar());

                    var citations = conclusionWord.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    if (citations.Count > 0)
                    {
                        var strList = WordUtil.ExtractCitation(citations, conclusionWord);
                        if (strList.Count > 0) //引文
                        {
                            foreach (var item in strList)
                            {
                                string content = item.Trim();
                                sqlDataAdapter = new SqlDataAdapter($"SELECT Id FROM [dbo].[Citations] WHERE [CitationContent] =@CitationContent", conn);
                                sqlDataAdapter.SelectCommand.Transaction = tran;
                                sqlDataAdapter.SelectCommand.Parameters.Add(new SqlParameter("CitationContent", content));
                                dt = new DataTable();
                                sqlDataAdapter.Fill(dt);
                                decimal citationId = 0;
                                if (dt.Rows.Count == 1)
                                {
                                    citationId = Convert.ToDecimal(dt.Rows[0]["Id"]);
                                }
                                else // 插入引文
                                {
                                    // 生成引文
                                    var citationWord = WordUtil.GenerateCitationFile(item);
                                    cmd.Parameters.Clear();
                                    cmd.CommandText = @"
INSERT INTO [dbo].[Citations]([CitationContent],[Word])VALUES(@CitationContent,@Word)
SELECT SCOPE_IDENTITY()";
                                    cmd.Parameters.Add(new SqlParameter("@CitationContent", content));
                                    wordPara = new SqlParameter("@Word", SqlDbType.Image);
                                    cmd.Parameters.Add(wordPara);
                                    wordPara.Value = citationWord.ToByte();
                                    citationId = Convert.ToDecimal(cmd.ExecuteScalar());
                                }
                                //插入关系表
                                cmd.Parameters.Clear();
                                cmd.CommandText = "INSERT INTO [dbo].[ChemicalCitationMappings]([ChemicalId],[CitationId])VALUES(@ChemicalId,@CitationId)";
                                cmd.Parameters.Add(new SqlParameter("@ChemicalId", cid));
                                cmd.Parameters.Add(new SqlParameter("@CitationId", citationId));
                                int cnt = cmd.ExecuteNonQuery();
                            }
                        }

                    }
                }
                tran.Commit();
            }
            catch (Exception ex)
            {
                tran.Rollback();
                Console.WriteLine(ex.Message);
                Console.WriteLine("导入失败,按Enter退出程序");
                Console.ReadLine();
                Environment.Exit(0);
            }

        }

        private static void Clear()
        {
            foreach (var dirInfo in GetAllCase())
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(dirInfo.FullName, SplitFileName));
                if (!directoryInfo.Exists)
                {
                    continue;
                }
                foreach (var file in directoryInfo.GetFileSystemInfos())
                {
                    if (file is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(file.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        File.Delete(file.FullName);      //删除指定文件
                    }
                }
            }
        }

        private static IEnumerable<DirectoryInfo> GetAllCase()
        {
            string filePathCase1 = Utils.GetSettings("FileDir:EN_Case_1");
            string filePathCase2 = Utils.GetSettings("FileDir:EN_Case_2");
            string filePathCase3 = Utils.GetSettings("FileDir:CN_Case_1");
            string filePathCase4 = Utils.GetSettings("FileDir:CN_Case_2");
            yield return new DirectoryInfo(filePathCase1);
            yield return new DirectoryInfo(filePathCase2);
            yield return new DirectoryInfo(filePathCase3);
            yield return new DirectoryInfo(filePathCase4);
        }

        //执行任务1
        public static void TaskCase1()
        {
            //读取文件
            //string filePathCase1 = Directory.GetCurrentDirectory() + "/files/en_case1";
            string filePathCase1 = Utils.GetSettings("FileDir:EN_Case_1");
            //Console.WriteLine("读取【英文Case1】文件夹中的结论...........");
            Console.WriteLine("Read the conclusion in the [EN_Case1] folder...........");
            if (!Directory.Exists(filePathCase1))
            {
                //Console.WriteLine("不存在此文件夹，请提前放入...........");
                Console.WriteLine("This folder does not exist, please put it in advance...........");
            }
            //读取英文case1
            var listCase1 = BatchImportHelper.ImportCase1(filePathCase1);
            //Console.WriteLine($"读取完毕，获得记录总数：{listCase1.Count}...........");
            Console.WriteLine($"After reading, get the total number of records:{listCase1.Count}...........");
            if (listCase1.Count > 0)
            {
                //写入excel
                DataTable dt = Utils.ToDataTable(listCase1);
                ExportExcelHandler.ExecuteExportExcelData(filePathCase1, "ChemicaList1.xls", dt);
                //Console.WriteLine("写入excel成功...........");
                Console.WriteLine("Successfully write to Excel...........");
            }
        }
        /// <summary>
        /// 执行任务2
        /// </summary>
        public static void TaskCase2()
        {
            string filePathCase2 = Utils.GetSettings("FileDir:EN_Case_2");
            //Console.WriteLine("读取【英文Case2】文件夹中的结论...........");
            Console.WriteLine("Read the conclusion in the [EN_Case2] folder...........");
            if (!Directory.Exists(filePathCase2))
            {
                Console.WriteLine("This folder does not exist, please put it in advance...........");
            }
            //读取英文case2
            var listCase2 = BatchImportHelper.ImportCase2(filePathCase2);
            Console.WriteLine($"After reading, get the total number of records:{listCase2.Count}...........");
            if (listCase2.Count > 0)
            {
                //写入excel
                DataTable dt = Utils.ToDataTable(listCase2);
                ExportExcelHandler.ExecuteExportExcelData(filePathCase2, "ChemicaList2.xls", dt);
                Console.WriteLine("Successfully write to Excel...........");
            }
        }
        /// <summary>
        /// 执行任务3
        /// </summary>
        public static void TaskCase3()
        {
            string filePathCase3 = Utils.GetSettings("FileDir:CN_Case_1");
            Console.WriteLine("Read the conclusion in the [CN_Case1] folder...........");
            if (!Directory.Exists(filePathCase3))
            {
                Console.WriteLine("This folder does not exist, please put it in advance...........");
            }
            //读取中文case1
            var listCase3 = BatchImportHelper.ImportCase3(filePathCase3);
            Console.WriteLine($"After reading, get the total number of records:{listCase3.Count}...........");
            if (listCase3.Count > 0)
            {
                //写入excel
                DataTable dt = Utils.ToDataTable(listCase3);
                ExportExcelHandler.ExecuteExportExcelData(filePathCase3, "化合物集合1.xls", dt);
                Console.WriteLine("Successfully write to Excel...........");
            }
        }
        /// <summary>
        /// 执行任务4
        /// </summary>
        public static void TaskCase4()
        {
            string filePathCase4 = Utils.GetSettings("FileDir:CN_Case_2");
            Console.WriteLine("Read the conclusion in the [CN_Case2] folder...........");
            if (!Directory.Exists(filePathCase4))
            {
                Console.WriteLine("This folder does not exist, please put it in advance...........");
            }
            //读取中文case2
            var listCase4 = BatchImportHelper.ImportCase4(filePathCase4);
            Console.WriteLine($"After reading, get the total number of records:{listCase4.Count}...........");
            if (listCase4.Count > 0)
            {
                //写入excel
                DataTable dt = Utils.ToDataTable(listCase4);
                ExportExcelHandler.ExecuteExportExcelData(filePathCase4, "化合物集合2.xls", dt);
                Console.WriteLine("Successfully write to Excel...........");
            }
        }

        public static bool IsNullOrBlank(Row row)
        {
            return row == null || row.IsBlank;
        }
    }
}
