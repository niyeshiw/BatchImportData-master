using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Threading;
using Aspose.Cells;
using Aspose.Words;
using System.Data.SqlClient;

namespace BatchImportData
{
    class Program
    {
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
            Console.ReadLine();
            
        }

        private static void Import()
        {
            string filePathCase1 = Utils.GetSettings("FileDir:EN_Case_1");
            string filePathCase2 = Utils.GetSettings("FileDir:EN_Case_2");
            string filePathCase3 = Utils.GetSettings("FileDir:CN_Case_1");
            string filePathCase4 = Utils.GetSettings("FileDir:CN_Case_2");
            DirectoryInfo d1 = new DirectoryInfo(filePathCase1);
            DirectoryInfo d2 = new DirectoryInfo(filePathCase2);
            DirectoryInfo d3 = new DirectoryInfo(filePathCase3);
            DirectoryInfo d4 = new DirectoryInfo(filePathCase4);
            if (d1.GetFiles("*.xls").Length > 1 ||
                d2.GetFiles("*.xls").Length > 1 ||
                d3.GetFiles("*.xls").Length > 1 ||
                d4.GetFiles("*.xls").Length > 1)
            {
                Console.WriteLine("检测到存在历史文件,是否要执行导入到数据库操作? (Y/N)");
                string result = Console.ReadLine().Trim();
                if (!result.Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
                Console.WriteLine("确认执行导入操作? (Y/N)");
                result = Console.ReadLine().Trim();
                if (!result.Equals("y", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
                Console.WriteLine("正在导入...");
            }
            ImportExcel(d1, 0);
        }

        private static void ImportExcel(DirectoryInfo directoryInfo, int language)
        {
            var files = directoryInfo.GetFiles("*.xls");
            if (files.Length == 0)
            {
                return;
            }
            var file = files[0];

            Workbook workbook = new Workbook(file.FullName);
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;
            var firstCell = cells.Find("FileName", null);
            int rowIndex = firstCell.Row + 1;
            int colIndex = firstCell.Column;

            SqlConnection conn = new SqlConnection(Utils.GetSettings("ConnectionStr"));
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("导入失败,按任意键退出程序");
                Console.ReadLine();
                Environment.Exit(0);
            }
            var tran = conn.BeginTransaction();
            try
            {
                string preFilePath = null;
                byte[] sourceWord = null;

                for (; rowIndex < 10000; rowIndex++)
                {
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
                        sourceWord = new Document(filePath).ToByte();
                    }
                    Document conclusionWord = new Document(subFilePath);

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
                    int pid = (int)cmd.ExecuteScalar();

                    // 子表
                    cmd.CommandText = @"INSERT INTO [dbo].[Chemicals]([IsDeleted],[CreationTime],
[Casrn],[TiValue],[Details],[GroupName],[State],[PDE],[ChemicalName],[StudyReportId],[Word])
VALUES(0,GETDATE(),@Casrn, @TiValue,@Details, @GroupName, 0, @PDE, @ChemicalName, @PID, @Word);
";
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@Casrn", cASRN));
                    cmd.Parameters.Add(new SqlParameter("@TiValue", tiValue));
                    cmd.Parameters.Add(new SqlParameter("@Details", remark));
                    cmd.Parameters.Add(new SqlParameter("@GroupName", groupName));
                    cmd.Parameters.Add(new SqlParameter("@PDE", pDE));
                    cmd.Parameters.Add(new SqlParameter("@ChemicalName", chemicalName));
                    cmd.Parameters.Add(new SqlParameter("@PID", pid));
                    wordPara.Value = conclusionWord.ToByte();
                    cmd.Parameters.Add(new SqlParameter("@Word", wordPara));
                    cmd.ExecuteNonQuery();

                    // 引文
                    if (!string.IsNullOrEmpty(references))
                    {
                        foreach (var item in references.Split("\f"))
                        {

                        }
                    }
                    WordUtil.GenerateCitationFile(subFilePath);

                }
            }
            catch (Exception ex)
            {
                tran.Rollback();
                Console.WriteLine(ex.Message);
                Console.WriteLine("导入失败,按任意键退出程序");
                Console.ReadLine();
                Environment.Exit(0);
            }
            
            

            

        }
        private static void Clear()
        {
            var directory = Utils.GetSettings("FileDir:SplitFolder");
            DirectoryInfo directoryInfo = new DirectoryInfo(directory);
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
    }
}
