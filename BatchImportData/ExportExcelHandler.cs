using System;
using System.Collections.Generic;
using System.Data;

namespace BatchImportData
{
    public class ExportExcelHandler
    {
        /// <summary>
        /// 执行导出
        /// </summary>
        /// <param name="filePathCase"></param>
        /// <param name="name"></param>
        /// <param name="dt"></param>
        public static void ExecuteExportExcelData(string filePathCase,string name,DataTable dt)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add(nameof(ParagraphKeyModel.FileName), nameof(ParagraphKeyModel.FileName));
            dic.Add(nameof(ParagraphKeyModel.SubFilePath), nameof(ParagraphKeyModel.SubFilePath));
            dic.Add(nameof(ParagraphKeyModel.StudyNo), nameof(ParagraphKeyModel.StudyNo));
            dic.Add(nameof(ParagraphKeyModel.GroupName), nameof(ParagraphKeyModel.GroupName));
            dic.Add(nameof(ParagraphKeyModel.CASRN), nameof(ParagraphKeyModel.CASRN));
            dic.Add(nameof(ParagraphKeyModel.ChemicalName), nameof(ParagraphKeyModel.ChemicalName));
            dic.Add(nameof(ParagraphKeyModel.TiValue), nameof(ParagraphKeyModel.TiValue));
            dic.Add(nameof(ParagraphKeyModel.PDE), nameof(ParagraphKeyModel.PDE));
            dic.Add(nameof(ParagraphKeyModel.BodyContact), nameof(ParagraphKeyModel.BodyContact));
            dic.Add(nameof(ParagraphKeyModel.Duration), nameof(ParagraphKeyModel.Duration));
            dic.Add(nameof(ParagraphKeyModel.Population), nameof(ParagraphKeyModel.Population));   
            dic.Add("Remark", nameof(ParagraphKeyModel.DicValue));
            dic.Add(nameof(ParagraphKeyModel.References), nameof(ParagraphKeyModel.References));
            //导出ecel
            string path = filePathCase + "/" + name;
            ExcelHelper.ExportDataTable(dic, dt, path);
        }
    }
}
