using System;
using System.Collections.Generic;
using Aspose.Words;

namespace BatchImportData
{
    public class ParagraphKeyModel
    {
        
        /// <summary>
        /// StudyNo
        /// </summary>
        public string StudyNo { get; set; }
        /// <summary>
        /// 文件名称
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// 分组名称
        /// </summary>
        public string GroupName { get; set; }
        /// <summary>
        /// 段落的Key
        /// </summary>
        public string DicKey { get; set; }
        /// <summary>
        /// 段落开始索引
        /// </summary>
        public string StarIndex { get; set; }
        /// <summary>
        /// 段落结束索引
        /// </summary>
        public string EndIndex { get; set; }
        /// <summary>
        /// BodyContact
        /// </summary>
        public string BodyContact { get; set; }
        /// <summary>
        /// Duration
        /// </summary>
        public string Duration { get; set; }
        /// <summary>
        /// Population
        /// </summary>
        public string Population { get; set; }
        /// <summary>
        /// CARSN
        /// </summary>
        public string CASRN { get; set; }
        /// <summary>
        /// 化合物名称
        /// </summary>
        public string ChemicalName { get; set; }
        /// <summary>
        /// Ti值
        /// </summary>
        public string TiValue { get; set; }
        /// <summary>
        /// PDE
        /// </summary>
        public string PDE { get; set; }
        /// <summary>
        /// 段落的值
        /// </summary>
        public string DicValue { get; set; }
        /// <summary>
        /// 源文件连接
        /// </summary>
        public string FilePath { get; set; }
        /// <summary>
        /// 语言
        /// </summary>
        public int Language { get; set; }
        /// <summary>
        /// 引文
        /// </summary>
        public string References { get; set; }

        public Paragraph Paragraph { get; set; }
        public string SubFilePath { get; set; }
    }
}
