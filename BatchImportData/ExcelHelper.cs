using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using Aspose.Cells;
using Aspose.Words;

namespace BatchImportData
{
    public class ExcelHelper
    {
        public static void ExportDataTable(Dictionary<string, string> header, DataTable dt, string filename)
        {
            Workbook workbook = new Workbook(); //工作簿
            Worksheet sheet = workbook.Worksheets[0]; //工作表
            Cells cells = sheet.Cells;//单元格
                                      //为标题设置样式
                                      //生成  列头行
            int headerNum = 0;//当前表头所在列
            Aspose.Cells.Style style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;  //设置居中
            style.Font.Size = 12;//文字大小
            style.Font.IsBold = true;//粗体
            style.BackgroundColor = Color.SkyBlue;
            style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = CellBorderType.Thin; //应用边界线 左边界线  
            style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = CellBorderType.Thin; //应用边界线 右边界线  
            style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = CellBorderType.Thin; //应用边界线 上边界线  
            style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = CellBorderType.Thin; //应用边界线 下边界线
            
            foreach (string item in header.Keys)
            {
                cells[1, headerNum].PutValue(item);
                cells[1, headerNum].SetStyle(style);
                cells.SetColumnWidthPixel(headerNum, 200);//设置单元格200宽度
                cells.SetRowHeight(1, 30);//第一行，30px高
                headerNum++;
            }
            Aspose.Cells.Style style2 = workbook.CreateStyle();
            style2.IsTextWrapped = true;
            //生成数据行
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cells.SetRowHeight(2 + i, 24);
                int contentNum = 0;//当前内容所在列
                foreach (string item in header.Keys)
                {
                    
                    string value = dt.Rows[i][header[item]] == null ? "" : dt.Rows[i][header[item]].ToString();
                    if (header[item] == "FileName"|| header[item] == "SubFilePath")
                    {
                        var index = sheet.Hyperlinks.Add(2+i, contentNum, 1, 1, value);
                        Hyperlink link = sheet.Hyperlinks[index];
                    }
                    else
                    {
                      
                        cells[2 + i, contentNum].SetStyle(style2);
                        cells[2 + i, contentNum].PutValue(value);
                    }
                  
                    contentNum++;
                }
            }

            // 对excel进行单元格合并
            int casrnIndex = 0, remarkIndex = 0;
            for (int j = 0; j < sheet.Cells.Columns.Count; j++)
            {
                if (sheet.Cells[1, j].Value.ToString() == "CASRN") casrnIndex = j;
                if (sheet.Cells[1, j].Value.ToString() == "Remark") remarkIndex = j;
            }
            int firstRow = 0;
            for (int i = 1; i < sheet.Cells.Rows.Count; i++)
            {
                if (sheet.Cells[i, casrnIndex].Value.ToString() == "NA") firstRow = i;

                if (firstRow != 0 && sheet.Cells[i, remarkIndex].Value.ToString() != string.Empty)
                {
                    int lastRow = i;
                    string data = sheet.Cells[i, remarkIndex].Value.ToString();

                    sheet.Cells.Merge(firstRow, remarkIndex, lastRow - firstRow + 1, 1);
                    sheet.Cells[firstRow, remarkIndex].PutValue(data);
                    firstRow = 0;
                }
            }
            workbook.Save(filename, Aspose.Cells.SaveFormat.Excel97To2003);
        }

    }
}
