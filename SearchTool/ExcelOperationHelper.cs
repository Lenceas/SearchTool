using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.IO;

namespace SearchTool
{
    /// <summary>
    /// Excel操作帮助类
    /// </summary>
    public class ExcelOperationHelper
    {
        /// <summary>
        /// Excel转换成DataTable
        /// </summary>
        /// <param name="hSSFWorkbook"></param>
        /// <returns></returns>
        public static List<DataTable> ToExcelDataTable(IWorkbook hSSFWorkbook)
        {
            List<DataTable> datatablelist = new List<DataTable>();
            for (int sheetIndex = 0; sheetIndex < hSSFWorkbook.NumberOfSheets; sheetIndex++)
            {
                ISheet sheet = hSSFWorkbook.GetSheetAt(sheetIndex);
                // 获取表头 FirstRowNum 第一行索引 0
                IRow header = sheet.GetRow(sheet.FirstRowNum);// 获取第一行
                if (header == null)
                {
                    break;
                }
                int startRow = 0;// 数据的第一行索引

                DataTable dtNpoi = new DataTable();
                startRow = sheet.FirstRowNum + 1;
                for (int i = header.FirstCellNum; i < header.LastCellNum; i++)
                {
                    ICell cell = header.GetCell(i);
                    if (cell != null)
                    {
                        string cellValue = $"{cell}";
                        if (cellValue != null)
                        {
                            DataColumn col = new DataColumn(cellValue);
                            dtNpoi.Columns.Add(col);
                        }
                        else
                        {
                            DataColumn col = new DataColumn();
                            dtNpoi.Columns.Add(col);
                        }
                    }
                }

                // 数据 LastRowNum 最后一行的索引 如第九行---索引8
                for (int i = startRow; i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);// 获取第i行
                    if (row == null)
                    {
                        continue;
                    }
                    DataRow dr = dtNpoi.NewRow();
                    // 遍历每行的单元格
                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        if (row.GetCell(j) != null)
                            dr[j] = row.GetCell(j).ToString();
                    }
                    dtNpoi.Rows.Add(dr);
                }
                dtNpoi.TableName = sheet.SheetName;
                datatablelist.Add(dtNpoi);
            }
            return datatablelist;
        }

        /// <summary>
        /// Excel转换成DataTable
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static List<DataTable> ExcelStreamToDataTable(Stream stream)
        {
            IWorkbook hSSFWorkbook = WorkbookFactory.Create(stream);
            return ToExcelDataTable(hSSFWorkbook);
        }
    }
}