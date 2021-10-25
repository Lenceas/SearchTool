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
            try
            {
                List<DataTable> datatablelist = new List<DataTable>();
                for (int sheetIndex = 0; sheetIndex < hSSFWorkbook.NumberOfSheets; sheetIndex++)
                {
                    ISheet sheet = hSSFWorkbook.GetSheetAt(sheetIndex);
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

                    //初始化列头
                    DataTable dt = new DataTable();
                    dt.TableName = sheet.SheetName;
                    IRow row0 = sheet.GetRow(0);
                    if (row0 != null)
                    {
                        for (int i = 0; i < row0.LastCellNum; i++)
                        {
                            ICell cell = row0.GetCell(i);
                            if (cell == null)
                            {
                                dt.Columns.Add("cell" + i.ToString());
                            }
                            else
                            {
                                switch (i)
                                {
                                    case 0:
                                        dt.Columns.Add($"{nameof(ExcelModel.id)}");
                                        break;
                                    case 1:
                                        dt.Columns.Add($"{nameof(ExcelModel.type)}");
                                        break;
                                    case 2:
                                        dt.Columns.Add($"{nameof(ExcelModel.item)}");
                                        break;
                                    default:
                                        dt.Columns.Add(cell.ToString());
                                        break;
                                }
                            }
                        }
                    }

                    int colCount = dt.Columns.Count;
                    //获取行数据
                    int rowCount = 0;
                    while (rows.MoveNext())
                    {
                        if (rowCount > 0)
                        {
                            var row = (IRow)rows.Current;
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < colCount; i++)
                            {
                                ICell cell = row.GetCell(i);
                                if (cell == null)
                                {
                                    dr[i] = null;
                                }
                                else
                                {
                                    dr[i] = cell.ToString();
                                }
                            }
                            dt.Rows.Add(dr);
                        }
                        rowCount++;
                    }
                    datatablelist.Add(dt);
                }
                return datatablelist;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
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