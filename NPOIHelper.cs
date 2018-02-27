using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.POIFS;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.SS;
using NPOI.DDF;
using NPOI.SS.Util;
using System.Collections;
using System.Text.RegularExpressions;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;
using System.Windows.Forms;

public class NPOIHelper
{
    #region 在用函数，属性等，经过测试

    #region 样式共通
    static Int32 MyColWidth = 15;
    static short MyColRowHeight = 15 * 20;
    static short MyHeaderRowHeight = 18 * 20;

    static void SetCellValue(string drValue, String dataType, ICellStyle dateStyle, ref ICell newCell)
    {
        switch (dataType)
        {
            case "System.String": //字符串类型
                newCell.SetCellValue(drValue);
                break;
            case "System.DateTime": //日期类型
                DateTime dateV;
                DateTime.TryParse(drValue, out dateV);
                newCell.SetCellValue(dateV);
                newCell.CellStyle = dateStyle; //格式化显示
                break;
            case "System.Boolean": //布尔型
                bool boolV = false;
                bool.TryParse(drValue, out boolV);
                newCell.SetCellValue(boolV);
                break;
            case "System.Int16": //整型
            case "System.Int32":
            case "System.Int64":
            case "System.Byte":
                int intV = 0;
                int.TryParse(drValue, out intV);
                newCell.SetCellValue(intV);
                break;
            case "System.Single":
            case "System.Decimal": //浮点型
            case "System.Double":
                double doubV = 0;
                double.TryParse(drValue, out doubV);
                newCell.SetCellValue(doubV);
                break;
            case "System.DBNull": //空值处理
                newCell.SetCellValue("");
                break;
            default:
                newCell.SetCellValue(drValue);
                break;
        }
    }

    static void SetBorderStyle(ICellStyle style)
    {
        style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
    }

    /// <summary>
    /// 创建标题单元格样式
    /// </summary>
    /// <param name="wb"></param>
    /// <returns></returns>
    static ICellStyle CreateTitleStyle(IWorkbook wb)
    {
        ICellStyle titleStyle = wb.CreateCellStyle() as ICellStyle;
        titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        IFont font = wb.CreateFont() as IFont;
        font.FontHeightInPoints = 20;
        font.Boldweight = 700;
        titleStyle.SetFont(font);
        SetBorderStyle(titleStyle);
        return titleStyle;
    }

    /// <summary>
    /// 创建列名单元格样式
    /// </summary>
    /// <param name="wb"></param>
    /// <returns></returns>
    static ICellStyle CreateHeadStyle(IWorkbook wb)
    {
        ICellStyle headStyle = wb.CreateCellStyle() as ICellStyle;
        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
        IFont font = wb.CreateFont() as IFont;
        font.FontHeightInPoints = 10;
        font.Boldweight = 700;
        headStyle.SetFont(font);
        headStyle.FillPattern = FillPattern.SolidForeground;
        headStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
        SetBorderStyle(headStyle);
        return headStyle;
    }

    /// <summary>
    ///  创建单元格样式
    /// </summary>
    /// <param name="wb"></param>
    /// <returns></returns>
    static ICellStyle CreateCellStyle(IWorkbook wb)
    {
        ICellStyle cellStyle = wb.CreateCellStyle() as ICellStyle;
        SetBorderStyle(cellStyle);
        return cellStyle;
    }

    /// <summary>
    /// 创建日期类型单元格样式
    /// </summary>
    /// <param name="wb"></param>
    /// <returns></returns>
    static ICellStyle CreateDateStyle(IWorkbook wb)
    {
        ICellStyle dateStyle = wb.CreateCellStyle() as ICellStyle;
        IDataFormat format = wb.CreateDataFormat() as IDataFormat;
        dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
        SetBorderStyle(dateStyle);
        return dateStyle;
    }
    #endregion

    #region 从datatable中将数据导出到excel
    /// <summary>
    /// DataTable导出到Excel的MemoryStream
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    static MemoryStream ExportData(Object dtSource, DataGridViewColumnCollection Columns = null)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;
        #region 右击文件 属性信息

        //{
        //    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        //    dsi.Company = "http://www.baidu.com/";
        //    workbook.DocumentSummaryInformation = dsi;

        //    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        //    si.Author = "石书伟"; //填加xls文件作者信息
        //    si.ApplicationName = "NPOI测试程序"; //填加xls文件创建程序信息
        //    si.LastAuthor = "石书伟2"; //填加xls文件最后保存者信息
        //    si.Comments = "说明信息"; //填加xls文件作者信息
        //    si.Title = "NPOI测试"; //填加xls文件标题信息
        //    si.Subject = "NPOI测试Demo"; //填加文件主题信息
        //    si.CreateDateTime = DateTime.Now;
        //    workbook.SummaryInformation = si;
        //}

        #endregion
        ICellStyle dateStyle = CreateDateStyle(workbook);
        ICellStyle headStyle = CreateHeadStyle(workbook);
        ICellStyle cellStyle = CreateCellStyle(workbook);

        DataTable dt = null;
        bool isDt = true;
        DataGridView dgv = null;
        if (dtSource is DataTable)
        {
            dt = dtSource as DataTable;
            isDt = true;
        }
        if (dtSource is DataGridView)
        {
            dgv = dtSource as DataGridView;
            dt = dgv.DataSource as DataTable;
            isDt = false;
            if (Columns == null || Columns.Count == 0)
            {
                Columns = dgv.Columns;
            }
        }
        
        int rowIndex = 0;

        foreach (DataRow row in dt.Rows)
        {
            #region 新建表，填充表头，填充列头，样式
            if (rowIndex == 65535 || rowIndex == 0)
            {
                if (rowIndex != 0)
                {
                    sheet = workbook.CreateSheet() as HSSFSheet;
                }
                #region 列头及样式
                HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
                headerRow.Height = MyHeaderRowHeight;
                if (isDt && Columns == null)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                        //设置列宽
                        sheet.SetColumnWidth(column.Ordinal, (MyColWidth + 1) * 256);
                    }
                }
                else
                {
                    int index = 0;
                    foreach (DataGridViewColumn column in Columns)
                    {
                        // 去掉隐藏列和checkbox列
                        if (column.Visible && column.HeaderText != "")
                        {
                            headerRow.CreateCell(index).SetCellValue(column.HeaderText);
                            headerRow.GetCell(index).CellStyle = headStyle;
                            //设置列宽
                            sheet.SetColumnWidth(index, (MyColWidth + 1) * 256);

                            index++;
                        }
                    }
                }
                #endregion
                rowIndex = 1;
            }
            #endregion

            #region 填充内容
            HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;
            dataRow.Height = MyColRowHeight;
            if (isDt && Columns == null)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal) as HSSFCell;
                    newCell.CellStyle = cellStyle;
                    string drValue = row[column].ToString();
                    SetCellValue(drValue, column.DataType.ToString(), dateStyle, ref newCell);
                }
            }
            else
            {
                int index = 0;
                foreach (DataGridViewColumn column in Columns)
                {
                    // 去掉隐藏列和checkbox列
                    if (column.Visible && column.HeaderText != "")
                    {
                        ICell newCell = dataRow.CreateCell(index) as HSSFCell;
                        newCell.CellStyle = cellStyle;
                        try
                        {
                            if (dt.Columns.Contains(column.DataPropertyName))
                            {
                                object objVal = row[column.DataPropertyName];
                                string drValue = objVal.ToString();
                                SetCellValue(drValue, objVal.GetType().ToString(), dateStyle, ref newCell);
                            }
                            else
                            {
                                newCell.SetCellValue("");
                            }
                        }
                        catch (Exception exe)
                        {
                            newCell.SetCellValue("");
                        }
                        index++;
                    }
                }
            }
            #endregion

            rowIndex++;
        }
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            //sheet.Dispose();
            //workbook.Dispose();

            return ms;
        }
    }

    /// <summary>
    /// DataTable导出到Excel的MemoryStream
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    static void ExportDataI(Object dtSource, FileStream fs, DataGridViewColumnCollection Columns = null)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.CreateSheet() as XSSFSheet;

        #region 右击文件 属性信息

        //{
        //    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        //    dsi.Company = "http://www.baidu.com/";
        //    workbook.DocumentSummaryInformation = dsi;

        //    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        //    si.Author = "石书伟"; //填加xls文件作者信息
        //    si.ApplicationName = "NPOI测试程序"; //填加xls文件创建程序信息
        //    si.LastAuthor = "石书伟2"; //填加xls文件最后保存者信息
        //    si.Comments = "说明信息"; //填加xls文件作者信息
        //    si.Title = "NPOI测试"; //填加xls文件标题信息
        //    si.Subject = "NPOI测试Demo"; //填加文件主题信息
        //    si.CreateDateTime = DateTime.Now;
        //    workbook.SummaryInformation = si;
        //}

        #endregion

        ICellStyle dateStyle = CreateDateStyle(workbook);
        ICellStyle headStyle = CreateHeadStyle(workbook);
        ICellStyle cellStyle = CreateCellStyle(workbook);

        DataTable dt = null;
        bool isDt = true;
        DataGridView dgv = null;
        if (dtSource is DataTable)
        {
            dt = dtSource as DataTable;
            isDt = true;
        }
        if (dtSource is DataGridView)
        {
            dgv = dtSource as DataGridView;
            dt = dgv.DataSource as DataTable;
            isDt = false;
        }
        
        int rowIndex = 0;
        foreach (DataRow row in dt.Rows)
        {
            #region 新建表，填充表头，填充列头，样式

            if (rowIndex == 0)
            {
                #region 列头及样式
                XSSFRow headerRow = sheet.CreateRow(0) as XSSFRow;
                headerRow.Height = MyHeaderRowHeight;
                if (isDt)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                        //设置列宽
                        sheet.SetColumnWidth(column.Ordinal, (MyColWidth + 1) * 256);
                    }
                }
                else
                {
                    int index = 0;
                    foreach (DataGridViewColumn column in dgv.Columns)
                    {
                        // 去掉隐藏列和checkbox列
                        if (column.Visible && column.HeaderText != "")
                        {
                            headerRow.CreateCell(index).SetCellValue(column.HeaderText);
                            headerRow.GetCell(index).CellStyle = headStyle;
                            //设置列宽
                            sheet.SetColumnWidth(index, (MyColWidth + 1) * 256);

                            index++;
                        }
                    }
                }
                #endregion

                rowIndex = 1;
            }

            #endregion

            #region 填充内容
            XSSFRow dataRow = sheet.CreateRow(rowIndex) as XSSFRow;
            dataRow.Height = MyColRowHeight;
            if (isDt)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal) as XSSFCell;
                    newCell.CellStyle = cellStyle;
                    string drValue = row[column.ColumnName].ToString();
                    SetCellValue(drValue, column.DataType.ToString(), dateStyle, ref newCell);
                }
            }
            else
            {
                int index = 0;
                foreach (DataGridViewColumn column in dgv.Columns)
                {
                    // 去掉隐藏列和checkbox列
                    if (column.Visible && column.HeaderText != "")
                    {
                        ICell newCell = dataRow.CreateCell(index) as XSSFCell;
                        newCell.CellStyle = cellStyle;
                        try
                        {
                            if (dt.Columns.Contains(column.DataPropertyName))
                            {
                                object objVal = row[column.DataPropertyName];
                                string drValue = objVal.ToString();
                                SetCellValue(drValue, objVal.GetType().ToString(), dateStyle, ref newCell);
                            }
                            else
                            {
                                newCell.SetCellValue("");
                            }
                            
                        }
                        catch (Exception exe)
                        {
                            newCell.SetCellValue("");
                        }
                        index++;
                    }
                }
            }
           
            #endregion
            rowIndex++;
        }
        workbook.Write(fs);
        fs.Close();
    }

    /// <summary>
    /// DataTable导出到Excel文件
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    /// <param name="strFileName">保存位置</param>
    static void ExportDataToExcel(Object dtSource, string strFileName, DataGridViewColumnCollection Columns = null)
    {
        string[] temp = strFileName.Split('.');

        Int32 colCount = 0;
        Int32 rowsCount = 0;
        if (dtSource is DataTable)
        {
            colCount = (dtSource as DataTable).Columns.Count;
            rowsCount = (dtSource as DataTable).Rows.Count;

            if (Columns != null && Columns.Count > 0)
            {
                for (int i = 0; i < Columns.Count; i++)
                {
                    // 去掉隐藏列和checkbox列
                    if (Columns[i].Visible && Columns[i].HeaderText != "")
                    {
                        colCount++;
                    }
                }
            }
        }
        if (dtSource is DataGridView)
        {
            colCount = 0;
            DataGridView dgv = dtSource as DataGridView;
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                // 去掉隐藏列和checkbox列
                if (dgv.Columns[i].Visible && dgv.Columns[i].HeaderText != "")
                {
                    colCount++;
                }
            }
            rowsCount = (dtSource as DataGridView).Rows.Count;
        }
        if (temp[temp.Length - 1] == "xls" && colCount < 256 && rowsCount < 65536)
        {
            using (MemoryStream ms = ExportData(dtSource, Columns))
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }
        else
        {
            if (temp[temp.Length - 1] == "xls")
                strFileName = strFileName + "x";

            using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
            {
                ExportDataI(dtSource, fs, Columns);
            }
        }
    }

    /// <summary>
    /// DataGridView导出到Excel文件
    /// </summary>
    /// <param name="dtSource">源DataGridView</param>
    /// <param name="strHeaderText">表头文本</param>
    /// <param name="strFileName">保存位置</param>
    public static void ExportDataGridViewToExcel(DataGridView dtSource, string strFileName = "")
    {
        SaveFileDialog sfd = new SaveFileDialog();
        sfd.FileName = strFileName;
        sfd.Filter = @"Excel 2003 工作表(*.xls)|*.xls|Excel 2007 工作表(*.xlsx)|*.xlsx";
        if (sfd.ShowDialog() != DialogResult.OK)
        {
            return;
        }
        ExportDataToExcel(dtSource, sfd.FileName);
    }

    /// <summary>
    /// DataTable导出到Excel文件
    /// </summary>
    /// <param name="dtSource">源DataGridView</param>
    /// <param name="strHeaderText">表头文本</param>
    /// <param name="strFileName">保存位置</param>
    public static void ExportDataTableToExcel(DataTable dtSource, string strFileName = "")
    {
        SaveFileDialog sfd = new SaveFileDialog();
        sfd.FileName = strFileName;
        sfd.Filter = @"Excel 2003 工作表(*.xls)|*.xls|Excel 2007 工作表(*.xlsx)|*.xlsx";
        if (sfd.ShowDialog() != DialogResult.OK)
        {
            return;
        }
        ExportDataToExcel(dtSource, sfd.FileName);
    }

    /// <summary>
    /// 根据datagridview的列导出datatable的数据
    /// </summary>
    /// <param name="dtSource"></param>
    /// <param name="DGVColumns"></param>
    /// <param name="strFileName"></param>
    public static void ExportDataTableToExcel(DataTable dtSource, DataGridViewColumnCollection DGVColumns, string strFileName = "")
    {
        SaveFileDialog sfd = new SaveFileDialog();
        sfd.FileName = strFileName;
        sfd.Filter = @"Excel 2003 工作表(*.xls)|*.xls|Excel 2007 工作表(*.xlsx)|*.xlsx";
        if (sfd.ShowDialog() != DialogResult.OK)
        {
            return;
        }
        ExportDataToExcel(dtSource, sfd.FileName, DGVColumns);
    }
    #endregion

    #region 多个datatable导出到excel
    static bool CheckIsXlsx(List<DataTable> dtList)
    {
        bool bl = false;
        foreach (DataTable dtSource in dtList)
        {
            if (dtSource.Columns.Count >= 256 || dtSource.Rows.Count >= 65536)
            {
                bl = true;
                break; // TODO: might not be correct. Was : Exit For
            }
        }
        return bl;
    }
    public static void ExportDTStoExcel(List<DataTable> dtList, string strFileName)
    {
        string[] temp = strFileName.Split('.');
        if (temp[temp.Length - 1] == "xls" && !CheckIsXlsx(dtList))
        {
            using (MemoryStream ms = ExportDTS(dtList))
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }
        else
        {
            if (temp[temp.Length - 1] == "xls")
            {
                strFileName = strFileName + "x";
            }
            using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
            {
                ExportDTSI(dtList, fs);
            }
        }
    }
    static MemoryStream ExportDTS(List<DataTable> dtList)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        ICellStyle headStyle = CreateHeadStyle(workbook);
        ICellStyle cellStyle = CreateCellStyle(workbook);
        ICellStyle dateStyle = CreateDateStyle(workbook);
        foreach (DataTable dtSource in dtList)
        {
            HSSFSheet sheet = workbook.CreateSheet(dtSource.TableName) as HSSFSheet;
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = MyColWidth;
            }
            int i = 0;
            while (i < dtSource.Rows.Count)
            {
                int j = 0;
                while (j < dtSource.Columns.Count)
                {
                    int intTemp = MyColWidth;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                    j++;
                }
                i++;
            }
            int rowIndex = 0;
            foreach (DataRow row in dtSource.Rows)
            {
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet(dtSource.TableName) as HSSFSheet;
                    }
                    rowIndex = 0;
                    if (true)
                    {
                        HSSFRow headerRow = sheet.CreateRow(rowIndex) as HSSFRow;
                        headerRow.Height = MyHeaderRowHeight;
                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                        }
                    }
                    rowIndex += 1;
                }
                HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;
                dataRow.Height = MyColRowHeight;
                foreach (DataColumn column in dtSource.Columns)
                {
                   ICell newCell = dataRow.CreateCell(column.Ordinal) as HSSFCell;
                    newCell.CellStyle = cellStyle;
                    string drValue = row[column].ToString();
                    SetCellValue(drValue, column.DataType.ToString(), dateStyle, ref newCell);
                }
                rowIndex++;
            }
        }
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            return ms;
        }
    }
    static void ExportDTSI(List<DataTable> dtList, FileStream fs)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        ICellStyle headStyle = CreateHeadStyle(workbook);
        ICellStyle cellStyle = CreateCellStyle(workbook);
        ICellStyle dateStyle = CreateDateStyle(workbook);
        foreach (DataTable dtSource in dtList)
        {
            XSSFSheet sheet = workbook.CreateSheet(dtSource.TableName) as XSSFSheet;
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = MyColWidth;
            }
            int i = 0;
            while (i < dtSource.Rows.Count)
            {
                int j = 0;
                while (j < dtSource.Columns.Count)
                {
                    int intTemp = MyColWidth;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                    j++;
                }
                i++;
            }
            int rowIndex = 0;
            foreach (DataRow row in dtSource.Rows)
            {
                if (rowIndex == 0)
                {
                    if (true)
                    {
                        XSSFRow headerRow = sheet.CreateRow(rowIndex) as XSSFRow;
                        headerRow.Height = MyHeaderRowHeight;
                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                        }
                    }
                    rowIndex += 1;
                }
                XSSFRow dataRow = sheet.CreateRow(rowIndex) as XSSFRow;
                dataRow.Height = MyColRowHeight;
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal) as XSSFCell;
                    newCell.CellStyle = cellStyle;
                    string drValue = row[column].ToString();
                    SetCellValue(drValue, column.DataType.ToString(), dateStyle, ref newCell);
                }
                rowIndex++;
            }
        }
        workbook.Write(fs);
        fs.Close();
    }
    #endregion
    #endregion
    #region 未经过测试代码，可参考

    #region 从datatable中将数据导出到excel
    /// <summary>
    /// DataTable导出到Excel的MemoryStream
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    static MemoryStream ExportDT(DataTable dtSource, string strHeaderText)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;

        #region 右击文件 属性信息

        //{
        //    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        //    dsi.Company = "http://www.baidu.com/";
        //    workbook.DocumentSummaryInformation = dsi;

        //    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        //    si.Author = "石书伟"; //填加xls文件作者信息
        //    si.ApplicationName = "NPOI测试程序"; //填加xls文件创建程序信息
        //    si.LastAuthor = "石书伟2"; //填加xls文件最后保存者信息
        //    si.Comments = "说明信息"; //填加xls文件作者信息
        //    si.Title = "NPOI测试"; //填加xls文件标题信息
        //    si.Subject = "NPOI测试Demo"; //填加文件主题信息
        //    si.CreateDateTime = DateTime.Now;
        //    workbook.SummaryInformation = si;
        //}

        #endregion

        HSSFCellStyle dateStyle = workbook.CreateCellStyle() as HSSFCellStyle;
        HSSFDataFormat format = workbook.CreateDataFormat() as HSSFDataFormat;
        dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

        //取得列宽
        int[] arrColWidth = new int[dtSource.Columns.Count];
        foreach (DataColumn item in dtSource.Columns)
        {
            arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
        }
        for (int i = 0; i < dtSource.Rows.Count; i++)
        {
            for (int j = 0; j < dtSource.Columns.Count; j++)
            {
                int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                if (intTemp > arrColWidth[j])
                {
                    arrColWidth[j] = intTemp;
                }
            }
        }
        int rowIndex = 0;

        foreach (DataRow row in dtSource.Rows)
        {
            #region 新建表，填充表头，填充列头，样式

            if (rowIndex == 65535 || rowIndex == 0)
            {
                if (rowIndex != 0)
                {
                    sheet = workbook.CreateSheet() as HSSFSheet;
                }

                #region 表头及样式

                {
                    HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
                    headerRow.HeightInPoints = 25;
                    headerRow.CreateCell(0).SetCellValue(strHeaderText);

                    HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                    headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    HSSFFont font = workbook.CreateFont() as HSSFFont;
                    font.FontHeightInPoints = 20;
                    font.Boldweight = 700;
                    headStyle.SetFont(font);

                    headerRow.GetCell(0).CellStyle = headStyle;

                    sheet.AddMergedRegion(new Region(0, 0, 0, dtSource.Columns.Count - 1));
                    //headerRow.Dispose();
                }

                #endregion


                #region 列头及样式

                {
                    HSSFRow headerRow = sheet.CreateRow(1) as HSSFRow;


                    HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                    headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    HSSFFont font = workbook.CreateFont() as HSSFFont;
                    font.FontHeightInPoints = 10;
                    font.Boldweight = 700;
                    headStyle.SetFont(font);


                    foreach (DataColumn column in dtSource.Columns)
                    {
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                        //设置列宽
                        sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);

                    }
                    //headerRow.Dispose();
                }

                #endregion

                rowIndex = 2;
            }

            #endregion

            #region 填充内容

            HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;
            foreach (DataColumn column in dtSource.Columns)
            {
                HSSFCell newCell = dataRow.CreateCell(column.Ordinal) as HSSFCell;

                string drValue = row[column].ToString();

                switch (column.DataType.ToString())
                {
                    case "System.String": //字符串类型
                        double result;
                        if (isNumeric(drValue, out result))
                        {

                            double.TryParse(drValue, out result);
                            newCell.SetCellValue(result);
                            break;
                        }
                        else
                        {
                            newCell.SetCellValue(drValue);
                            break;
                        }

                    case "System.DateTime": //日期类型
                        DateTime dateV;
                        DateTime.TryParse(drValue, out dateV);
                        newCell.SetCellValue(dateV);

                        newCell.CellStyle = dateStyle; //格式化显示
                        break;
                    case "System.Boolean": //布尔型
                        bool boolV = false;
                        bool.TryParse(drValue, out boolV);
                        newCell.SetCellValue(boolV);
                        break;
                    case "System.Int16": //整型
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Byte":
                        int intV = 0;
                        int.TryParse(drValue, out intV);
                        newCell.SetCellValue(intV);
                        break;
                    case "System.Single":
                    case "System.Decimal": //浮点型
                    case "System.Double":
                        double doubV = 0;
                        double.TryParse(drValue, out doubV);
                        newCell.SetCellValue(doubV);
                        break;
                    case "System.DBNull": //空值处理
                        newCell.SetCellValue("");
                        break;
                    default:
                        newCell.SetCellValue(drValue);
                        break;
                }

            }

            #endregion

            rowIndex++;
        }
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            //sheet.Dispose();
            //workbook.Dispose();

            return ms;
        }
    }

    /// <summary>
    /// DataTable导出到Excel的MemoryStream
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    static void ExportDTI(DataTable dtSource, string strHeaderText, FileStream fs)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.CreateSheet() as XSSFSheet;

        #region 右击文件 属性信息

        //{
        //    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        //    dsi.Company = "http://www.baidu.com/";
        //    workbook.DocumentSummaryInformation = dsi;

        //    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
        //    si.Author = "石书伟"; //填加xls文件作者信息
        //    si.ApplicationName = "NPOI测试程序"; //填加xls文件创建程序信息
        //    si.LastAuthor = "石书伟2"; //填加xls文件最后保存者信息
        //    si.Comments = "说明信息"; //填加xls文件作者信息
        //    si.Title = "NPOI测试"; //填加xls文件标题信息
        //    si.Subject = "NPOI测试Demo"; //填加文件主题信息
        //    si.CreateDateTime = DateTime.Now;
        //    workbook.SummaryInformation = si;
        //}

        #endregion

        XSSFCellStyle dateStyle = workbook.CreateCellStyle() as XSSFCellStyle;
        XSSFDataFormat format = workbook.CreateDataFormat() as XSSFDataFormat;
        dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

        //取得列宽
        int[] arrColWidth = new int[dtSource.Columns.Count];
        foreach (DataColumn item in dtSource.Columns)
        {
            arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
        }
        for (int i = 0; i < dtSource.Rows.Count; i++)
        {
            for (int j = 0; j < dtSource.Columns.Count; j++)
            {
                int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                if (intTemp > arrColWidth[j])
                {
                    arrColWidth[j] = intTemp;
                }
            }
        }
        int rowIndex = 0;

        foreach (DataRow row in dtSource.Rows)
        {
            #region 新建表，填充表头，填充列头，样式

            if (rowIndex == 0)
            {
                #region 表头及样式
                //{
                //    XSSFRow headerRow = sheet.CreateRow(0) as XSSFRow;
                //    headerRow.HeightInPoints = 25;
                //    headerRow.CreateCell(0).SetCellValue(strHeaderText);

                //    XSSFCellStyle headStyle = workbook.CreateCellStyle() as XSSFCellStyle;
                //    headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                //    XSSFFont font = workbook.CreateFont() as XSSFFont;
                //    font.FontHeightInPoints = 20;
                //    font.Boldweight = 700;
                //    headStyle.SetFont(font);

                //    headerRow.GetCell(0).CellStyle = headStyle;

                //    //sheet.AddMergedRegion(new Region(0, 0, 0, dtSource.Columns.Count - 1));
                //    //headerRow.Dispose();
                //}

                #endregion


                #region 列头及样式

                {
                    XSSFRow headerRow = sheet.CreateRow(0) as XSSFRow;


                    XSSFCellStyle headStyle = workbook.CreateCellStyle() as XSSFCellStyle;
                    headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    XSSFFont font = workbook.CreateFont() as XSSFFont;
                    font.FontHeightInPoints = 10;
                    font.Boldweight = 700;
                    headStyle.SetFont(font);


                    foreach (DataColumn column in dtSource.Columns)
                    {
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                        //设置列宽
                        sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);

                    }
                    //headerRow.Dispose();
                }

                #endregion

                rowIndex = 1;
            }

            #endregion

            #region 填充内容

            XSSFRow dataRow = sheet.CreateRow(rowIndex) as XSSFRow;
            foreach (DataColumn column in dtSource.Columns)
            {
                XSSFCell newCell = dataRow.CreateCell(column.Ordinal) as XSSFCell;

                string drValue = row[column].ToString();

                switch (column.DataType.ToString())
                {
                    case "System.String": //字符串类型
                        double result;
                        if (isNumeric(drValue, out result))
                        {

                            double.TryParse(drValue, out result);
                            newCell.SetCellValue(result);
                            break;
                        }
                        else
                        {
                            newCell.SetCellValue(drValue);
                            break;
                        }

                    case "System.DateTime": //日期类型
                        DateTime dateV;
                        DateTime.TryParse(drValue, out dateV);
                        newCell.SetCellValue(dateV);

                        newCell.CellStyle = dateStyle; //格式化显示
                        break;
                    case "System.Boolean": //布尔型
                        bool boolV = false;
                        bool.TryParse(drValue, out boolV);
                        newCell.SetCellValue(boolV);
                        break;
                    case "System.Int16": //整型
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Byte":
                        int intV = 0;
                        int.TryParse(drValue, out intV);
                        newCell.SetCellValue(intV);
                        break;
                    case "System.Single":
                    case "System.Decimal": //浮点型
                    case "System.Double":
                        double doubV = 0;
                        double.TryParse(drValue, out doubV);
                        newCell.SetCellValue(doubV);
                        break;
                    case "System.DBNull": //空值处理
                        newCell.SetCellValue("");
                        break;
                    default:
                        newCell.SetCellValue(drValue);
                        break;
                }

            }

            #endregion

            rowIndex++;
        }
        workbook.Write(fs);
        fs.Close();
    }

    /// <summary>
    /// DataTable导出到Excel文件
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    /// <param name="strFileName">保存位置</param>
    public static void ExportDTtoExcel(DataTable dtSource, string strHeaderText, string strFileName)
    {
        string[] temp = strFileName.Split('.');

        if (temp[temp.Length - 1] == "xls" && dtSource.Columns.Count < 256 && dtSource.Rows.Count < 65536)
        {
            using (MemoryStream ms = ExportDT(dtSource, strHeaderText))
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
        }
        else
        {
            if (temp[temp.Length - 1] == "xls")
                strFileName = strFileName + "x";

            using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
            {
                ExportDTI(dtSource, strHeaderText, fs);
            }
        }
    }
    #endregion

    #region 从excel中将数据导出到datatable
    /// <summary>
    /// 读取excel 默认第一行为标头
    /// </summary>
    /// <param name="strFileName">excel文档路径</param>
    /// <returns></returns>
    public static DataTable ImportExceltoDt(string strFileName)
    {
        DataTable dt = new DataTable();
        IWorkbook wb;
        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
        {
            wb = WorkbookFactory.Create(file);
        }
        ISheet sheet = wb.GetSheetAt(0);
        dt = ImportDt(sheet, 0, true);
        return dt;
    }

    /// <summary>
    /// 读取Excel流到DataTable
    /// </summary>
    /// <param name="stream">Excel流</param>
    /// <returns>第一个sheet中的数据</returns>
    public static DataTable ImportExceltoDt(Stream stream)
    {
        try
        {
            DataTable dt = new DataTable();
            IWorkbook wb;
            using (stream)
            {
                wb = WorkbookFactory.Create(stream);
            }
            ISheet sheet = wb.GetSheetAt(0);
            dt = ImportDt(sheet, 0, true);
            return dt;
        }
        catch (Exception)
        {

            throw;
        }
    }

    /// <summary>
    /// 读取Excel流到DataTable
    /// </summary>
    /// <param name="stream">Excel流</param>
    /// <param name="sheetName">表单名</param>
    /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
    /// <returns>指定sheet中的数据</returns>
    public static DataTable ImportExceltoDt(Stream stream, string sheetName, int HeaderRowIndex)
    {
        try
        {
            DataTable dt = new DataTable();
            IWorkbook wb;
            using (stream)
            {
                wb = WorkbookFactory.Create(stream);
            }
            ISheet sheet = wb.GetSheet(sheetName);
            dt = ImportDt(sheet, HeaderRowIndex, true);
            return dt;
        }
        catch (Exception)
        {

            throw;
        }
    }

    /// <summary>
    /// 读取Excel流到DataSet
    /// </summary>
    /// <param name="stream">Excel流</param>
    /// <returns>Excel中的数据</returns>
    public static DataSet ImportExceltoDs(Stream stream)
    {
        try
        {
            DataSet ds = new DataSet();
            IWorkbook wb;
            using (stream)
            {
                wb = WorkbookFactory.Create(stream);
            }
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                DataTable dt = new DataTable();
                ISheet sheet = wb.GetSheetAt(i);
                dt = ImportDt(sheet, 0, true);
                ds.Tables.Add(dt);
            }
            return ds;
        }
        catch (Exception)
        {

            throw;
        }
    }

    /// <summary>
    /// 读取Excel流到DataSet
    /// </summary>
    /// <param name="stream">Excel流</param>
    /// <param name="dict">字典参数，key：sheet名，value：列头所在行号，-1表示没有列头</param>
    /// <returns>Excel中的数据</returns>
    public static DataSet ImportExceltoDs(Stream stream, Dictionary<string, int> dict)
    {
        try
        {
            DataSet ds = new DataSet();
            IWorkbook wb;
            using (stream)
            {
                wb = WorkbookFactory.Create(stream);
            }
            foreach (string key in dict.Keys)
            {
                DataTable dt = new DataTable();
                ISheet sheet = wb.GetSheet(key);
                dt = ImportDt(sheet, dict[key], true);
                ds.Tables.Add(dt);
            }
            return ds;
        }
        catch (Exception)
        {

            throw;
        }
    }

    /// <summary>
    /// 读取excel
    /// </summary>
    /// <param name="strFileName">excel文件路径</param>
    /// <param name="sheet">需要导出的sheet</param>
    /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
    /// <returns></returns>
    public static DataTable ImportExceltoDt(string strFileName, string SheetName, int HeaderRowIndex)
    {
        HSSFWorkbook workbook;
        IWorkbook wb;
        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
        {
            wb = new HSSFWorkbook(file);
        }
        ISheet sheet = wb.GetSheet(SheetName);
        DataTable table = new DataTable();
        table = ImportDt(sheet, HeaderRowIndex, true);
        //ExcelFileStream.Close();
        workbook = null;
        sheet = null;
        return table;
    }

    /// <summary>
    /// 读取excel
    /// </summary>
    /// <param name="strFileName">excel文件路径</param>
    /// <param name="sheet">需要导出的sheet序号</param>
    /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
    /// <returns></returns>
    public static DataTable ImportExceltoDt(string strFileName, int SheetIndex, int HeaderRowIndex)
    {
        HSSFWorkbook workbook;
        IWorkbook wb;
        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
        {
            wb = WorkbookFactory.Create(file);
        }
        ISheet isheet = wb.GetSheetAt(SheetIndex);
        DataTable table = new DataTable();
        table = ImportDt(isheet, HeaderRowIndex, true);
        //ExcelFileStream.Close();
        workbook = null;
        isheet = null;
        return table;
    }

    /// <summary>
    /// 读取excel
    /// </summary>
    /// <param name="strFileName">excel文件路径</param>
    /// <param name="sheet">需要导出的sheet</param>
    /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
    /// <returns></returns>
    public static DataTable ImportExceltoDt(string strFileName, string SheetName, int HeaderRowIndex, bool needHeader)
    {
        HSSFWorkbook workbook;
        IWorkbook wb;
        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
        {
            wb = WorkbookFactory.Create(file);
        }
        ISheet sheet = wb.GetSheet(SheetName);
        DataTable table = new DataTable();
        table = ImportDt(sheet, HeaderRowIndex, needHeader);
        //ExcelFileStream.Close();
        workbook = null;
        sheet = null;
        return table;
    }

    /// <summary>
    /// 读取excel
    /// </summary>
    /// <param name="strFileName">excel文件路径</param>
    /// <param name="sheet">需要导出的sheet序号</param>
    /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
    /// <returns></returns>
    public static DataTable ImportExceltoDt(string strFileName, int SheetIndex, int HeaderRowIndex, bool needHeader)
    {
        HSSFWorkbook workbook;
        IWorkbook wb;
        using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
        {
            wb = WorkbookFactory.Create(file);
        }
        ISheet sheet = wb.GetSheetAt(SheetIndex);
        DataTable table = new DataTable();
        table = ImportDt(sheet, HeaderRowIndex, needHeader);
        //ExcelFileStream.Close();
        workbook = null;
        sheet = null;
        return table;
    }

    /// <summary>
    /// 将制定sheet中的数据导出到datatable中
    /// </summary>
    /// <param name="sheet">需要导出的sheet</param>
    /// <param name="HeaderRowIndex">列头所在行号，-1表示没有列头</param>
    /// <returns></returns>
    static DataTable ImportDt(ISheet sheet, int HeaderRowIndex, bool needHeader)
    {
        DataTable table = new DataTable();
        IRow headerRow;
        int cellCount;
        try
        {
            if (HeaderRowIndex < 0 || !needHeader)
            {
                headerRow = sheet.GetRow(0);
                cellCount = headerRow.LastCellNum;

                for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                {
                    DataColumn column = new DataColumn(Convert.ToString(i));
                    table.Columns.Add(column);
                }
            }
            else
            {
                headerRow = sheet.GetRow(HeaderRowIndex);
                cellCount = headerRow.LastCellNum;

                for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                {
                    if (headerRow.GetCell(i) == null)
                    {
                        if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                        {
                            DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn(Convert.ToString(i));
                            table.Columns.Add(column);
                        }

                    }
                    else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                    {
                        DataColumn column = new DataColumn(Convert.ToString("重复列名" + i));
                        table.Columns.Add(column);
                    }
                    else
                    {
                        DataColumn column = new DataColumn(headerRow.GetCell(i).ToString());
                        table.Columns.Add(column);
                    }
                }
            }
            int rowCount = sheet.LastRowNum;
            for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                try
                {
                    IRow row;
                    if (sheet.GetRow(i) == null)
                    {
                        row = sheet.CreateRow(i);
                    }
                    else
                    {
                        row = sheet.GetRow(i);
                    }

                    DataRow dataRow = table.NewRow();

                    for (int j = row.FirstCellNum; j <= cellCount; j++)
                    {
                        try
                        {
                            if (row.GetCell(j) != null)
                            {
                                switch (row.GetCell(j).CellType)
                                {
                                    case CellType.String:
                                        string str = row.GetCell(j).StringCellValue;
                                        if (str != null && str.Length > 0)
                                        {
                                            dataRow[j] = str.ToString();
                                        }
                                        else
                                        {
                                            dataRow[j] = null;
                                        }
                                        break;
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                        {
                                            dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                        }
                                        else
                                        {
                                            dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                        }
                                        break;
                                    case CellType.Boolean:
                                        dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                        break;
                                    case CellType.Error:
                                        dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                        break;
                                    case CellType.Formula:
                                        switch (row.GetCell(j).CachedFormulaResultType)
                                        {
                                            case CellType.String:
                                                string strFORMULA = row.GetCell(j).StringCellValue;
                                                if (strFORMULA != null && strFORMULA.Length > 0)
                                                {
                                                    dataRow[j] = strFORMULA.ToString();
                                                }
                                                else
                                                {
                                                    dataRow[j] = null;
                                                }
                                                break;
                                            case CellType.Numeric:
                                                dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                break;
                                            case CellType.Boolean:
                                                dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                break;
                                            case CellType.Error:
                                                dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                break;
                                            default:
                                                dataRow[j] = "";
                                                break;
                                        }
                                        break;
                                    default:
                                        dataRow[j] = "";
                                        break;
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            throw exception;
                        }
                    }
                    table.Rows.Add(dataRow);
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
        }
        catch (Exception exception)
        {
            throw exception;
        }
        return table;
    }

    #endregion

    public static void InsertSheet(string outputFile, string sheetname, DataTable dt)
    {
        FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);
        IWorkbook hssfworkbook = WorkbookFactory.Create(readfile);
        //HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
        int num = hssfworkbook.GetSheetIndex(sheetname);
        ISheet sheet1;
        if (num >= 0)
            sheet1 = hssfworkbook.GetSheet(sheetname);
        else
        {
            sheet1 = hssfworkbook.CreateSheet(sheetname);
        }


        try
        {
            if (sheet1.GetRow(0) == null)
            {
                sheet1.CreateRow(0);
            }
            for (int coluid = 0; coluid < dt.Columns.Count; coluid++)
            {
                if (sheet1.GetRow(0).GetCell(coluid) == null)
                {
                    sheet1.GetRow(0).CreateCell(coluid);
                }

                sheet1.GetRow(0).GetCell(coluid).SetCellValue(dt.Columns[coluid].ColumnName);
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }


        for (int i = 1; i <= dt.Rows.Count; i++)
        {
            try
            {
                if (sheet1.GetRow(i) == null)
                {
                    sheet1.CreateRow(i);
                }
                for (int coluid = 0; coluid < dt.Columns.Count; coluid++)
                {
                    if (sheet1.GetRow(i).GetCell(coluid) == null)
                    {
                        sheet1.GetRow(i).CreateCell(coluid);
                    }

                    sheet1.GetRow(i).GetCell(coluid).SetCellValue(dt.Rows[i - 1][coluid].ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        try
        {
            readfile.Close();

            FileStream writefile = new FileStream(outputFile, FileMode.OpenOrCreate, FileAccess.Write);
            hssfworkbook.Write(writefile);
            writefile.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    #region 更新excel中的数据
    /// <summary>
    /// 更新Excel表格
    /// </summary>
    /// <param name="outputFile">需更新的excel表格路径</param>
    /// <param name="sheetname">sheet名</param>
    /// <param name="updateData">需更新的数据</param>
    /// <param name="coluid">需更新的列号</param>
    /// <param name="rowid">需更新的开始行号</param>
    public static void UpdateExcel(string outputFile, string sheetname, string[] updateData, int coluid, int rowid)
    {
        //FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);
        IWorkbook hssfworkbook = null;// WorkbookFactory.Create(outputFile);
        //HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
        ISheet sheet1 = hssfworkbook.GetSheet(sheetname);
        for (int i = 0; i < updateData.Length; i++)
        {
            try
            {
                if (sheet1.GetRow(i + rowid) == null)
                {
                    sheet1.CreateRow(i + rowid);
                }
                if (sheet1.GetRow(i + rowid).GetCell(coluid) == null)
                {
                    sheet1.GetRow(i + rowid).CreateCell(coluid);
                }

                sheet1.GetRow(i + rowid).GetCell(coluid).SetCellValue(updateData[i]);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        try
        {
            //readfile.Close();
            FileStream writefile = new FileStream(outputFile, FileMode.OpenOrCreate, FileAccess.Write);
            hssfworkbook.Write(writefile);
            writefile.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }

    }

    /// <summary>
    /// 更新Excel表格
    /// </summary>
    /// <param name="outputFile">需更新的excel表格路径</param>
    /// <param name="sheetname">sheet名</param>
    /// <param name="updateData">需更新的数据</param>
    /// <param name="coluids">需更新的列号</param>
    /// <param name="rowid">需更新的开始行号</param>
    public static void UpdateExcel(string outputFile, string sheetname, string[][] updateData, int[] coluids, int rowid)
    {
        FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);

        HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
        readfile.Close();
        ISheet sheet1 = hssfworkbook.GetSheet(sheetname);
        for (int j = 0; j < coluids.Length; j++)
        {
            for (int i = 0; i < updateData[j].Length; i++)
            {
                try
                {
                    if (sheet1.GetRow(i + rowid) == null)
                    {
                        sheet1.CreateRow(i + rowid);
                    }
                    if (sheet1.GetRow(i + rowid).GetCell(coluids[j]) == null)
                    {
                        sheet1.GetRow(i + rowid).CreateCell(coluids[j]);
                    }
                    sheet1.GetRow(i + rowid).GetCell(coluids[j]).SetCellValue(updateData[j][i]);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        try
        {
            FileStream writefile = new FileStream(outputFile, FileMode.Create);
            hssfworkbook.Write(writefile);
            writefile.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// 更新Excel表格
    /// </summary>
    /// <param name="outputFile">需更新的excel表格路径</param>
    /// <param name="sheetname">sheet名</param>
    /// <param name="updateData">需更新的数据</param>
    /// <param name="coluid">需更新的列号</param>
    /// <param name="rowid">需更新的开始行号</param>
    public static void UpdateExcel(string outputFile, string sheetname, double[] updateData, int coluid, int rowid)
    {
        FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);

        HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
        ISheet sheet1 = hssfworkbook.GetSheet(sheetname);
        for (int i = 0; i < updateData.Length; i++)
        {
            try
            {
                if (sheet1.GetRow(i + rowid) == null)
                {
                    sheet1.CreateRow(i + rowid);
                }
                if (sheet1.GetRow(i + rowid).GetCell(coluid) == null)
                {
                    sheet1.GetRow(i + rowid).CreateCell(coluid);
                }

                sheet1.GetRow(i + rowid).GetCell(coluid).SetCellValue(updateData[i]);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        try
        {
            readfile.Close();
            FileStream writefile = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
            hssfworkbook.Write(writefile);
            writefile.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }

    }

    /// <summary>
    /// 更新Excel表格
    /// </summary>
    /// <param name="outputFile">需更新的excel表格路径</param>
    /// <param name="sheetname">sheet名</param>
    /// <param name="updateData">需更新的数据</param>
    /// <param name="coluids">需更新的列号</param>
    /// <param name="rowid">需更新的开始行号</param>
    public static void UpdateExcel(string outputFile, string sheetname, double[][] updateData, int[] coluids, int rowid)
    {
        FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);

        HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
        readfile.Close();
        ISheet sheet1 = hssfworkbook.GetSheet(sheetname);
        for (int j = 0; j < coluids.Length; j++)
        {
            for (int i = 0; i < updateData[j].Length; i++)
            {
                try
                {
                    if (sheet1.GetRow(i + rowid) == null)
                    {
                        sheet1.CreateRow(i + rowid);
                    }
                    if (sheet1.GetRow(i + rowid).GetCell(coluids[j]) == null)
                    {
                        sheet1.GetRow(i + rowid).CreateCell(coluids[j]);
                    }
                    sheet1.GetRow(i + rowid).GetCell(coluids[j]).SetCellValue(updateData[j][i]);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        try
        {
            FileStream writefile = new FileStream(outputFile, FileMode.Create);
            hssfworkbook.Write(writefile);
            writefile.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    #endregion

    public static int GetSheetNumber(string outputFile)
    {
        int number = 0;
        try
        {
            FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);

            HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
            number = hssfworkbook.NumberOfSheets;

        }
        catch (Exception exception)
        {
            throw exception;
        }
        return number;
    }

    public static ArrayList GetSheetName(string outputFile)
    {
        ArrayList arrayList = new ArrayList();
        try
        {
            FileStream readfile = new FileStream(outputFile, FileMode.Open, FileAccess.Read);

            HSSFWorkbook hssfworkbook = new HSSFWorkbook(readfile);
            for (int i = 0; i < hssfworkbook.NumberOfSheets; i++)
            {
                arrayList.Add(hssfworkbook.GetSheetName(i));
            }
        }
        catch (Exception exception)
        {
            throw exception;
        }
        return arrayList;
    }

    public static bool isNumeric(String message, out double result)
    {
        Regex rex = new Regex(@"^[-]?\d+[.]?\d*$");
        result = -1;
        if (rex.IsMatch(message))
        {
            result = double.Parse(message);
            return true;
        }
        else
            return false;

    }

    /// <summary>
    /// DataTable导出到Excel的MemoryStream                                                                      第二步
    /// </summary>
    /// <param name="dtSource">源DataTable</param>
    /// <param name="strHeaderText">表头文本</param>
    public static MemoryStream Export(DataTable dtSource, string strHeaderText)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.CreateSheet() as HSSFSheet;

        #region 右击文件 属性信息
        {
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI";
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Author = "文件作者信息"; //填加xls文件作者信息
            si.ApplicationName = "创建程序信息"; //填加xls文件创建程序信息
            si.LastAuthor = "最后保存者信息"; //填加xls文件最后保存者信息
            si.Comments = "作者信息"; //填加xls文件作者信息
            si.Title = "标题信息"; //填加xls文件标题信息
            si.Subject = "主题信息";//填加文件主题信息

            si.CreateDateTime = DateTime.Now;
            workbook.SummaryInformation = si;
        }
        #endregion

        HSSFCellStyle dateStyle = workbook.CreateCellStyle() as HSSFCellStyle;
        HSSFDataFormat format = workbook.CreateDataFormat() as HSSFDataFormat;
        dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

        //取得列宽
        int[] arrColWidth = new int[dtSource.Columns.Count];
        foreach (DataColumn item in dtSource.Columns)
        {
            arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
        }
        for (int i = 0; i < dtSource.Rows.Count; i++)
        {
            for (int j = 0; j < dtSource.Columns.Count; j++)
            {
                int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                if (intTemp > arrColWidth[j])
                {
                    arrColWidth[j] = intTemp;
                }
            }
        }
        int rowIndex = 0;
        foreach (DataRow row in dtSource.Rows)
        {
            #region 新建表，填充表头，填充列头，样式
            if (rowIndex == 65535 || rowIndex == 0)
            {
                if (rowIndex != 0)
                {
                    sheet = workbook.CreateSheet() as HSSFSheet;
                }

                #region 表头及样式
                {
                    if (string.IsNullOrEmpty(strHeaderText))
                    {
                        HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);
                        HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                        //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                        HSSFFont font = workbook.CreateFont() as HSSFFont;
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        headerRow.GetCell(0).CellStyle = headStyle;
                        sheet.AddMergedRegion(new Region(0, 0, 0, dtSource.Columns.Count - 1));
                        //headerRow.Dispose();
                    }
                }
                #endregion

                #region 列头及样式
                {
                    HSSFRow headerRow = sheet.CreateRow(0) as HSSFRow;
                    HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                    //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                    HSSFFont font = workbook.CreateFont() as HSSFFont;
                    font.FontHeightInPoints = 10;
                    font.Boldweight = 700;
                    headStyle.SetFont(font);
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                        headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                        //设置列宽
                        sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                    }
                    //headerRow.Dispose();
                }
                #endregion

                rowIndex = 1;
            }
            #endregion


            #region 填充内容
            HSSFRow dataRow = sheet.CreateRow(rowIndex) as HSSFRow;
            foreach (DataColumn column in dtSource.Columns)
            {
                HSSFCell newCell = dataRow.CreateCell(column.Ordinal) as HSSFCell;

                string drValue = row[column].ToString();

                switch (column.DataType.ToString())
                {
                    case "System.String"://字符串类型
                        newCell.SetCellValue(drValue);
                        break;
                    case "System.DateTime"://日期类型
                        DateTime dateV;
                        DateTime.TryParse(drValue, out dateV);
                        newCell.SetCellValue(dateV);

                        newCell.CellStyle = dateStyle;//格式化显示
                        break;
                    case "System.Boolean"://布尔型
                        bool boolV = false;
                        bool.TryParse(drValue, out boolV);
                        newCell.SetCellValue(boolV);
                        break;
                    case "System.Int16"://整型
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Byte":
                        int intV = 0;
                        int.TryParse(drValue, out intV);
                        newCell.SetCellValue(intV);
                        break;
                    case "System.Single":
                    case "System.Decimal"://浮点型
                    case "System.Double":
                        double doubV = 0;
                        double.TryParse(drValue, out doubV);
                        newCell.SetCellValue(doubV);
                        break;
                    case "System.DBNull"://空值处理
                        newCell.SetCellValue("");
                        break;
                    default:
                        newCell.SetCellValue(drValue);
                        break;
                }
            }
            #endregion

            rowIndex++;
        }
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            //sheet.Dispose();
            //workbook.Dispose();//一般只用写这一个就OK了，他会遍历并释放所有资源，但当前版本有问题所以只释放sheet
            return ms;
        }
    }

    /// <summary>
    /// 由DataSet导出Excel
    /// </summary>
    /// <param name="sourceTable">要导出数据的DataTable</param>
    /// <param name="sheetName">工作表名称</param>
    /// <returns>Excel工作表</returns>
    private static MemoryStream ExportDataSetToExcel(DataSet sourceDs, string sheetName)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();
        MemoryStream ms = new MemoryStream();
        string[] sheetNames = sheetName.Split(',');
        for (int i = 0; i < sheetNames.Length; i++)
        {
            ISheet sheet = workbook.CreateSheet(sheetNames[i]);

            #region 列头
            IRow headerRow = sheet.CreateRow(0);
            HSSFCellStyle headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
            HSSFFont font = workbook.CreateFont() as HSSFFont;
            font.FontHeightInPoints = 10;
            font.Boldweight = 700;
            headStyle.SetFont(font);

            //取得列宽
            int[] arrColWidth = new int[sourceDs.Tables[i].Columns.Count];
            foreach (DataColumn item in sourceDs.Tables[i].Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }

            // 处理列头
            foreach (DataColumn column in sourceDs.Tables[i].Columns)
            {
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                //设置列宽
                sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);

            }
            #endregion

            #region 填充值
            int rowIndex = 1;
            foreach (DataRow row in sourceDs.Tables[i].Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in sourceDs.Tables[i].Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }
                rowIndex++;
            }
            #endregion
        }
        workbook.Write(ms);
        ms.Flush();
        ms.Position = 0;
        workbook = null;
        return ms;
    }

    /// <summary>
    /// 验证导入的Excel是否有数据
    /// </summary>
    /// <param name="excelFileStream"></param>
    /// <returns></returns>
    public static bool HasData(Stream excelFileStream)
    {
        using (excelFileStream)
        {
            IWorkbook workBook = new HSSFWorkbook(excelFileStream);
            if (workBook.NumberOfSheets > 0)
            {
                ISheet sheet = workBook.GetSheetAt(0);
                return sheet.PhysicalNumberOfRows > 0;
            }
        }
        return false;
    }
    #endregion
}