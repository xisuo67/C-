using ICSharpCode.SharpZipLib.Zip;
using Models;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Tools
{
    /// <summary>
    /// NPOI操作EXECL帮助类
    /// </summary>
    public class NPOIHelper
    {
        /// <summary>
        /// EXECL最大列宽
        /// </summary>
        public static readonly int MAX_COLUMN_WIDTH = 100 * 256;

        /// <summary>
        /// 最大行索引
        /// </summary>
        private static readonly int MAX_ROW_INDEX = 65530;

        /// <summary>
        /// 生成EXECL文件，通过读取DataTable和列头映射信息
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="excelInfo">Excel导出信息</param>
        /// <param name="guid">导出Excel任务的标识，以此为key,在Session中更新进度</param>
        /// <returns>文件流</returns>
        public static MemoryStream Export(DataTable dt, ExcelInfo excelInfo, string guid = null)
        {
            if (dt == null || excelInfo == null || excelInfo.ColumnInfoList == null)
            {
                throw new ArgumentNullException();
            }
            bool isMoreHeader = excelInfo.GroupHeader != null && excelInfo.GroupHeader.Count > 0;
            if (isMoreHeader)
            {
                return CreateMoreHeaderXls(dt, excelInfo, guid);
            }
            int rowHeight = 20;
            List<ColumnInfo> ColumnInfoList = excelInfo.ColumnInfoList;
            //每个标签页最多行数
            int sheetRow = 65536;
            HSSFWorkbook workbook = new HSSFWorkbook();

            //文本样式
            ICellStyle centerStyle = workbook.CreateCellStyle();
            centerStyle.VerticalAlignment = VerticalAlignment.CENTER;
            centerStyle.Alignment = HorizontalAlignment.CENTER;

            ICellStyle leftStyle = workbook.CreateCellStyle();
            leftStyle.VerticalAlignment = VerticalAlignment.CENTER;
            leftStyle.Alignment = HorizontalAlignment.LEFT;

            ICellStyle rightStyle = workbook.CreateCellStyle();
            rightStyle.VerticalAlignment = VerticalAlignment.CENTER;
            rightStyle.Alignment = HorizontalAlignment.RIGHT;

            //寻找列头和DataTable之间映射关系
            foreach (DataColumn col in dt.Columns)
            {
                ColumnInfo info = ColumnInfoList.FirstOrDefault<ColumnInfo>(e => e.Field.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase));
                if (info != null)
                {
                    switch (info.Align.ToLower())
                    {
                        case "left":
                            info.Style = leftStyle;
                            break;
                        case "center":
                            info.Style = centerStyle;
                            break;
                        case "right":
                            info.Style = rightStyle;
                            break;
                    }
                    info.IsMapDT = true;
                }
            }

            int sheetNum = (int)Math.Ceiling(dt.Rows.Count * 1.0 / MAX_ROW_INDEX);
            int completeCount = 0;
            int total = dt.Rows.Count;
            int curprogress = 30;//查询数据已经花了30%

            //超链接字体颜色
            IFont blueFont = workbook.CreateFont();
            blueFont.Color = HSSFColor.BLUE.index;

            //最多生成5个标签页的数据
            sheetNum = sheetNum > 3 ? 3 : (sheetNum == 0 ? 1 : sheetNum);
            ICell cell = null;
            object cellValue = null;

            //标题头索引
            int headIndex = string.IsNullOrEmpty(excelInfo.Remark) ? 0 : 1;
            for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++)
            {
                ISheet sheet = workbook.CreateSheet();
                sheet.CreateFreezePane(0, headIndex + 1, 0, headIndex + 1);

                if (headIndex > 0)
                {
                    //输出备注行
                    IRow RemarkRow = sheet.CreateRow(0);
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, ColumnInfoList.Count - 1));
                    ICell rcell = RemarkRow.CreateCell(0);
                    ICellStyle remarkStyle = workbook.CreateCellStyle();
                    remarkStyle.WrapText = true;
                    remarkStyle.VerticalAlignment = VerticalAlignment.TOP;
                    remarkStyle.Alignment = HorizontalAlignment.LEFT;
                    IFont rfont = workbook.CreateFont();
                    rfont.FontHeightInPoints = 12;
                    remarkStyle.SetFont(rfont);
                    rcell.CellStyle = remarkStyle;
                    RemarkRow.HeightInPoints = rowHeight * 5;
                    rcell.SetCellValue(excelInfo.Remark);
                }


                //输出表头
                IRow headerRow = sheet.CreateRow(headIndex);
                //设置行高
                headerRow.HeightInPoints = rowHeight;
                //首行样式
                ICellStyle HeaderStyle = workbook.CreateCellStyle();
                HeaderStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                HeaderStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
                IFont font = workbook.CreateFont();
                font.Boldweight = short.MaxValue;
                HeaderStyle.SetFont(font);
                HeaderStyle.VerticalAlignment = VerticalAlignment.CENTER; ;
                HeaderStyle.Alignment = HorizontalAlignment.CENTER;


                //输出表头信息 并设置表头样式
                int i = 0;
                foreach (var data in ColumnInfoList)
                {
                    cell = headerRow.CreateCell(i);
                    cell.SetCellValue(data.Header.Trim());
                    cell.CellStyle = HeaderStyle;
                    i++;
                }

                //开始循环所有行
                int iRow = 1 + headIndex;

                int startRow = sheetIndex * (sheetRow - 1);
                int endRow = (sheetIndex + 1) * (sheetRow - 1);
                endRow = endRow <= dt.Rows.Count ? endRow : dt.Rows.Count;

                for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
                {
                    IRow row = sheet.CreateRow(iRow);
                    row.HeightInPoints = rowHeight;
                    i = 0;
                    foreach (var item in ColumnInfoList)
                    {
                        cell = row.CreateCell(i);
                        if (item.IsMapDT)
                        {
                            cellValue = dt.Rows[rowIndex][item.Field];
                            cell.SetCellValue(cellValue != DBNull.Value ? cellValue.ToString() : string.Empty);
                            cell.CellStyle = item.Style;

                            if (item.IsLink)
                            {
                                cellValue = dt.Rows[rowIndex][item.Field + "Link"];
                                if (cellValue != DBNull.Value && cellValue != null)
                                {
                                    //建一个HSSFHyperlink实体，指明链接类型为URL（这里是枚举，可以根据需求自行更改）  
                                    HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.URL);
                                    //给HSSFHyperlink的地址赋值 ，默认为该列加上Link
                                    link.Address = cellValue.ToString();
                                    cell.Hyperlink = link;
                                    cell.CellStyle.SetFont(blueFont);
                                }
                            }
                        }
                        i++;
                    }

                    //记录进度
                    completeCount++;
                    int temp = 30 + (int)(completeCount * 65 / total);
                    if (temp > curprogress)
                    {
                        //当temp >curprogress 才写Session,避免无用的写Session
                        curprogress = temp;
                        if (!string.IsNullOrEmpty(guid))
                        {
                            HttpContext.Current.Session[guid] = curprogress;
                        }
                    }

                    iRow++;
                }

                //自适应列宽度
                for (int j = 0; j < ColumnInfoList.Count; j++)
                {
                    sheet.AutoSizeColumn(j);
                    int width = sheet.GetColumnWidth(j) + 2560;
                    sheet.SetColumnWidth(j, width > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : width);
                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            //记录进度
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 99;
            }

            return ms;
        }

        /// <summary>
        /// 生成EXECL文件，通过读取DataTable和列头映射信息
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="excelInfo">Excel导出信息</param>
        /// <param name="guid">导出Excel任务的标识，以此为key,在Session中更新进度</param>
        /// <returns>文件流</returns>
        public static MemoryStream CreateMoreHeaderXls(DataTable dt, ExcelInfo excelInfo, string guid = null)
        {
            int rowHeight = 20;
            List<ColumnInfo> ColumnInfoList = excelInfo.ColumnInfoList;

            HSSFWorkbook workbook = new HSSFWorkbook();

            //文本样式
            ICellStyle centerStyle = workbook.CreateCellStyle();
            centerStyle.VerticalAlignment = VerticalAlignment.CENTER;
            centerStyle.Alignment = HorizontalAlignment.CENTER;

            ICellStyle leftStyle = workbook.CreateCellStyle();
            leftStyle.VerticalAlignment = VerticalAlignment.CENTER;
            leftStyle.Alignment = HorizontalAlignment.LEFT;

            ICellStyle rightStyle = workbook.CreateCellStyle();
            rightStyle.VerticalAlignment = VerticalAlignment.CENTER;
            rightStyle.Alignment = HorizontalAlignment.RIGHT;

            //首行样式
            ICellStyle HeaderStyle = workbook.CreateCellStyle();
            HeaderStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            HeaderStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LIGHT_CORNFLOWER_BLUE.index;
            HeaderStyle.BorderTop = BorderStyle.THIN;
            HeaderStyle.BorderLeft = BorderStyle.THIN;
            HeaderStyle.BorderRight = BorderStyle.THIN;
            HeaderStyle.BorderBottom = BorderStyle.THIN;
            HeaderStyle.TopBorderColor = NPOI.HSSF.Util.HSSFColor.BLACK.index;
            HeaderStyle.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.BLACK.index;
            HeaderStyle.RightBorderColor = NPOI.HSSF.Util.HSSFColor.BLACK.index;
            HeaderStyle.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.BLACK.index;
            IFont font = workbook.CreateFont();
            font.Boldweight = short.MaxValue;
            HeaderStyle.SetFont(font);
            HeaderStyle.VerticalAlignment = VerticalAlignment.CENTER;
            HeaderStyle.Alignment = HorizontalAlignment.CENTER;

            Dictionary<string, int> dictGroupMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Dictionary<ColumnInfo, bool> dictColumn = new Dictionary<ColumnInfo, bool>();

            //寻找列头和DataTable之间映射关系

            foreach (DataColumn col in dt.Columns)
            {
                ColumnInfo info = ColumnInfoList.FirstOrDefault<ColumnInfo>(e => e.Field.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase));
                if (info != null)
                {
                    dictColumn[info] = col.DataType == typeof(int) || col.DataType == typeof(float) || col.DataType == typeof(double)
                        || col.DataType == typeof(long);
                    switch (info.Align.ToLower())
                    {
                        case "left":
                            info.Style = leftStyle;
                            break;
                        case "center":
                            info.Style = centerStyle;
                            break;
                        case "right":
                            info.Style = rightStyle;
                            break;
                    }
                    info.IsMapDT = true;
                }
            }
            int index = 0;
            foreach (var item in ColumnInfoList)
            {
                if (excelInfo.GroupHeader.FirstOrDefault(e => e.StartColumnName.Equals(item.Field, StringComparison.OrdinalIgnoreCase)) != null)
                {
                    dictGroupMap[item.Field] = index;
                }
                index++;
            }

            List<int> listColumnIndex = new List<int>(ColumnInfoList.Count);
            for (int i = 0; i < ColumnInfoList.Count; i++)
            {
                listColumnIndex.Add(i);
            }
            foreach (var item in excelInfo.GroupHeader)
            {
                int startCol = dictGroupMap[item.StartColumnName];
                int lastCol = startCol + item.NumberOfColumns - 1;
                for (int j = startCol; j <= lastCol; j++)
                {
                    listColumnIndex.Remove(j);
                }
            }

            int sheetNum = (int)Math.Ceiling(dt.Rows.Count * 1.0 / MAX_ROW_INDEX);
            int completeCount = 0;
            int total = dt.Rows.Count;
            int curprogress = 30;//查询数据已经花了30%

            //超链接字体颜色
            IFont blueFont = workbook.CreateFont();
            blueFont.Color = HSSFColor.BLUE.index;

            //最多生成5个标签页的数据
            sheetNum = sheetNum > 3 ? 3 : (sheetNum == 0 ? 1 : sheetNum);
            ICell cell = null;
            object cellValue = null;

            //标题头索引
            int headIndex = string.IsNullOrEmpty(excelInfo.Remark) ? 0 : 1;
            for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++)
            {
                ISheet sheet = workbook.CreateSheet();

                if (headIndex > 0)
                {
                    //输出备注行
                    IRow RemarkRow = sheet.CreateRow(0);
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, ColumnInfoList.Count - 1));
                    ICell rcell = RemarkRow.CreateCell(0);
                    ICellStyle remarkStyle = workbook.CreateCellStyle();
                    remarkStyle.WrapText = true;
                    remarkStyle.VerticalAlignment = VerticalAlignment.TOP;
                    remarkStyle.Alignment = HorizontalAlignment.LEFT;
                    IFont rfont = workbook.CreateFont();
                    rfont.FontHeightInPoints = 12;
                    remarkStyle.SetFont(rfont);
                    rcell.CellStyle = remarkStyle;
                    RemarkRow.HeightInPoints = rowHeight * 5;
                    rcell.SetCellValue(excelInfo.Remark);
                }

                //输出表头
                IRow firstHeaderRow = sheet.CreateRow(headIndex);
                IRow secondHeaderRow = sheet.CreateRow(headIndex + 1);

                //设置行高
                firstHeaderRow.HeightInPoints = rowHeight;
                secondHeaderRow.HeightInPoints = rowHeight;

                //输出表头信息 并设置表头样式
                int i = 0, groupIndex = 0;
                MoreHeader groupheader = null;
                foreach (var data in ColumnInfoList)
                {
                    cell = secondHeaderRow.CreateCell(i);
                    cell.SetCellValue(data.Header.Trim());
                    cell.CellStyle = HeaderStyle;

                    cell = firstHeaderRow.CreateCell(i);
                    cell.SetCellValue(data.Header.Trim());
                    cell.CellStyle = HeaderStyle;

                    groupheader = excelInfo.GroupHeader[groupIndex];
                    if (groupheader.StartColumnName.Equals(data.Field, StringComparison.CurrentCultureIgnoreCase))
                    {
                        cell.SetCellValue(groupheader.TitleText.Trim());
                        groupIndex++;
                        if (groupIndex >= excelInfo.GroupHeader.Count)
                        {
                            groupIndex--;
                        }
                    }
                    i++;
                }
                foreach (var item in listColumnIndex)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(headIndex, headIndex + 1, item, item));
                }
                int startCol, lastCol;
                foreach (var item in excelInfo.GroupHeader)
                {
                    startCol = dictGroupMap[item.StartColumnName];
                    lastCol = startCol + item.NumberOfColumns - 1;
                    sheet.AddMergedRegion(new CellRangeAddress(headIndex, headIndex, startCol, lastCol));
                }
                //冻结列 行
                sheet.CreateFreezePane(excelInfo.FixColumns, headIndex + 2, excelInfo.FixColumns, headIndex + 2);

                //开始循环所有行
                int iRow = 2 + headIndex;

                int startRow = sheetIndex * (MAX_ROW_INDEX - 1);
                int endRow = (sheetIndex + 1) * (MAX_ROW_INDEX - 1);
                endRow = endRow <= dt.Rows.Count ? endRow : dt.Rows.Count;

                for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
                {
                    IRow row = sheet.CreateRow(iRow);
                    row.HeightInPoints = rowHeight;
                    i = 0;
                    foreach (var item in ColumnInfoList)
                    {
                        cell = row.CreateCell(i);
                        if (item.IsMapDT)
                        {
                            cellValue = dt.Rows[rowIndex][item.Field];
                            cell.SetCellValue((cellValue != null && cellValue != DBNull.Value) ? cellValue.ToString() : (dictColumn[item] ? "--" : string.Empty));
                            cell.CellStyle = item.Style;

                            if (item.IsLink)
                            {
                                cellValue = dt.Rows[rowIndex][item.Field + "Link"];
                                if (cellValue != DBNull.Value && cellValue != null)
                                {
                                    //建一个HSSFHyperlink实体，指明链接类型为URL（这里是枚举，可以根据需求自行更改）  
                                    HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.URL);
                                    //给HSSFHyperlink的地址赋值 ，默认为该列加上Link
                                    link.Address = cellValue.ToString();
                                    cell.Hyperlink = link;
                                    cell.CellStyle.SetFont(blueFont);
                                }
                            }
                        }
                        i++;
                    }

                    //记录进度
                    completeCount++;
                    int temp = 30 + (int)(completeCount * 65 / total);
                    if (temp > curprogress)
                    {
                        //当temp >curprogress 才写Session,避免无用的写Session
                        curprogress = temp;
                        if (!string.IsNullOrEmpty(guid))
                        {
                            HttpContext.Current.Session[guid] = curprogress;
                        }
                    }

                    iRow++;
                }

                //自适应列宽度
                for (int j = 0; j < ColumnInfoList.Count; j++)
                {
                    sheet.AutoSizeColumn(j);
                    int width = sheet.GetColumnWidth(j) + 2560;
                    sheet.SetColumnWidth(j, width > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : width);
                }
            }

            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);
            //记录进度
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 99;
            }

            return ms;
        }

        /// <summary>
        /// 通过多标签页数据将数据写入响应流
        /// </summary>
        /// <param name="multiSheet">多标签页数据</param>
        /// <param name="Response">响应流</param>
        public static void ExportToMutilSheet(ExportMultiSheet multiSheet, HttpResponseBase Response, string guid = "")
        {
            ExportToMutilSheet(multiSheet, Response.OutputStream, guid);

            Response.Headers.Add("Content-Disposition", string.Format("attachment;filename={0}", Path.GetFileName(multiSheet.FileName)));
            Response.ContentType = MimeHelper.GetMineType(multiSheet.FileName);

            Response.OutputStream.Flush();
            Response.End();
        }

        /// <summary>
        /// 导出数据到多个Excel标签页中
        /// </summary>
        /// <param name="multiSheet">多标签页信息</param>
        /// <param name="s">输出流</param>
        /// <param name="guid">进度控制</param>
        public static void ExportToMutilSheet(ExportMultiSheet multiSheet, Stream s, string guid = "")
        {
            if (multiSheet == null || multiSheet.ListSheet == null || multiSheet.ListSheet.Count == 0)
            {
                throw new ArgumentNullException();
            }
            string fileExt = ".xls";
            if (string.IsNullOrEmpty(multiSheet.FileName))
            {
                multiSheet.FileName = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + fileExt;
            }
            else
            {
                if (!multiSheet.FileName.EndsWith(fileExt))
                {
                    multiSheet.FileName = multiSheet.FileName + fileExt;
                }
            }
            //进度信息
            ProgressResult pr = null;
            bool isAddProcess = false;
            if (!string.IsNullOrEmpty(guid))
            {
                pr = HttpContext.Current.Session[guid] as ProgressResult;
                if (pr != null && pr.Value < 90)
                {
                    isAddProcess = true;
                }
            }

            IWorkbook workbook = new HSSFWorkbook();
            //文本样式
            ICellStyle centerStyle = workbook.CreateCellStyle();
            centerStyle.VerticalAlignment = VerticalAlignment.CENTER;
            centerStyle.Alignment = HorizontalAlignment.CENTER;

            ICellStyle leftStyle = workbook.CreateCellStyle();
            leftStyle.VerticalAlignment = VerticalAlignment.CENTER;
            leftStyle.Alignment = HorizontalAlignment.LEFT;

            ICellStyle rightStyle = workbook.CreateCellStyle();
            rightStyle.VerticalAlignment = VerticalAlignment.CENTER;
            rightStyle.Alignment = HorizontalAlignment.RIGHT;


            //超链接字体颜色
            IFont blueFont = workbook.CreateFont();
            blueFont.Color = HSSFColor.BLUE.index;
            ICellStyle leftStyleLink = workbook.CreateCellStyle();
            leftStyleLink.VerticalAlignment = VerticalAlignment.CENTER;
            leftStyleLink.Alignment = HorizontalAlignment.LEFT;
            leftStyleLink.SetFont(blueFont);

            ISheet sheet = null;
            DataTable dt = null;
            DataRow dr = null;
            List<ColumnInfo> ColumnInfoList = null;
            int rowHeight = 20;
            object cellValue = null;
            IRow row = null;
            ICell cell = null;

            //每写入100条数据进度更新一次
            int totalData = multiSheet.ListSheet.Sum(e => e.Data.Rows.Count);
            int complete = 0, frequency = 1;
            if (isAddProcess)
            {
                if (totalData <= (90 - pr.Value))
                {
                    complete = 90 - pr.Value;
                    frequency = totalData;
                }
                else
                {
                    complete = 1;
                    frequency = (int)Math.Ceiling(totalData * 1.0 / (90 - pr.Value));
                }
            }
            //写入总数
            int writeTotal = 0;
            foreach (ExportSheetInfo exportSheetInfo in multiSheet.ListSheet)
            {
                dt = exportSheetInfo.Data;
                ColumnInfoList = exportSheetInfo.ColumnInfoList;
                //寻找列头和DataTable之间映射关系
                foreach (DataColumn col in dt.Columns)
                {
                    ColumnInfo info = ColumnInfoList.FirstOrDefault<ColumnInfo>(e => e.Field.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase));
                    if (info != null)
                    {
                        info.Align = info.Align ?? "left";
                        switch (info.Align.ToLower())
                        {
                            case "left":
                                info.Style = leftStyle;
                                break;
                            case "center":
                                info.Style = centerStyle;
                                break;
                            case "right":
                                info.Style = rightStyle;
                                break;
                        }
                        info.IsMapDT = true;
                    }
                }
                //标题头索引
                int headIndex = string.IsNullOrEmpty(exportSheetInfo.Remark) ? 0 : 1;
                int total = dt.Rows.Count;
                int sheetNum = (int)Math.Ceiling(total * 1.0 / (MAX_ROW_INDEX - headIndex - 1));
                int drIndex = 0;

                for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++)
                {
                    string sheetName = string.IsNullOrEmpty(exportSheetInfo.SheetName) ? "Sheet " + workbook.NumberOfSheets : (sheetNum > 1 ? (exportSheetInfo.SheetName + sheetIndex) : exportSheetInfo.SheetName);
                    sheet = workbook.CreateSheet(sheetName);

                    sheet.CreateFreezePane(0, headIndex + 1, 0, headIndex + 1);
                    if (headIndex > 0)
                    {
                        //输出备注行
                        IRow RemarkRow = sheet.CreateRow(0);
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, ColumnInfoList.Count - 1));
                        ICell rcell = RemarkRow.CreateCell(0);
                        ICellStyle remarkStyle = workbook.CreateCellStyle();
                        remarkStyle.WrapText = true;
                        remarkStyle.VerticalAlignment = VerticalAlignment.TOP;
                        remarkStyle.Alignment = HorizontalAlignment.LEFT;
                        IFont rfont = workbook.CreateFont();
                        rfont.FontHeightInPoints = 12;
                        remarkStyle.SetFont(rfont);
                        rcell.CellStyle = remarkStyle;
                        RemarkRow.HeightInPoints = rowHeight * 5;
                        rcell.SetCellValue(exportSheetInfo.Remark);
                    }

                    //输出表头
                    IRow headerRow = sheet.CreateRow(headIndex);
                    //设置行高
                    headerRow.HeightInPoints = rowHeight;
                    //首行样式
                    ICellStyle HeaderStyle = workbook.CreateCellStyle();
                    HeaderStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
                    HeaderStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index;
                    IFont font = workbook.CreateFont();
                    font.Boldweight = short.MaxValue;
                    HeaderStyle.SetFont(font);
                    HeaderStyle.VerticalAlignment = VerticalAlignment.CENTER; ;
                    HeaderStyle.Alignment = HorizontalAlignment.CENTER;

                    //输出表头信息 并设置表头样式
                    int i = 0;
                    foreach (var data in ColumnInfoList)
                    {
                        cell = headerRow.CreateCell(i);
                        cell.SetCellValue(data.Header.Trim());
                        cell.CellStyle = HeaderStyle;
                        i++;
                    }

                    //开始循环所有行
                    int iRow = 1 + headIndex;
                    int startRow = iRow;

                    while (startRow < MAX_ROW_INDEX && drIndex < total)
                    {
                        row = sheet.CreateRow(startRow);
                        row.HeightInPoints = rowHeight;
                        i = 0;
                        dr = dt.Rows[drIndex];
                        foreach (var item in ColumnInfoList)
                        {
                            cell = row.CreateCell(i);
                            if (item.IsMapDT)
                            {
                                cellValue = dr[item.Field];
                                cell.SetCellValue(cellValue != DBNull.Value ? cellValue.ToString() : string.Empty);
                                cell.CellStyle = item.Style;

                                if (item.IsLink)
                                {
                                    cellValue = dr[item.Field + "Link"];
                                    if (cellValue != DBNull.Value && cellValue != null)
                                    {
                                        //建一个HSSFHyperlink实体，指明链接类型为URL（这里是枚举，可以根据需求自行更改）  
                                        HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.URL);
                                        //给HSSFHyperlink的地址赋值 ，默认为该列加上Link
                                        link.Address = cellValue.ToString();
                                        cell.Hyperlink = link;
                                        cell.CellStyle = leftStyleLink;
                                    }
                                }
                            }
                            i++;
                        }
                        drIndex++;
                        startRow++;
                        writeTotal++;

                        if (isAddProcess && writeTotal % frequency == 0)
                        {
                            pr.Value += complete;
                            HttpContext.Current.Session[guid] = pr;
                        }
                    }
                    //自适应列宽度
                    for (int j = 0; j < ColumnInfoList.Count; j++)
                    {
                        sheet.AutoSizeColumn(j);
                        int width = sheet.GetColumnWidth(j) + 2560;
                        sheet.SetColumnWidth(j, width > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : width);
                    }
                }
            }

            if (s is ZipOutputStream)
            {
                ZipOutputStream zs = s as ZipOutputStream;
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Write(ms);
                    byte[] m_buffer = ms.GetBuffer();
                    zs.Write(m_buffer, 0, m_buffer.Length);
                }
            }
            else
            {
                workbook.Write(s);
            }
            if (isAddProcess)
            {
                pr.Value += 5;
                HttpContext.Current.Session[guid] = pr;
            }
        }

        /// <summary>
        /// 从EXECL中读取数据 转换成DataTable
        /// 每个sheet页对应一个DataTable
        /// </summary>
        /// <param name="xlsxFile">文件路径</param>
        /// <returns></returns>
        public static List<DataTable> GetDataTablesFrom(string xlsxFile)
        {
            if (!File.Exists(xlsxFile))
                throw new FileNotFoundException("文件不存在");

            List<DataTable> result = new List<DataTable>();
            Stream stream = new MemoryStream(File.ReadAllBytes(xlsxFile));
            IWorkbook workbook = new HSSFWorkbook(stream);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                DataTable dt = new DataTable();
                ISheet sheet = workbook.GetSheetAt(i);
                IRow headerRow = sheet.GetRow(0);

                int cellCount = headerRow.LastCellNum;
                for (int j = headerRow.FirstCellNum; j < cellCount; j++)
                {
                    DataColumn column = new DataColumn(headerRow.GetCell(j).StringCellValue);
                    dt.Columns.Add(column);
                }

                int rowCount = sheet.LastRowNum;
                for (int a = (sheet.FirstRowNum + 1); a < rowCount; a++)
                {
                    IRow row = sheet.GetRow(a);
                    if (row == null) continue;

                    DataRow dr = dt.NewRow();
                    for (int b = row.FirstCellNum; b < cellCount; b++)
                    {
                        if (row.GetCell(b) == null) continue;
                        dr[b] = row.GetCell(b).ToString();
                    }

                    dt.Rows.Add(dr);
                }
                result.Add(dt);
            }
            stream.Close();

            return result;
        }

        /// <summary>
        /// 获取第一个Sheet
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>Sheet</returns>
        public static ISheet GetFirstSheet(string filePath)
        {
            using (Stream stream = new MemoryStream(File.ReadAllBytes(filePath)))
            {
                IWorkbook workbook = new HSSFWorkbook(stream);
                if (workbook.NumberOfSheets > 0)
                {
                    return workbook.GetSheetAt(0);
                }
            }
            return null;
        }

        #region "Excel模版数据读取相关"

        /// <summary>
        /// 通过输入流初始化workbook
        /// </summary>
        /// <param name="ins">输入流</param>
        /// <returns>workbook对象</returns>
        public static IWorkbook InitWorkBook(Stream ins)
        {
            return new HSSFWorkbook(ins);
        }

        /// <summary>
        /// 从excel第一个sheet中读取数据
        /// </summary>
        /// <param name="ins">输入流</param>
        /// <param name="headRowIndex">标题行索引 默认为第6行</param>
        /// <param name="fSheet">第一个sheet</param>
        /// <returns>DataTable</returns>
        public static DataTable GetDataFromExcel(Stream ins, out ISheet fSheet, int headRowIndex = 5)
        {
            IWorkbook workbook = InitWorkBook(ins);
            fSheet = null;
            DataTable dt = new DataTable();
            if (workbook.NumberOfSheets > 0)
            {
                fSheet = workbook.GetSheetAt(0);
                if (fSheet.LastRowNum < headRowIndex)
                {
                    throw new ArgumentException("Excel模版错误,标题行索引大于总行数");
                }

                //读取标题行
                IRow row = null;
                ICell cell = null;

                row = fSheet.GetRow(headRowIndex);
                object objColumnName = null;
                for (int i = 0, length = row.LastCellNum; i < length; i++)
                {
                    cell = row.GetCell(i);
                    if (cell == null)
                    {
                        continue;
                    }
                    objColumnName = GetCellVale(cell);
                    if (objColumnName != null)
                    {
                        try
                        {
                            dt.Columns.Add(objColumnName.ToString().Trim());
                        }
                        catch (Exception e)
                        {
                            throw new Exception("上传文件格式与下载模板格式不符");
                        }

                    }
                    else
                    {
                        dt.Columns.Add("");
                    }
                }

                //读取数据行
                object[] entityValues = null;
                int columnCount = dt.Columns.Count;

                for (int i = headRowIndex + 1, length = fSheet.LastRowNum; i < length; i++)
                {
                    row = fSheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    entityValues = new object[columnCount];
                    //用于判断是否为空行
                    bool isHasData = false;
                    int dataColumnLength = row.LastCellNum < columnCount ? row.LastCellNum : columnCount;
                    for (int j = 0; j < dataColumnLength; j++)
                    {
                        cell = row.GetCell(j);
                        if (cell == null)
                        {
                            continue;
                        }
                        entityValues[j] = GetCellVale(cell);
                        if (!isHasData && j < columnCount && entityValues[j] != null)
                        {
                            isHasData = true;
                        }
                    }
                    if (isHasData)
                    {
                        dt.Rows.Add(entityValues);
                    }
                }
            }
            return dt;
        }


        /// <summary>
        /// 从excel中所有sheet中读取数据
        /// </summary>
        /// <param name="ins">输入流</param>
        /// <param name="headRowIndex">标题行索引 默认为第6行</param>
        /// <param name="listSheet">所有有数据的Sheet</param>
        /// <returns>DataTable</returns>
        public static DataSet GetDataSetFromExcel(Stream ins, List<string> sheetNames, int headRowIndex, out List<ISheet> listSheet)
        {
            IWorkbook workbook = InitWorkBook(ins);
            DataSet ds = new DataSet();
            List<ISheet> sheets = new List<ISheet>();
            if (workbook.NumberOfSheets > 0)
            {
                //读取标题行
                IRow row = null;
                ICell cell = null;
                ISheet fSheet = null;
                DataTable dt = null;
                foreach (string sheetName in sheetNames)
                {
                    fSheet = workbook.GetSheet(sheetName);
                    if (fSheet == null || fSheet.LastRowNum < headRowIndex)
                    {
                        continue;
                    }
                    dt = new DataTable();
                    dt.TableName = sheetName;
                    row = fSheet.GetRow(headRowIndex);
                    object objColumnName = null;
                    for (int i = 0, length = row.LastCellNum; i < length; i++)
                    {
                        cell = row.GetCell(i);
                        if (cell == null)
                        {
                            continue;
                        }
                        objColumnName = GetCellVale(cell);
                        if (objColumnName != null)
                        {
                            dt.Columns.Add(objColumnName.ToString().Trim());
                        }
                        else
                        {
                            dt.Columns.Add("");
                        }
                    }

                    //读取数据行
                    object[] entityValues = null;
                    int columnCount = dt.Columns.Count;

                    for (int i = headRowIndex + 1, length = fSheet.LastRowNum; i < length; i++)
                    {
                        row = fSheet.GetRow(i);
                        if (row == null)
                        {
                            continue;
                        }
                        entityValues = new object[columnCount];
                        //用于判断是否为空行
                        bool isHasData = false;
                        int dataColumnLength = row.LastCellNum < columnCount ? row.LastCellNum : columnCount;
                        for (int j = 0; j < dataColumnLength; j++)
                        {
                            cell = row.GetCell(j);
                            if (cell == null)
                            {
                                continue;
                            }
                            entityValues[j] = GetCellVale(cell);
                            if (!isHasData && j < columnCount && entityValues[j] != null)
                            {
                                isHasData = true;
                            }
                        }
                        if (isHasData)
                        {
                            dt.Rows.Add(entityValues);
                        }
                    }
                    ds.Tables.Add(dt);
                    sheets.Add(fSheet);
                }
            }
            listSheet = sheets;
            return ds;
        }

        /// <summary>
        /// 设置excel模版错误信息
        /// </summary>
        /// <param name="sheet">数据标签</param>
        /// <param name="rowindex">错误信息显示行</param>
        /// <param name="msg">错误信息</param>
        public static void SetTemplateErrorMsg(ISheet sheet, int rowindex, string msg)
        {
            IRow row = sheet.GetRow(rowindex);
            row = sheet.CreateRow(rowindex);
            if (row != null && !string.IsNullOrEmpty(msg))
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowindex, rowindex, 0, row.LastCellNum));

                ICell cell = row.GetCell(0);
                if (cell == null)
                {
                    cell = row.CreateCell(0);
                }
                ICellStyle cellStyle = sheet.Workbook.CreateCellStyle();
                cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
                cellStyle.Alignment = HorizontalAlignment.LEFT;
                IFont font = sheet.Workbook.CreateFont();
                font.FontHeightInPoints = 12;
                font.Color = HSSFColor.RED.index;
                cellStyle.SetFont(font);
                cell.CellStyle = cellStyle;
                cell.SetCellValue(msg);
            }
        }

        /// <summary>
        /// 获取数据行的错误信息提示样式
        /// </summary>
        /// <returns>错误数据行样式</returns>
        public static ICellStyle GetErrorCellStyle(IWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            cellStyle.Alignment = HorizontalAlignment.LEFT;
            IFont font = wb.CreateFont();
            //font.FontHeightInPoints = 12;
            font.Color = HSSFColor.RED.index;
            cellStyle.SetFont(font);
            return cellStyle;
        }

        /// <summary>
        /// 获取标题行的错误信息提示样式
        /// </summary>
        /// <returns>错误标题行样式</returns>
        public static ICellStyle GetErrorHeadCellStyle(IWorkbook wb)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.CENTER;
            cellStyle.Alignment = HorizontalAlignment.CENTER;
            IFont font = wb.CreateFont();
            font.Boldweight = short.MaxValue;
            font.Color = HSSFColor.RED.index;
            cellStyle.SetFont(font);
            cellStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
            return cellStyle;
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>单元格值</returns>
        private static object GetCellVale(ICell cell)
        {
            object obj = null;
            switch (cell.CellType)
            {
                case CellType.NUMERIC:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        obj = cell.DateCellValue;
                    }
                    else
                    {
                        obj = cell.NumericCellValue;
                    }
                    break;
                case CellType.STRING:
                    if (string.IsNullOrEmpty(cell.StringCellValue))
                        obj = string.Empty;
                    else
                        obj = cell.StringCellValue.Trim();
                    break;
                case CellType.BOOLEAN:
                    obj = cell.BooleanCellValue;
                    break;
                case CellType.FORMULA:
                    obj = cell.CellFormula;
                    break;

            }
            return obj;
        }
        #endregion

        #region "设置下拉选项"
        /// <summary>
        /// 设置某些列的值只能输入预制的数据,显示下拉框
        /// </summary>
        /// <param name="sheet">要设置的sheet</param>
        /// <param name="textlist">下拉框显示的内容</param>
        /// <param name="firstRow">开始行</param>
        /// <param name="firstCol">开始列</param>
        /// <param name="ValidationData">是否验证数据只能从下拉框中选择 默认为true   为false时既可以选择也可以输入</param>
        /// <returns>设置好的sheet</returns>
        public static ISheet SetHSSFValidation(ISheet sheet,
                string[] textlist, int firstRow, int firstCol, bool ValidationData = true)
        {
            return SetHSSFValidation(sheet, textlist, firstRow, sheet.LastRowNum, firstCol, firstCol, ValidationData);
        }

        /// <summary>
        /// 设置某些列的值只能输入预制的数据,显示下拉框
        /// </summary>
        /// <param name="sheet">要设置的sheet</param>
        /// <param name="textlist">下拉框显示的内容</param>
        /// <param name="firstRow">开始行</param>
        /// <param name="endRow">结束行</param>
        /// <param name="firstCol">开始列</param>
        /// <param name="endCol">结束列</param>
        /// <param name="ValidationData">是否验证数据只能从下拉框中选择 默认为true   为false时既可以选择也可以输入</param>
        /// <returns>设置好的sheet</returns>
        public static ISheet SetHSSFValidation(ISheet sheet,
                string[] textlist, int firstRow, int endRow, int firstCol,
                int endCol, bool ValidationData = true)
        {
            IWorkbook workbook = sheet.Workbook;
            if (endRow > sheet.LastRowNum)
            {
                endRow = sheet.LastRowNum;
            }
            ISheet hidden = null;
            string hiddenSheetName = "hidden" + sheet.SheetName.GetFirstPinyin();
            int hIndex = workbook.GetSheetIndex(hiddenSheetName);
            if (hIndex < 0)
            {
                hidden = workbook.CreateSheet(hiddenSheetName);
                workbook.SetSheetHidden(sheet.Workbook.NumberOfSheets - 1, SheetState.HIDDEN);
            }
            else
            {
                hidden = workbook.GetSheetAt(hIndex);
            }
            if (textlist == null || textlist.Length == 0)
            {
                textlist = new string[] { "" };
            }
            IRow row = null;
            ICell cell = null;
            for (int i = 0, length = textlist.Length; i < length; i++)
            {
                row = hidden.GetRow(i);
                if (row == null)
                {
                    row = hidden.CreateRow(i);
                }
                cell = row.GetCell(firstCol);
                if (cell == null)
                {
                    cell = row.CreateCell(firstCol);
                }
                cell.SetCellValue(textlist[i]);
            }

            // 加载下拉列表内容  
            string nameCellKey = hiddenSheetName + firstCol;
            IName namedCell = workbook.GetName(nameCellKey);
            if (namedCell == null)
            {
                namedCell = workbook.CreateName();
                namedCell.NameName = nameCellKey;
                namedCell.RefersToFormula = string.Format("{0}!${1}$1:${1}${2}", hiddenSheetName, NumberToChar(firstCol + 1), textlist.Length);
            }
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(nameCellKey);

            // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列  
            CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);

            // 数据有效性对象  
            HSSFDataValidation validation = new HSSFDataValidation(regions, constraint);
            //// 取消弹出错误框
            validation.ShowErrorBox = ValidationData;
            sheet.AddValidationData(validation);
            return sheet;
        }
        #endregion

        #region "私有方法"
        /// 
        /// 把1,2,3,...,35,36转换成A,B,C,...,Y,Z
        /// 
        /// 要转换成字母的数字（数字范围在闭区间[1,36]）
        /// 
        private static string NumberToChar(int number)
        {
            if (1 <= number && 36 >= number)
            {
                int num = number + 64;
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)num };
                return asciiEncoding.GetString(btNumber);
            }
            return "A";
        }
        #endregion

    }
}
/// <summary>
/// NPOI拓展方法
/// </summary>
public static class NPOIExtend
{
    /// <summary>
    /// 获取RGB对应NPOI颜色值
    /// </summary>
    /// <param name="workbook">当前wb</param>
    /// <param name="R"></param>
    /// <param name="G"></param>
    /// <param name="B"></param>
    /// <returns></returns>
    public static short GetXLColor(this HSSFWorkbook workbook, int R, int G, int B)
    {
        short s = 0;
        HSSFPalette XlPalette = workbook.GetCustomPalette();
        HSSFColor XlColour = XlPalette.FindColor((byte)R, (byte)G, (byte)B);
        if (XlColour == null)
        {
            if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
            {
                if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 64)
                {
                    NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE = 64;
                    NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE += 1;
                    XlColour = XlPalette.AddColor((byte)R, (byte)G, (byte)B);
                }
                else
                {
                    XlColour = XlPalette.FindSimilarColor((byte)R, (byte)G, (byte)B);
                }

                s = XlColour.GetIndex();
            }
        }
        else
        {
            s = XlColour.GetIndex();
        }
        return s;
    }

    /// <summary>
    /// 冻结表格
    /// </summary>
    /// <param name="sheet">sheet</param>
    /// <param name="colCount">冻结的列数</param>
    /// <param name="rowCount">冻结的行数</param>
    /// <param name="startCol">右边区域可见的首列序号，从1开始计算</param>
    /// <param name="startRow">下边区域可见的首行序号，也是从1开始计算</param>
    /// <example>
    /// sheet1.CreateFreezePane(0, 1, 0, 1); 冻结首行
    /// sheet1.CreateFreezePane(1, 0, 1, 0);冻结首列
    /// </example>
    public static void FreezePane(this ISheet sheet, int colCount, int rowCount, int startCol, int startRow)
    {
        sheet.CreateFreezePane(colCount, rowCount, startCol, startRow);
    }
}