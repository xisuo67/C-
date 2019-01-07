using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace FileExportHelper
{
    /// <summary>
    /// Pdf生成帮助类
    /// </summary>
    public class PdfHelper
    {
        private static BaseFont BF_Light = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        /// <summary>
        /// 生成PDF文件，通过读取DataTable和列头映射信息
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="excelInfo">Excel导出信息</param>
        /// <returns>文件流</returns>
        public static MemoryStream Export(DataTable dt, ExcelInfo excelInfo, string guid = null)
        {
            if (dt == null || excelInfo == null || excelInfo.ColumnInfoList == null)
            {
                throw new ArgumentNullException();
            }
            List<ColumnInfo> ColumnInfoList = excelInfo.ColumnInfoList;
            int minRowHeight = 25;
            int dataTotal = dt.Rows.Count;

            //寻找列头和DataTable之间映射关系
            foreach (DataColumn col in dt.Columns)
            {
                ColumnInfo info = ColumnInfoList.FirstOrDefault(e => e.Field.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase));
                if (info != null)
                {
                    info.IsMapDT = true;
                }
            }
            MemoryStream ms = new MemoryStream();
            using (Document document = new Document())
            {
                PdfWriter.GetInstance(document, ms);
                document.Open();

                PdfPTable table = new PdfPTable(ColumnInfoList.Count);
                table.WidthPercentage = 100f;
                //Font font = new Font(Font.NORMAL, 13, Font.NORMAL, BaseColor.BLACK);
                for (int i = 0, length = ColumnInfoList.Count; i < length; i++)
                {
                    Paragraph p = new Paragraph(ColumnInfoList[i].Header, new Font(BF_Light, 13));
                    PdfPCell cell = new PdfPCell(p);
                    cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    cell.MinimumHeight = minRowHeight;
                    table.AddCell(cell);
                }
                //开始循环所有行
                object objCellValue = null;
                string cellValue = string.Empty;

                int completeCount = 0;
                int curprogress = 30;//查询数据已经花了30%

                for (int rowIndex = 0; rowIndex < dataTotal; rowIndex++)
                {
                    int i = 0;
                    foreach (var item in ColumnInfoList)
                    {
                        if (item.IsMapDT)
                        {
                            objCellValue = dt.Rows[rowIndex][item.Field];
                            cellValue = objCellValue != DBNull.Value ? objCellValue.ToString() : string.Empty;
                        }
                        else
                        {
                            cellValue = string.Empty;
                        }
                        Paragraph p = new Paragraph(cellValue, new Font(BF_Light, 13));
                        PdfPCell cell = new PdfPCell(p);
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell.MinimumHeight = minRowHeight;

                        switch (item.Align.ToLower())
                        {
                            case "left":
                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                break;
                            case "center":
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                break;
                            case "right":
                                cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                break;
                        }
                        table.AddCell(cell);
                    }
                    i++;
                    //记录进度
                    completeCount++;
                    int temp = 30 + (int)(completeCount * 65 / dataTotal);
                    if (temp > curprogress)
                    {
                        //当temp >curprogress 才写Session,避免无用的写Session
                        curprogress = temp;
                        if (!string.IsNullOrEmpty(guid))
                        {
                            HttpContext.Current.Session[guid] = curprogress;
                        }
                    }
                }
                document.Add(table);
            }

            //记录进度
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 99;
            }
            return ms;
        }
    }
}
