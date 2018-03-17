using Models;
using Novacode;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI.WebControls;

namespace Tools
{
    /// <summary>
    /// word生成帮助类
    /// </summary>
    public class WordHelper
    {
        /// <summary>
        /// 生成Word文件，通过读取DataTable和列头映射信息
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="excelInfo">导出信息</param>
        /// <returns>文件流</returns>
        public static MemoryStream Export(DataTable dt, ExcelInfo excelInfo, string guid = null)
        {
            if (dt == null || excelInfo == null || excelInfo.ColumnInfoList == null)
            {
                throw new ArgumentNullException();
            }
            List<ColumnInfo> ColumnInfoList = excelInfo.ColumnInfoList;
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
            int fontSize = 9;
            int minRowHeight = 25;
            int dataTotal = dt.Rows.Count;
            using (DocX doc = DocX.Create(ms, DocumentTypes.Document))
            {
                Novacode.Table table = doc.InsertTable(dataTotal + 1, ColumnInfoList.Count);
                table.Design = TableDesign.TableGrid;
                table.Alignment = Alignment.center;
                List<Row> rows = table.Rows;
                //首行标题头
                Row rowHeader = rows[0];
                rowHeader.TableHeader = true;
                rowHeader.MinHeight = minRowHeight;
                for (int i = 0, length = ColumnInfoList.Count; i < length; i++)
                {
                    rowHeader.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                    rowHeader.Cells[i].FillColor = Color.FromArgb(192, 192, 192);
                    Paragraph p = rowHeader.Cells[i].Paragraphs[0];
                    p.Alignment = Alignment.center;
                    p.Append(ColumnInfoList[i].Header).Bold().FontSize(fontSize);
                }

                //开始循环所有行
                object cellValue = null;

                int completeCount = 0;
                int curprogress = 30;//查询数据已经花了30%

                for (int rowIndex = 0; rowIndex < dataTotal; rowIndex++)
                {
                    Row dataRow = rows[rowIndex + 1];
                    dataRow.MinHeight = minRowHeight;
                    int i = 0;
                    foreach (var item in ColumnInfoList)
                    {
                        if (item.IsMapDT)
                        {
                            cellValue = dt.Rows[rowIndex][item.Field];
                            Paragraph p = dataRow.Cells[i].Paragraphs[0];
                            dataRow.Cells[i].VerticalAlignment = VerticalAlignment.Center;
                            switch (item.Align.ToLower())
                            {
                                case "left":
                                    p.Alignment = Alignment.left;
                                    break;
                                case "center":
                                    p.Alignment = Alignment.center;
                                    break;
                                case "right":
                                    p.Alignment = Alignment.right;
                                    break;
                            }
                            p.Append(cellValue != DBNull.Value ? cellValue.ToString() : string.Empty);
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
                }
                doc.Save();
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
