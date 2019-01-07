using System.Data;
using System.IO;
using System.Web;

namespace FileExportHelper
{
    /// <summary>
    /// Execl操作帮助类
    /// </summary>
    public static class ExcelUtil
    {
        /// <summary>
        /// 拓展方法,生成EXECL
        /// </summary>
        /// <param name="info">EXECL相关信息</param>
        /// <param name="guid">导出文件时，前端传入的标识,标记一次导出任务,为此为key,从Session中更新或读取该次导出任务的进度</param>
        /// <returns>Execl路径</returns>
        public static MemoryStream ExportExeclStream(this ExcelInfo info, string guid = null)
        {
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 0;
            }
            //1.获取列表对应数据
            DataTable dt = info.Data;
            if (!string.IsNullOrEmpty(guid))
            {
                HttpContext.Current.Session[guid] = 30;//查询出结果,直接设置为30%
            }

            //2.创建文档
            MemoryStream ms = null;
            switch (info.FileFormat)
            {
                case ExportFileFormat.Excel:
                    //导出Excel
                    ms = NPOIHelper.Export(dt, info, guid);
                    break;
                case ExportFileFormat.Word:
                    //导出Word
                    ms = WordHelper.Export(dt, info, guid);
                    break;
                case ExportFileFormat.Pdf:
                    //导出Pdf
                    ms = PdfHelper.Export(dt, info, guid);
                    break;
            }
            return ms;
        }
    }
}
