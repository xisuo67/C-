using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileExportHelper
{
    /// <summary>
    /// Execl相关信息
    /// </summary>
    public class ExcelInfo
    {
        /// <summary>
        /// Execl列信息
        /// </summary>
        public List<ColumnInfo> ColumnInfoList { get; set; }

        /// <summary>
        /// Execl文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 应用程序根目录
        /// </summary>
        public string RootPath { get; set; }

        /// <summary>
        ///获取数据源的方法信息
        /// </summary>
        public string Api { get; set; }

        /// <summary>
        /// 获取数据请求类型 post get
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// 查询条件
        /// </summary>
        public QueryCondition Condition { get; set; }

        private List<Dictionary<string, object>> _dataEx { get; set; }

        /// <summary>
        /// 解决复杂列表序列化报错问题 2017-09-21
        /// </summary>
        [JsonProperty(PropertyName = "Data")]
        public List<Dictionary<string, object>> DataEx
        {
            get
            {
                return _dataEx;
            }
            set
            {
                _dataEx = value;
                Data = ConvertDataEx2Data(value);
            }
        }

        /// <summary>
        /// 需要导出的数据
        /// </summary>
        [JsonIgnore]
        public DataTable Data { get; set; }

        private bool isExportSelectData = false;
        /// <summary>
        /// 是否为导出当前选中数据
        /// 如果wei true 则不进行远程查询
        /// </summary>
        public bool IsExportSelectData
        {
            get { return isExportSelectData; }
            set { isExportSelectData = value; }
        }

        /// <summary>
        /// 记录导出进度的标识
        /// </summary>
        public string Guid { get; set; }

        /// <summary>
        /// 备注信息-不为空将放置在第一行
        /// </summary>
        public string Remark { get; set; }

        /// <summary>
        /// 文件格式
        /// </summary>
        public ExportFileFormat FileFormat { get; set; }

        /// <summary>
        /// 当前页面FunctionCode
        /// </summary>
        public string FunctionCode { get; set; }

        public string GetFileExt()
        {
            string result = ".xls";
            switch (FileFormat)
            {
                case ExportFileFormat.Excel:
                    result = ".xls";
                    break;
                case ExportFileFormat.Word:
                    result = ".docx";
                    break;
                case ExportFileFormat.Pdf:
                    result = ".pdf";
                    break;
            }
            return result;
        }

        /// <summary>
        /// 是否按照前台传的列来序列化DataTable
        /// </summary>
        public bool ColAsSerialize { get; set; }

        /// <summary>
        /// 将DataEx转换成Data
        /// </summary>
        public DataTable ConvertDataEx2Data(List<Dictionary<string, object>> data)
        {

            DataTable dt = new DataTable();
            if (data == null)
            {
                return dt;
            }
            if (ColumnInfoList == null || ColAsSerialize == false)
            {
                return JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(data));
            }

            List<string> columns = new List<string>(ColumnInfoList.Count);
            foreach (var item in ColumnInfoList)
            {
                if (!string.IsNullOrEmpty(item.Field))
                {
                    dt.Columns.Add(item.Field);
                    columns.Add(item.Field);
                }
            }
            if (data != null)
            {
                DataRow dr = null;
                object value = null;
                foreach (var item in data)
                {
                    dr = dt.NewRow();
                    foreach (string column in columns)
                    {
                        if (!item.TryGetValue(column, out value))
                        {
                            value = null;
                        }
                        dr[column] = value;
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        public System.Dynamic.ExpandoObject Tag { get; set; }

        /// <summary>
        /// 固定列数量
        /// </summary>
        public int FixColumns { get; set; }

        /// <summary>
        /// 合并表头信息
        /// </summary>
        public List<MoreHeader> GroupHeader { get; set; }
    }

    /// <summary>
    /// 多表头信息
    /// </summary>
    public class MoreHeader
    {
        /// <summary>
        ///  开始列列名
        /// </summary>
        public string StartColumnName { get; set; }

        /// <summary>
        ///  合并列数量
        /// </summary>
        public int NumberOfColumns { get; set; }

        /// <summary>
        ///  合并表头名称
        /// </summary>
        public string TitleText { get; set; }
    }

    /// <summary>
    /// Excel多标签页导出Sheet相关信息
    /// </summary>
    public class ExportSheetInfo
    {
        /// <summary>
        /// 标签页名称
        /// </summary>
        public string SheetName { get; set; }
        /// <summary>
        /// Sheet列信息
        /// </summary>
        public List<ColumnInfo> ColumnInfoList { get; set; }

        /// <summary>
        ///Sheet对应的数据
        /// </summary>
        public DataTable Data { get; set; }

        /// <summary>
        /// 备注信息-不为空将放置在第一行
        /// </summary>
        public string Remark { get; set; }
    }

    /// <summary>
    /// 多标签页导出信息
    /// </summary>
    public class ExportMultiSheet
    {
        /// <summary>
        /// 当前页面FunctionCode
        /// </summary>
        public string FunctionCode { get; set; }

        /// <summary>
        /// Execl文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 多标签页信息
        /// </summary>
        public List<ExportSheetInfo> ListSheet { get; set; }
    }

    public class ProgressResult
    {

        public string Guid { get; set; }
        public string Status { get; set; }
        public int Value { get; set; }
        public string Msg { get; set; }
    }

    /// <summary>
    /// 导出文件类型
    /// </summary>
    public enum ExportFileFormat
    {
        Excel = 0,
        Word = 1,
        Pdf = 2
    }
}
