using Models;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Data;
using System.IO;
using System.Linq;
using Tools;
using Utility;

namespace ImportDemo.Util
{
    /// <summary>
    /// 学生信息批量导入
    /// </summary>
    [Export(typeof(ExcelImport))]
    public class StudentExcelImport : ExcelImport
    {
        /// <summary>
        /// Excel字段映射
        /// </summary>
        private static Dictionary<string, ImportVerify> dictFields = new List<ImportVerify>
        {
            new ImportVerify{ ColumnName= "学生编号",FieldName="SNO",VerifyFunc =UniqueVerify },
            new ImportVerify{ ColumnName= "学生姓名",FieldName="SNAME",VerifyFunc =(e,extra)=> ExcelImportHelper.GetCellMsg(e.CellValue,e.ColName,50,true,true) },
            new ImportVerify{ ColumnName= "年龄",FieldName="AGE",VerifyFunc =VerifyNum},
            new ImportVerify{ ColumnName= "性别",FieldName="SEX",VerifyFunc =SelectVerify },
        }.ToDictionary(e => e.ColumnName, e => e);
        /// <summary>
        /// 学生编号唯一性校验
        /// </summary>
        /// <param name="e">校验参数</param>
        /// <param name="extra">错误信息</param>
        /// <returns></returns>
        private static string UniqueVerify(ImportVerifyParam e, object extra)
        {
            string result = "";
            result = ExcelImportHelper.GetCellMsg(e.CellValue, e.ColName, 50, true);
            if (string.IsNullOrEmpty(result))
            {
                var StudentsDict = extra as Dictionary<string, int>;
                //校验是否唯一
                int total = 0;
                if (StudentsDict.TryGetValue(e.CellValue.ToString(), out total))
                {
                    if (total > 1)
                    {
                        result += string.Format("{0}:“{1}”已经存在", e.ColName, e.CellValue);
                    }
                }
            }
            return result;
        }
        /// <summary>
        ///下拉选项校验
        /// </summary>
        /// <param name="e">校验参数</param>
        /// <returns>错误信息</returns>
        private static string SelectVerify(ImportVerifyParam e, object extra)
        {
            string result = "";
            result = ExcelImportHelper.GetCellMsg(e.CellValue, e.ColName, 0, true);
            if (string.IsNullOrEmpty(result))
            {
                //校验是否唯一
                var dict = extra as Dictionary<string, string>;
                if (dict != null)
                {
                    if (!dict.ContainsKey(e.CellValue.ToString()))
                    {
                        result += e.ColName + "下拉选项" + e.CellValue + "不存在";
                    }
                }
            }
            return result;
        }
        /// <summary>
        /// 年龄验证
        /// </summary>
        /// <param name="e"></param>
        /// <param name="extra"></param>
        /// <returns></returns>
        private static string VerifyNum(ImportVerifyParam e, object extra)
        {
            string result = "";
            result = ExcelImportHelper.GetCellMsg(e.CellValue, e.ColName, 100, false, false);
            if (string.IsNullOrEmpty(result))
            {
                if (!string.IsNullOrEmpty(e.CellValue.ToString().Trim()))
                {
                    if (!System.Text.RegularExpressions.Regex.IsMatch(e.CellValue.ToString(), @"^[0-9]*[1-9][0- ]*$"))
                    {
                        result += "请输入正确年龄!";
                    }
                }
            }
            return result;
        }
        /// <summary>
        /// 学生编号数量缓存
        /// </summary>
        private static Dictionary<string, int> GetStudentNoDict(DataTable dt)
        {
            //人员编号 数量缓存
            string sql = "select sno from students";
            List<string> listNo = DBHelper.GetDataTable(sql).ToList<Students>().Select(e => e.SNO).ToList() ;
            var q = (from p in dt.AsEnumerable()
                     where !string.IsNullOrEmpty(p.Field<string>("sno"))
                     select p.Field<string>("sno")
                   ).ToList();
            listNo.AddRange(q);

            return (from p in listNo
                    group p by p into g
                    select new
                    {
                        No = g.Key,
                        Total = g.Count()
                    }).ToDictionary(p => p.No, p => p.Total);
        }
        /// <summary>
        /// 性别校验
        /// </summary>
        /// <returns></returns>
        private static Dictionary<string, string> GetSexDict()
        {
            //默认值
            Dictionary<string, string> genderInfo = new Dictionary<string, string>() {
                { "男","0" },
                { "女" ,"1" }
            };
            return genderInfo;
        }
        /// <summary>
        /// 返回对应的导出模版数据
        /// </summary>
        /// <param name="FilePath">模版的路径</param>
        /// <param name="s">响应流</param>
        public override void GetExportTemplate(string FilePath, Stream s)
        {
            //写入下拉框值 人员类型 所属单位 警衔
            var sheet = NPOIHelper.GetFirstSheet(FilePath);
            int dataRowIndex = StartRowIndex + 1;
            string[] sex = GetSexDict().Keys.ToArray();
            NPOIHelper.SetHSSFValidation(sheet, sex, dataRowIndex, 3);
            sheet.Workbook.Write(s);
        }
        #region "override方法"
        /// <summary>
        /// 获取额外的校验所需信息
        /// </summary>
        /// <param name="listColumn">所有列名集合</param>
        /// <param name="dt">dt</param>
        /// <returns>额外信息</returns>
        /// <remarks>
        /// 例如导入excel中含有下拉框 导入时需要判断选项值是否还存在，可以通过该方法查询选项值
        /// </remarks>
        public override Dictionary<string, object> GetExtraInfo(List<string> listColumn, DataTable dt)
        {
            Dictionary<string, object> extraInfo = new Dictionary<string, object>();
            foreach (string name in listColumn)
            {
                switch (name)
                {
                    case "SNO":
                        extraInfo[name] = GetStudentNoDict(dt);
                        break;
                    case "SEX":
                        extraInfo[name] = GetSexDict();
                        break;
                    default:
                        break;
                }
            }
            return extraInfo;
        }
        /// <summary>
        /// 业务类型
        /// </summary>
        public override ExcelImportType Type => ExcelImportType.StudentsInfo;
        /// <summary>
        /// Excel字段映射及校验缓存
        /// </summary>
        public override Dictionary<string, ImportVerify> DictFields => dictFields;
        /// <summary>
        /// 保存导入信息
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="extraInfo"></param>
        /// <returns></returns>
        public override object SaveImportData(DataTable dt, Dictionary<string, object> extraInfo)
        {
            throw new System.NotImplementedException();
        }
    }
    #endregion
}