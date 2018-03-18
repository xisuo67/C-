using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Data;
using System.Linq;
using Utility;

namespace ImportDemo.Util
{
    /// <summary>
    /// 学生信息批量导入
    /// </summary>
    [Export(typeof(ExcelImport))]
    public class StudentExcelImport : ExcelImport
    {
        private static Dictionary<string, ImportVerify> dictFields = new List<ImportVerify>
        {

        }.ToDictionary(e => e.ColumnName, e => e);
        #region "override方法"
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