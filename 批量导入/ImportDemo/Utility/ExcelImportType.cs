using System.Collections.Concurrent;
using Tools;

namespace Utility
{
    public enum ExcelImportType
    {
        /// <summary>
        /// 学生信息
        /// </summary>
        StudentsInfo = 0,
    }
    public class ExcelImporMapper
    {
        /// <summary>
        /// 业务类型模板文件路径缓存
        /// </summary>
        private static ConcurrentDictionary<ExcelImportType, string> _fileMappingDict = null;

        /// <summary>
        /// 根据业务类型获取模版文件路径
        /// </summary>
        /// <param name="type">业务类型</param>
        /// <returns>模版文件路径</returns>
        public static string GetTemplatePath(ExcelImportType type)
        {
            InitMapping();
            return _fileMappingDict[type];
        }
        /// <summary>
        /// 初始化模版文件路径缓存
        /// </summary>
        private static void InitMapping()
        {
            if (_fileMappingDict == null)
            {
                _fileMappingDict = new ConcurrentDictionary<ExcelImportType, string>();
                _fileMappingDict.TryAdd(ExcelImportType.StudentsInfo, FileHelper.GetAbsolutePath("/Template/Excel/学生信息批量导入.xls"));

            }
        }
    }
}
