using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    /// <summary>
    /// 定义调用各业务接口返回结果的格式
    /// </summary>
    public class ExcelApiResult
    {
        /// <summary>
        /// 执行是否成功：true false
        /// </summary>
        /// <remarks></remarks>
        public bool HasError { get; set; }


        /// <summary>
        /// 执行返回消息
        /// </summary>
        /// <remarks></remarks>
        public string Message { get; set; }

        /// <summary>
        /// 异常信息
        /// </summary>
        public Exception Exception { get; set; }

        /// <summary>
        /// 返回的主要内容.
        /// </summary>
        public List<Dictionary<string, object>> Result { get; set; }
    }
}
