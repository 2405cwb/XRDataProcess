using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XRDataProcess
{
    /// <summary>
    /// 农村路进行图片矫正矩阵可能由于用户选择的模块错误导致超出索引
    /// </summary>
    class VillageIndexException:ApplicationException
    {
        public VillageIndexException()
        {

        }
        public VillageIndexException(string message)
            : base(message)
        {
        
        }
        public VillageIndexException(string message,Exception ex):base(message,ex)
        {

        }
    }
}
