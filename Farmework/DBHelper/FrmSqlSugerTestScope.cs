using HNDtos;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Framework.DBHelper
{
    public class FrmSqlSugerTestScope<T> where T : EntityBase, new()
    {
        public FrmSqlSugerTestScope(string connStr, SqlSugar.DbType dbType)
        {

            Db = new SqlSugarScope(new ConnectionConfig()
            {
                ConnectionString = connStr,//连接符字串
                DbType = dbType,//数据库类型
                IsAutoCloseConnection = true //不设成true要手动close


            },
           db =>
           {
               //(A)全局生效配置点，一般AOP和程序启动的配置扔这里面 ，所有上下文生效
               //调试SQL事件，可以删掉
               db.Aop.OnLogExecuting = (sql, pars) =>
               {
                   Console.WriteLine(sql);//输出sql,查看执行sql 性能无影响


                   //5.0.8.2 获取无参数化 SQL  对性能有影响，特别大的SQL参数多的，调试使用
                   //UtilMethods.GetSqlString(DbType.SqlServer,sql,pars)
               };

               //多个配置就写下面
               //db.Ado.IsDisableMasterSlaveSeparation=true;

               //注意多租户 有几个设置几个
               //db.GetConnection(i).Aop
           });
        }

        public SqlSugarScope Db = null;


        public List<T> LoadSysAdmin()
        {
            List<T> ListAdmin = Db.Queryable<T>().ToList();
            return ListAdmin;
        }
        //按要求删除某一数据
        public void DetelData(string LoginID)
        {
            Db.Deleteable<T>().Where(C => C.RoadNum == LoginID).ExecuteCommand();
        }
        //增加数据
        public void AddData(T objAdmin)
        {
            Db.Insertable<T>(objAdmin).ExecuteCommand();
        }
        //改变某一数据内容
        public void Update(T objAdmin)
        {
            Db.Updateable<T>(objAdmin).WhereColumns(c => c.RoadNum).ExecuteCommand();
        }

        public List<T> Queryable(T objAdmin)
        {
            List<T> List = Db.Queryable<T>().Where(c => c.RoadNum == objAdmin.RoadNum).ToList();
            return List;
        }

        public List<T> QueryableHeFei2(T objAdmin)
        {
            List<T> List = Db.Queryable<T>().Where(c => c.RoadNum.Contains( objAdmin.RoadNum)).ToList();
            return List;
        }
        public List<T> QueryableHeFei2(string roadNum)
        {
            List<T> List = Db.Queryable<T>().Where(c => c.RoadNum.Equals(roadNum)).ToList();
            return List;
        }

        public void InsertAll(List<T> lstData)
        {
            //WinFrom中 不能直接用，会出现卡住不动，因为WinFrom不支持直接调用异步 ，需要加委托 
            //bulkCopy 和 bulkCopyAsync 底层都是异步实现
            //Func<int> func = () => Db.Fastest<T>().BulkCopy(lstData);
            //func.BeginInvoke(x =>
            //{
            //    var result = func.EndInvoke(x);//获取返回值
            //    MessageBox.Show($"成功添加数据：{result}条");
            //}, null);

            Db.Fastest<T>().BulkCopy(lstData);

        }

    }
}
