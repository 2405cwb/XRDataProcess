//using log4net;
//using log4net.Config;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace HNRoadFormatConverter.Commons
//{
//    public class LogHelper
//    {
//        //public void Log()
//        //{
//        //    XmlConfigurator.Configure(new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CfgFiles\\log4net.cfg.xml")));
//        //    ILog log = LogManager.GetLogger(typeof(Logger));
//        //    log.Warn("");
//        //}
//        static LogHelper()
//        {
//            XmlConfigurator.Configure(new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config\\log4net.cfg.xml")));
//           // ILog Log = LogManager.GetLogger(typeof(LogHelper));
//           // Log.Info("系统初始化Logger模块");
//        }

//        private ILog loger = null;
//        public LogHelper(Type type)
//        {
//            loger = LogManager.GetLogger(type);
//        }
       
//        /// <summary>
//        /// Log4日志
//        /// </summary>
//        /// <param name="msg"></param>
//        /// <param name="ex"></param>
//        public void Error(string msg = "出现异常", Exception ex = null)
//        {
           
//            loger.Error(msg, ex);
//        }

//        /// <summary>
//        /// Log4日志
//        /// </summary>
//        /// <param name="msg"></param>
//        public void Warn(string msg)
//        {
           
//            loger.Warn(msg);
//        }

//        /// <summary>
//        /// Log4日志
//        /// </summary>
//        /// <param name="msg"></param>
//        public void Info(string msg)
//        {
         
//            loger.Info(msg);
//        }

//        /// <summary>
//        /// Log4日志
//        /// </summary>
//        /// <param name="msg"></param>
//        public void Debug(string msg)
//        {
//            Console.WriteLine(msg);
//            loger.Debug(msg);
//        }

//    }

//}
