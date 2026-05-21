using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using System.Reflection;
using System.IO;
using System.Collections;
using System.Data;
using DevExpress.XtraPrinting;

namespace XRDataProcess
{
    /// <summary>
    /// 道路信息
    /// </summary>
    [Serializable]
    public class RoadInfoClass
    {
        /// <summary>
        /// 道路信息ID
        /// </summary>
        public string m_id = null;

        /// <summary>
        /// 省（直辖市）
        /// </summary>
        public string m_province = null;

        /// <summary>
        /// 市
        /// </summary>
        public string m_city = null;

        /// <summary>
        /// (区)县
        /// </summary>
        public string m_district = null;

        /// <summary>
        /// 乡（镇）
        /// </summary>
        public string m_town = null;

        /// <summary>
        /// 村（街道）
        /// </summary>
        public string m_village = null;

        /// <summary>
        /// 道路名称
        /// </summary>
        public string m_name = null;

        /// <summary>
        /// 道路编号
        /// </summary>
        public string m_code = null;

        /// <summary>
        /// 道路属性
        /// </summary>
        public string m_properity = null;

        /// <summary>
        /// 道路等级
        /// </summary>
        public string m_grade = null;

        /// <summary>
        /// 道路长度（m）
        /// </summary>
        public string m_length = null;

        /// <summary>
        /// 道路宽度（m）
        /// </summary>
        public string m_width = null;

        /// <summary>
        /// 道路面积（㎡）
        /// </summary>
        public string m_roadway_area = null;

        /// <summary>
        /// 人行道面积（㎡）
        /// </summary>
        public string m_sidewalk_area = null;

        /// <summary>
        /// 路面类型
        /// </summary>
        public string m_roadtype = null;

        /// <summary>
        /// 道路起点
        /// </summary>
        public string m_roadstartlocation = null;

        /// <summary>
        /// 道路终点
        /// </summary>
        public string m_roadendlocation = null;

        /// <summary>
        /// 道路起点桩号
        /// </summary>
        public string m_roadsartmile = null;

        /// <summary>
        /// 道路终点桩号
        /// </summary>
        public string m_roadendmile = null;

        /// <summary>
        /// 建设年代
        /// </summary>
        public string m_buildyear = null;

        /// <summary>
        /// 建设单位
        /// </summary>
        public string m_buildunit = null;

        /// <summary>
        /// 设计单位
        /// </summary>
        public string m_designunit = null;

        /// <summary>
        /// 施工单位
        /// </summary>
        public string m_constructionunit = null;

        /// <summary>
        /// 监理单位
        /// </summary>
        public string m_controlunit = null;

        /// <summary>
        /// 管理单位（省）
        /// </summary>
        public string m_managementunit_province = null;

        /// <summary>
        /// 管理单位（市）
        /// </summary>
        public string m_managementunit_city = null;

        /// <summary>
        /// 管理单位（县）
        /// </summary>
        public string m_managementunit_district = null;

        /// <summary>
        /// 管理单位（署）
        /// </summary>
        public string m_managementunit_department = null;

        /// <summary>
        /// 管理单位（所）
        /// </summary>
        public string m_maintenance_center = null;

        /// <summary>
        /// 养护标段
        /// </summary>
        public string m_maintenance_section = null;

        /// <summary>
        /// 养护单位
        /// </summary>
        public string m_maintenance_unit = null;

        /// <summary>
        /// 养护项目部
        /// </summary>
        public string m_project_department = null;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    


    /// <summary>
    /// 路段信息
    /// </summary>
    [Serializable]
    public class RoadPartInfoClass
    {
        /// <summary>
        /// 道路信息
        /// </summary>
        public RoadInfoClass m_roadinfo = new RoadInfoClass();

        /// <summary>
        /// 路段信息ID
        /// </summary>
        public string m_id = null;

        /// <summary>
        /// 路段起点
        /// </summary>
        public string m_startlocation = null;

        /// <summary>
        /// 路段终点
        /// </summary>
        public string m_endlocation = null;

        /// <summary>
        /// 路段起点桩号（km）
        /// </summary>
        public string m_startmile = null;

        /// <summary>
        /// 路段终点桩号（km）
        /// </summary>
        public string m_endmile = null;

        /// <summary>
        /// 路段_道路等级
        /// </summary>
        public string m_part_grade = null;

        /// <summary>
        /// 路段长度（m）
        /// </summary>
        public string m_length = null;

        /// <summary>
        /// 路段宽度（m）
        /// </summary>
        public string m_width = null;

        /// <summary>
        /// 路段面积（m2）
        /// </summary>
        public string m_area = null;

        /// <summary>
        /// 路面类型
        /// </summary>
        public string m_type = null;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 检测参数信息
    /// </summary>
    [Serializable]
    public class IndexInfoClass
    {
        /// <summary>
        /// 检测参数ID
        /// </summary>
        public string m_id;

        /// <summary>
        /// 项目ID
        /// </summary>
        public string m_projectid;

        /// <summary>
        /// 标准参数ID
        /// </summary>
        public string m_standardid;

        /// <summary>
        /// 参数名称
        /// </summary>
        public string m_name;

        /// <summary>
        /// 指标
        /// </summary>
        public string m_index;

        /// <summary>
        /// 路面类型
        /// </summary>
        public string m_pavementtype;

        /// <summary>
        /// 是否检测
        /// </summary>
        public string m_tesing;

        

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }
    
    /// <summary>
    /// 项目信息
    /// </summary>
    [Serializable]
    public class ProjectInfoClass
    {
        /// <summary>
        /// 项目ID
        /// </summary>
        public string m_id;

        /// <summary>
        /// 项目名称
        /// </summary>
        public string m_project_name;

        /// <summary>
        /// 委托单位
        /// </summary>
        public string m_entrust_client;

        /// <summary>
        /// 委托编号
        /// </summary>
        public string m_entrust_serial;

        /// <summary>
        /// 合同编号
        /// </summary>
        public string m_contract_num;

        /// <summary>
        /// 委托日期
        /// </summary>
        public string m_entrust_date;

        /// <summary>
        /// 检测单位
        /// </summary>
        public string m_testing_unit;

        /// <summary>
        /// 项目负责人
        /// </summary>
        public string m_project_dutyperson;

        /// <summary>
        /// 检测起始日期
        /// </summary>
        public string m_testing_start_date;

        /// <summary>
        /// 检测终止日期
        /// </summary>
        public string m_testing_end_date;

        /// <summary>
        /// 检测标准ID
        /// </summary>
        public TestingStardardClass m_testing_standard = new TestingStardardClass();

        /// <summary>
        /// 检测参数
        /// </summary>
        public List<IndexInfoClass> m_indexlist = new List<IndexInfoClass>();

        /// <summary>
        /// 项目报告日期
        /// </summary>
        public string m_date;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 报告信息
    /// </summary>
    [Serializable]
    public class ReportInfoClass
    {
        /// <summary>
        /// 项目总体信息
        /// </summary>
        public ProjectInfoClass m_projectinfo = new ProjectInfoClass();

        public List<RoadPartInfoClass> m_roadpartlist = new List<RoadPartInfoClass>();

        /// <summary>
        /// 检测报告ID
        /// </summary>
        public string m_id;

        /// <summary>
        /// 报告编号
        /// </summary>
        public string m_report_num;

        /// <summary>
        /// 报告名称
        /// </summary>
        public string m_report_name;

        /// <summary>
        /// 报告说明
        /// </summary>
        public string m_project_name;

        /// <summary>
        /// 报告起始日期
        /// </summary>
        public string m_report_start_date;
        
        /// <summary>
        /// 报告终止日期
        /// </summary>
        public string m_report_end_date;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 车道信息
    /// </summary>
    [Serializable]
    public class LaneInfoClass
    {
        /// <summary>
        /// 路段信息
        /// </summary>
        public RoadPartInfoClass m_roadpartinfo = new RoadPartInfoClass();

        /// <summary>
        /// 车道信息ID
        /// </summary>
        public string m_id = null;

        /// <summary>
        /// 行车方向
        /// </summary>
        public string m_direction = null;

        /// <summary>
        /// 车道号
        /// </summary>
        public string m_lanenum = null;

        /// <summary>
        /// 车道起点桩号（km）
        /// </summary>
        public string m_startmile = null;

        /// <summary>
        /// 车道终点桩号（km）
        /// </summary>
        public string m_endmile = null;

        /// <summary>
        /// 车道宽度（m）
        /// </summary>
        public string m_width = null;

        /// <summary>
        /// 车道类型，机动车道、非机动车道、人行道
        /// </summary>
        public string m_roadfunctiontype = null;

        /// <summary>
        /// 机动车道类型，标准车道、路口车道
        /// </summary>
        public string m_carwaytype = null;

        public string m_wcStartTime = "";
        /// <summary>
        /// 路面类型
        /// </summary>
        public string m_pavementtype = null;
        
        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 报告检测路段
    /// </summary>
    [Serializable]
    public class ReoprtRoadPartInfoClass
    {
        /// <summary>
        /// 报告检测路段ID
        /// </summary>
        public string m_id;

        /// <summary>
        /// 报告信息
        /// </summary>
        public ReportInfoClass m_reoprt = new ReportInfoClass();

        /// <summary>
        /// 路段信息
        /// </summary>
        public RoadPartInfoClass m_roadpart = new RoadPartInfoClass();

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 弯沉数据
    /// </summary>
    [Serializable]
    public class WcDataClass
    {
      
        public string wcFilePath = null;

        /// <summary>
        /// 交通量等级
        /// </summary>
        public string traffic = null;


        public double wcLength = 0;
        /// <summary>
        /// 单元信息
        /// </summary>
        public DataTable unitDatas = new DataTable();

        /// <summary>
        /// 原始弯沉数据
        /// </summary>
        public DataTable wcDatas = new DataTable();

        /// <summary>
        /// 结果弯沉数据
        /// </summary>
        public DataTable wcResultDatas = new DataTable();
        
        /// <summary>
        /// 全车道弯沉值
        /// </summary>
        public double WcValue = 0;

        /// <summary>
        /// 全车道路弯沉值评价
        /// </summary>
        public string WcJudge ="无";

        /// <summary>
        /// 代表路基类型 
        /// 如果各个分段的路基类型不一致则为 /
        /// </summary>
        public string WcLjlx = "/";

        public string time = "";
    }

    /// <summary>
    /// 车道数据
    /// </summary>
    [Serializable]
    public class LaneProjectClass
    {
        /// <summary>
        /// 该车道的原始工程数据路径
        /// </summary>
        public List<string> m_projectdatapathlist = new List<string>();

        /// <summary>
        /// 该车道的弯沉数据
        /// </summary>
        public List<FileInfo> m_projectwcDataFilePath = new List<FileInfo>();

        /// <summary>
        /// 弯沉数据
        /// </summary>
        public List<WcDataClass> m_wcDataClasses = new List<WcDataClass>();

        /// <summary>
        /// 该车道的工程数据报表
        /// </summary>
        public string m_xlsxpath;

        /// <summary>
        /// 该车道的实际检测采集长度，单位m
        /// </summary>
        public int m_laneRealLength;

        /// <summary>
        /// 车道信息
        /// </summary>
        public LaneInfoClass m_lane = new LaneInfoClass();

        /// <summary>
        /// 报告信息
        /// </summary>
        public ReportInfoClass m_report = new ReportInfoClass();

       

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 路段数据
    /// </summary>
    [Serializable]
    public class RoadPartProjectClass
    {
        /// <summary>
        /// 路段信息
        /// </summary>
        public RoadPartInfoClass m_roadpart = new RoadPartInfoClass();

        /// <summary>
        /// 多个车道信息
        /// </summary>
        public List<LaneProjectClass> m_lanelist = new List<LaneProjectClass>();

        /// <summary>
        /// 地理位置示意图
        /// </summary>
        public string m_MapImg = null;

        /// <summary>
        /// 车道布置，比如双向2车道
        /// </summary>
        public string m_LaneLayout = null;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }
    
    /// <summary>
    /// 报告数据
    /// </summary>
    [Serializable]
    public class ReportProjectClass
    {
        /// <summary>
        /// 报告信息
        /// </summary>
        public ReportInfoClass m_report = new ReportInfoClass();

        /// <summary>
        /// 多个路段信息
        /// </summary>
        public List<RoadPartProjectClass> m_roadpartlist = new List<RoadPartProjectClass>();

        /// <summary>
        /// 检测人员
        /// </summary>
        public List<TestingPersonClass> m_personList = new List<TestingPersonClass>();

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    /// <summary>
    /// 项目数据
    /// </summary>
    [Serializable]
    public class ProjectProjectClass
    {
        /// <summary>
        /// 项目信息
        /// </summary>
        public ProjectInfoClass m_project = new ProjectInfoClass();

        /// <summary>
        /// 多个报告信息
        /// </summary>
        public List<ReportProjectClass> m_reportlist = new List<ReportProjectClass>();

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            //20220705 cwb


            //object retval;
            //using (MemoryStream ms = new MemoryStream())
            //{
            //    BinaryFormatter bf = new BinaryFormatter();
            //    bf.Serialize(ms, obj);
            //    ms.Seek(0, SeekOrigin.Begin);
            //    retval = bf.Deserialize(ms);
            //    ms.Close();
            //}
            //return (T)retval;
            return obj;
        }
    }

    [Serializable]
    public class TestingStardardClass
    {
        /// <summary>
        /// 检测标准ID
        /// </summary>
        public string m_id;

        /// <summary>
        /// 标准名称
        /// </summary>
        public string m_name;

        /// <summary>
        /// 标准编号
        /// </summary>
        public string m_code;

        /// <summary>
        /// 标准用途
        /// </summary>
        public string m_function;

        /// <summary>
        /// 标准类型
        /// </summary>
        public string m_type;
        
        /// <summary>
        /// 行业
        /// </summary>
        public string m_industry;

        /// <summary>
        /// 地方
        /// </summary>
        public string m_dependency;

        /// <summary>
        /// 备注
        /// </summary>
        public string m_remarks;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }

    [Serializable]
    public class TestingPersonClass
    {
        /// <summary>
        /// 姓名
        /// </summary>
        public string m_name;

        /// <summary>
        /// 证书编号
        /// </summary>
        public string m_CertificateNo;

        /// <summary>
        /// 职称
        /// </summary>
        public string m_title;

        /// <summary>
        /// 岗位
        /// </summary>
        public string m_post;

        /// <summary>
        /// 分工
        /// </summary>
        public string m_duty;

        /// <summary>
        /// 深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyByBinary<T>(T obj)
        {
            object retval;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, obj);
                ms.Seek(0, SeekOrigin.Begin);
                retval = bf.Deserialize(ms);
                ms.Close();
            }
            return (T)retval;
        }
    }
}
