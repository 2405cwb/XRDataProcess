/*-----------------------------------------------------------------
//CopyRight (C) 2012 武汉汉宁轨道交通技术有限公司
//版权所有。
//MyWordSzechwanDQ
//合肥要求出表适配资产表
//这里提供一个供下端  路段编号及里程区间 获得道路信息
//

 //------------------------------------------------------------------*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSExcel = Microsoft.Office.Interop.Excel;
using Framework.Office.Excel;
using Framework.Log;
using Framework.DBHelper;
using HNDtos;
using System.Windows.Forms;

namespace XRDataProcess
{
 
    public class HefeiPropertyDataMange
    {
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();
        private static MyLogger log = new MyLogger(typeof(HefeiPropertyDataMange));
        private static List<RowRoad> SelectRoads { get; set; }
        /// <summary>
        /// 获取到的数据 没有经过桩号区间过滤
        /// </summary>
        public static List<RowRoad> AllGetRoads { get; set; }
        static XRSetting _Setting = XRSetting.GetInstance();
        public static List<HeFeiEntity> AllDatas { get; set; }


        private HefeiPropertyDataMange() { }

        private static HefeiPropertyDataMange singleInstance = null;
        public static bool Direction_SX { get; set; }

        public static ProjectInfo cPorject = null;
        public static HefeiPropertyDataMange GetInstance(ProjectInfo project)
        {
            Direction_SX = project._Direction == 1 ? true : false;
            cPorject = project;
            if (AllDatas == null)
            {
                string dbPath = string.Format(@"{0}\propertyInfo.db",
                System.Windows.Forms.Application.StartupPath);
                string connectStr = $" Data Source={dbPath};";
                var sqlScope = new FrmSqlSugerTestScope<HeFeiEntity>(connectStr, SqlSugar.DbType.Sqlite);
                AllDatas = sqlScope.LoadSysAdmin();
            }
            List<HeFeiEntity> datas = AllDatas.Where(
                t => project._RoadCode.Contains(t.RoadNum)||project._RoadCode.Contains(t.RoadNum)
                ).ToList();

            if (datas.Count < 1)
            {
                MessageBox.Show(project._RoadName + "在资产表中未找到匹配的道路项");
            }
            datas = datas.OrderBy(a => a.StartMile).ToList(); //排序
            SelectRoads = new List<RowRoad>();
            //检查一下有没有空的
            string grad = "四级";
            string width = "-1";
            foreach (var item in datas)
            {
                if (!string.IsNullOrWhiteSpace(item.Width))
                {
                    grad = item.Grad;
                    width = item.Width;
                }
            }
            foreach (var item in datas)
            {
                if (string.IsNullOrWhiteSpace(item.Width))
                {
                    item.Grad = grad;
                    item.Width = width;
                }
            }
            foreach (var data in datas)
            {
                RowRoad rowRoad = new RowRoad();
                if (data.Grad.Contains("一级"))
                {

                    rowRoad.Grad = 1;
                }
                else if (data.Grad.Contains("二级"))
                {
                    rowRoad.Grad = 2;
                }

                else if (data.Grad.Contains("三级"))
                {
                    rowRoad.Grad = 3;
                }
                else if (data.Grad.Contains("四级"))
                {
                    rowRoad.Grad = 4;
                }
                else if (data.Grad.Contains("高速"))
                {
                    rowRoad.Grad = 0;
                }
                rowRoad.RoadNum = data.RoadNum;
                // rowRoad.RoadWid = double.Parse( data.Width);
                double temp = data.StartMile * 1000;
                rowRoad.StartMile = int.Parse(temp.ToString());
                temp = data.EndMile * 1000;
                rowRoad.EndMile = int.Parse(temp.ToString());
                rowRoad.IsPub = data.IsPub.Contains("是") ? true : false;
                rowRoad.RoadType = project._RoadType == 0 ? "沥青" : project._RoadType == 1 ? "水泥" : "砂石";
                if (string.IsNullOrWhiteSpace(data.Width))
                {
                    rowRoad.RoadWid = -1;
                }
                else
                {
                    rowRoad.RoadWid = double.Parse(data.Width);
                }
                rowRoad.Unit = data.Unit;
                SelectRoads.Add(rowRoad);
            }
            AllGetRoads = new List<RowRoad>(SelectRoads);

            return singleInstance;
        }
        public static List<HeFeiEntity> SelectDatas { get; set; }

        public static void updateInfo(ref List<MilePart> roadpart, Dictionary<string, int> _RoadGradeDict)
        {
            List<RowRoad> tempMark = new List<RowRoad>(AllGetRoads);
            for  ( int i = tempMark.Count-1; i>0; --i)
            {
                RowRoad preMark = tempMark[i - 1];
                RowRoad nowMark = tempMark[i];

                if (preMark.IsPub == nowMark.IsPub 
                    &&
                    preMark.RoadWid == nowMark.RoadWid
                    &&
                    preMark.Unit == nowMark.Unit)
                {
                    AllGetRoads[i - 1].EndMile = nowMark.EndMile;
                    AllGetRoads.RemoveAt(i);
                }

            }

            int sAllRoadMile = cPorject._StartMile;
            int eAllRoadMile = cPorject._EndMile;
            if (cPorject._Direction != 1)
            {
                //下行 反转 
                roadpart.Reverse();
                sAllRoadMile = cPorject._EndMile;
                eAllRoadMile = cPorject._StartMile;
            }
            //实时记录已经存在的节点  每次判断 防止重复插入
            List<int> exitsMile = new List<int>();
            foreach (var item in roadpart)
            {
                exitsMile.Add(item.mile);
            }
            //处理 
            int sMile = roadpart.First().mile;
            int eMile = roadpart.Last().mile;
            List<MilePart> temp = new List<MilePart>(roadpart);
            #region   先只进行插入

            for (int j = 0; j < AllGetRoads.Count; ++j)
            {
                for (int i = 0; i < roadpart.Count - 1; i++)
                {
                    MilePart preMile = roadpart[i];
                    MilePart cMile = roadpart[i + 1];

                    RowRoad cMarkMile = AllGetRoads[j];
                    if (cMarkMile.StartMile >= preMile.mile)
                    {
                        if (cMarkMile.StartMile < sAllRoadMile || cMarkMile.StartMile > eAllRoadMile)
                        {
                            continue;
                        }
                        else if (exitsMile.Contains(cMarkMile.StartMile))
                        {
                            continue;
                        }
                        else
                        {
                            MilePart mile = new MilePart();
                            mile.mile = cMarkMile.StartMile;
                            mile.unit = cMarkMile.Unit;
                            mile.isPub = cMarkMile.IsPub;
                            mile.width = cMarkMile.RoadWid;
                            mile.dmi = cPorject.Mile2Dmi(mile.mile);
                            mile.roadtype = preMile.roadtype;
                            mile.roaddegree = cMarkMile.Grad;
                            //标记  道路等级字符需要处理
                            foreach (var tempDict in _RoadGradeDict)
                            {
                                if (tempDict.Value == roadpart[i].roaddegree)
                                {
                                    mile.degreestr = tempDict.Key;
                                }
                            }
                            mile.roadcross = preMile.roadcross;
                            mile.roadtypelist = preMile.roadtypelist;
                            mile.isZC = true;
                            exitsMile.Add(mile.mile);
                            temp.Insert(i + 1, mile);
                        }
                    }
                    if (cMarkMile.EndMile <= cMile.mile)
                    {
                        if (cMarkMile.EndMile < sAllRoadMile || cMarkMile.EndMile > eAllRoadMile)
                        {
                            continue;
                        }
                        else if (exitsMile.Contains(cMarkMile.EndMile))
                        {
                            continue;
                        }
                        else
                        {
                            if (i - 1 == -1)
                            {
                                continue;
                            }
                            MilePart mile = new MilePart();
                            mile.mile = cMarkMile.EndMile;
                            mile.unit = cMarkMile.Unit;
                            mile.isPub = cMarkMile.IsPub;
                            mile.roaddegree = cMarkMile.Grad;
                            mile.width = cMarkMile.RoadWid;
                            //标记  道路等级字符需要处理
                            foreach (var tempDict in _RoadGradeDict)
                            {
                                if (tempDict.Value == roadpart[i].roaddegree)
                                {
                                    mile.degreestr = tempDict.Key;
                                }
                            }
                            mile.dmi = cPorject.Mile2Dmi(mile.mile);
                            mile.roadtype = preMile.roadtype;
                            mile.roadcross = preMile.roadcross;
                            mile.roadtypelist = preMile.roadtypelist;
                            mile.isZC = true;
                            exitsMile.Add(mile.mile);
                            temp.Insert(i + 1, mile);
                        }
                    }
                }
            }
            #endregion
            roadpart = temp.OrderBy(t => t.mile).ToList();
            #region  对数组中所有元素进行赋值
            if (cPorject._Direction == 1)
            {
                for (int i = 0; i < roadpart.Count - 1; ++i)
                {
                    MilePart cMile = roadpart[i];
                    MilePart endMile = roadpart[i + 1];
                    for (int j = 0; j < AllGetRoads.Count; ++j)
                    {
                        RowRoad cMarkMile = AllGetRoads[j];
                        if (cMarkMile.StartMile <= cMile.mile && cMarkMile.EndMile >= cMile.mile)
                        {
                            if (cMarkMile.StartMile <= endMile.mile && cMarkMile.EndMile >= endMile.mile)
                            {
                                roadpart[i].isPub = cMarkMile.IsPub;
                                roadpart[i].unit = cMarkMile.Unit;
                                roadpart[i].width = cMarkMile.RoadWid;
                                roadpart[i].roaddegree = cMarkMile.Grad;
                                //标记  道路等级字符需要处理
                                foreach (var tempDict in _RoadGradeDict)
                                {
                                    if (tempDict.Value == roadpart[i].roaddegree)
                                    {
                                        roadpart[i].degreestr = tempDict.Key;
                                    }
                                }
                            }

                        }
                    }
                }
            }
            else
            {
                    for (int i = roadpart.Count - 1; i >0; --i)
                    {
                        MilePart cMile = roadpart[i];
                        MilePart endMile = roadpart[i - 1];
                        for (int j = 0; j < AllGetRoads.Count; ++j)
                        {
                            RowRoad cMarkMile = AllGetRoads[j];
                            if (cMarkMile.StartMile <= endMile.mile && cMarkMile.EndMile >= endMile.mile)
                            {
                                if (cMarkMile.StartMile <=cMile.mile && cMarkMile.EndMile >= cMile.mile)
                                {
                                    roadpart[i].isPub = cMarkMile.IsPub;
                                    roadpart[i].width = cMarkMile.RoadWid; 
                                roadpart[i].unit = cMarkMile.Unit;
                                roadpart[i].roaddegree = cMarkMile.Grad;
                                    //标记  道路等级字符需要处理
                                    foreach (var tempDict in _RoadGradeDict)
                                    {
                                        if (tempDict.Value == roadpart[i].roaddegree)
                                        {
                                            roadpart[i].degreestr = tempDict.Key;
                                        }
                                    }
                                }

                            }
                        }
                    }
                   
            }

            #endregion

            #region 对由于材质打标和资产表冲突导致的小区间进行过滤

            //if (roadpart.Count >= 3)
            //{
            //    int value = _Setting.hefei2MinSplit;

            //    if (value != -1)
            //    {
            //        for (int i = roadpart.Count - 1; i > 0; i--)
            //        {
            //            var prePart = roadpart[i - 1];
            //            var nowPart = roadpart[i];

            //            if (Math.Abs(nowPart.mile - prePart.mile) < value)
            //            {
            //                if (nowPart.roaddegree != prePart.roaddegree || nowPart.unit != prePart.unit
            //                    || nowPart.roadtype != prePart.roadtype|| nowPart.width != prePart.width)
            //                {
            //                    roadpart[i - 1] = roadpart[i];
            //                    roadpart.RemoveAt(i);

            //                }
            //            }

            //        }
            //    }


            //    }

                #endregion
                //处理后  反转回来  
                if (cPorject._Direction != 1)
            {
                roadpart.Reverse();
            }
        }

        private static void ModifyValue(int endMile1, int endMile2)
        {
            endMile1 = endMile2;
        }

        static HefeiPropertyDataMange()
        {

            singleInstance = new HefeiPropertyDataMange();

        }

        public static void Clear()
        {
            if (AllGetRoads != null)
            {
                AllGetRoads.Clear();
                AllGetRoads = null;
            }
            if (SelectRoads != null)
            {
                SelectRoads.Clear();
                SelectRoads = null;
            }
        }
        private static void readExcel(ProjectInfo project)
        {
        }


    }
}
