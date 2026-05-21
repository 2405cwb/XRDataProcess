//using Farmework.Other;
//using HNRoadFormatConverter.MyConfig;
//using HNRoadFormatConverter.MyEntitys;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Runtime.Remoting.Lifetime;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;

//namespace HNRoadFormatConverter
//{
//    public class GlobalExcel
//    {

//        static XRSetting _Setting = XRSetting.GetInstance();
//        static RoadConfig _RoadConfig = RoadConfig.GetInstance();

//        /// <summary>
//        /// 0-沥青，1-水泥，2-砂石
//        /// </summary>
//        public static string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };

//        /// <summary>
//        /// 0-高速、一级公路
//        /// 1-高速、一级公路
//        /// 2-二、三、四级公路
//        /// 3-二、三、四级公路
//        /// 4-二、三、四级公路
//        /// </summary>
//        public static string[] _RoadDegreeStr = { "高速、一级公路", "高速、一级公路", "二、三、四级公路", "二、三、四级公路", "二、三、四级公路", };
//        public static string[] _RoadTypeExcelStr = { "沥青路面", "水泥混凝土路面", "砂石路面" };


//        public static void GetAllMilePart(string projectpath, ProjectInfo prjinfo, int xlslen, int direction, string[] RoadGradeStr,
//           ref List<MilePart> roadpart, Dictionary<string, int> _RoadTypeDict, Dictionary<string, int> _RoadGradeDict)
//        {
//            //获取打标的信息
//            string filename = projectpath + "\\RoadStatuMarkInfo.txt";
//            List<MilePart> roadmark = new List<MilePart>();
//            List<MilePart> roadmarkUnit = new List<MilePart>();
//            if (File.Exists(filename) && !_Setting.banMarkSign)
//            {
//                string[] disinfo = File.ReadAllLines(filename, Encoding.UTF8);
//                foreach (string line in disinfo)
//                {
//                    MilePart markinfo = new MilePart();
//                    if (line.Contains("路面材质"))
//                    {
//                        string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                        string roadtype = s[s.Length - 1];

//                        //重庆农村路有砂石路面病害类型，其他规范没有
//                        if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadChongqing))
//                        {
//                            int dmival = Convert.ToInt32(s[2]);
//                            if (prjinfo._Is23dProject)
//                            {
//                                dmival = dmival - prjinfo._23dStartDmi;
//                            }
//                            markinfo.dmi = dmival;

//                            markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                            markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];

//                            //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                            markinfo.degreestr = prjinfo._RoadGrade;
//                            markinfo.roadcross = -1;
//                            roadmark.Add(markinfo);
//                        }
//                        //低等级农村公路
//                        else if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
//                            )
//                        {
//                            int dmival = Convert.ToInt32(s[2]);
//                            if (prjinfo._Is23dProject)
//                            {
//                                dmival = dmival - prjinfo._23dStartDmi;
//                            }
//                            markinfo.dmi = dmival;
//                            markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                            markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                            //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                            markinfo.degreestr = prjinfo._RoadGrade;
//                            markinfo.roadcross = -1;
//                            roadmark.Add(markinfo);
//                        }
//                        else if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
//                            )
//                        {
//                            int dmival = Convert.ToInt32(s[2]);
//                            if (prjinfo._Is23dProject)
//                            {
//                                dmival = dmival - prjinfo._23dStartDmi;
//                            }
//                            markinfo.dmi = dmival;
//                            markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                            markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                            //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                            markinfo.degreestr = prjinfo._RoadGrade;
//                            markinfo.roadcross = -1;
//                            roadmark.Add(markinfo);
//                        }
//                    }
//                    else if (line.Contains("公路等级"))
//                    {
//                        string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                        int dmival = Convert.ToInt32(s[2]);
//                        if (prjinfo._Is23dProject)
//                        {
//                            dmival = dmival - prjinfo._23dStartDmi;
//                        }
//                        markinfo.dmi = dmival;
//                        markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                        //markinfo.roadtype = roadmark[roadmark.Count - 1].roadtype;
//                        markinfo.roaddegree = _RoadGradeDict[s[s.Length - 1].Replace("主干路次干路", "主干路")];
//                        markinfo.degreestr = s[s.Length - 1];
//                        markinfo.roadcross = -1;
//                        roadmark.Add(markinfo);
//                    }
//                    else if (line.Contains("路面单元"))
//                    {
//                        string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                        string roadcross = s[s.Length - 1];
//                        int dmival = Convert.ToInt32(s[2]);
//                        if (prjinfo._Is23dProject)
//                        {
//                            dmival = dmival - prjinfo._23dStartDmi;
//                        }
//                        markinfo.dmi = dmival;
//                        markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                        markinfo.roadcross = 1;
//                        roadmarkUnit.Add(markinfo);
//                    }
//                    else if (line.Contains("路面情况"))
//                    {
//                        string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                        string roadCondition = s[s.Length - 1];
//                        int dmival = Convert.ToInt32(s[2]);
//                        if (prjinfo._Is23dProject)
//                        {
//                            dmival = dmival - prjinfo._23dStartDmi;
//                        }
//                        markinfo.dmi = dmival;
//                        markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                        markinfo.roadCondition = roadCondition;
//                        roadmarkUnit.Add(markinfo);
//                    }
//                }
//                //为了适配城镇道路 路口由出入口变为中心点的需求
//                //将路面单元单独进行处理
//                if (_Setting.roadCrossingShow)
//                {
//                    foreach (var item in roadmarkUnit)
//                    {
//                        roadmark.Add(item);
//                    }
//                }
//                else
//                {
//                    List<MilePart> newRoadPartUnit = new List<MilePart>();
//                    for (int i = 1; i < roadmarkUnit.Count; i += 2)
//                    {
//                        MilePart markinfo = new MilePart();
//                        var pre = roadmarkUnit[i - 1];
//                        var now = roadmarkUnit[i];
//                        var preDmi = pre.dmi;
//                        var nowDmi = now.dmi;
//                        int minVal = Math.Min(preDmi, nowDmi);
//                        int dmival = Math.Abs(preDmi - nowDmi) / 2 + minVal;

//                        markinfo.dmi = dmival;
//                        markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                        markinfo.roadcross = 1;
//                        roadmark.Add(markinfo);
//                    }
//                }
//            }
//            if (direction > 0)//升序
//            {
//                roadmark.Sort(delegate (MilePart x, MilePart y) { return x.mile.CompareTo(y.mile); });
//            }
//            else if (direction < 0)//降序
//            {
//                roadmark.Sort(delegate (MilePart x, MilePart y) { return y.mile.CompareTo(x.mile); });
//            }


//            if (_Setting.needSub)
//            {
//                for (int tt = 0; tt < roadmark.Count; ++tt)
//                {
//                    MilePart item = roadmark[tt];


//                    if (roadpart[0].dmi >= item.dmi)
//                    {
//                        if (item.roadtype != -1)
//                        {
//                            roadpart[0].roadtype = item.roadtype;

//                        }
//                    }

//                }
//                for (int ttt = roadmark.Count - 1; ttt > -1; --ttt)
//                {
//                    if (direction > 0)
//                    {
//                        if (roadmark[ttt].mile < prjinfo._StartMile || roadmark[ttt].mile > prjinfo._EndMile)
//                        {
//                            roadmark.RemoveAt(ttt);
//                        }
//                    }
//                    else
//                    {
//                        if (roadmark[ttt].mile > prjinfo._StartMile || roadmark[ttt].mile < prjinfo._EndMile)
//                        {
//                            roadmark.RemoveAt(ttt);
//                        }
//                    }


//                }
//            }
//            MilePart spart = new MilePart();
//            spart = roadpart[0];


//            if (roadmark.Count > 0 && roadmark[0].mile == spart.mile)
//            {
//                if (roadmark[0].roaddegree != -1)
//                {
//                    spart.roaddegree = roadmark[0].roaddegree;
//                    spart.degreestr = roadmark[0].degreestr;
//                }
//                if (roadmark[0].roaddegree != -1)
//                {
//                    spart.roaddegree = roadmark[0].roaddegree;
//                }
//                if (roadmark[0].roadcross != -1)
//                {
//                    spart.roadcross = roadmark[0].roadcross;
//                }

//                if (string.IsNullOrWhiteSpace(roadmark[0].roadCondition))
//                {
//                    spart.roadCondition = roadmark[0].roadCondition;
//                }
//            }
//            int curmile = spart.mile;

//            //上行时 当前桩号=上一个桩号 + 区间长度
//            //下行时 当前桩号=上一个桩号 - 区间长度

//            while (direction * (prjinfo._EndMile - curmile) > 0)
//            {
//                MilePart pmile = new MilePart();
//                if (direction > 0)
//                {
//                    curmile = (curmile / xlslen + direction) * xlslen;
//                }
//                else
//                {
//                    curmile = ((curmile + xlslen - 1) / xlslen + direction) * xlslen;
//                }
//                if (direction * (prjinfo._EndMile - curmile) < 0)
//                {
//                    curmile = prjinfo._EndMile;
//                }
//                pmile.mile = curmile;
//                pmile.roadtype = spart.roadtype;
//                pmile.roaddegree = spart.roaddegree;
//                pmile.degreestr = spart.degreestr;
//                pmile.isPub = false;
//                pmile.isZC = false;
//                // pmile.width = _RoadConfig.DetectWidth;
//                roadpart.Add(pmile);
//            }

//            if (_Setting.zcSplit || _Setting.shieldMark)
//            {
//                roadmark.Clear();
//            }

//            //由打标的信息，再将区间隔断
//            if (roadmark.Count > 0)
//            {
//                int roadtype = spart.roadtype;
//                int roaddegree = spart.roaddegree;
//                string degreestr = spart.degreestr;
//                if (roadmark[0].mile == spart.mile)
//                {
//                    if (roadmark[0].roaddegree != -1)
//                    {
//                        spart.roaddegree = roadmark[0].roaddegree;
//                        spart.degreestr = roadmark[0].degreestr;
//                    }
//                    if (roadmark[0].roadtype != -1)
//                    {
//                        spart.roadtype = roadmark[0].roadtype;
//                    }
//                    roadtype = spart.roadtype;
//                    roaddegree = spart.roaddegree;
//                    degreestr = spart.degreestr;
//                    roadpart[0] = spart;
//                }

//                for (int i = 1, j = 0; i < roadpart.Count; i++)
//                {
//                    while (j < roadmark.Count &&
//                        ((direction > 0 && roadpart[i - 1].mile <= roadmark[j].mile && roadpart[i].mile > roadmark[j].mile)
//                        || (direction < 0 && roadpart[i - 1].mile >= roadmark[j].mile && roadpart[i].mile < roadmark[j].mile)))
//                    {
//                        if (roadmark[j].roaddegree != -1)
//                        {
//                            if (roadpart[i - 1].roaddegree != roadmark[j].roaddegree)
//                            {   // 【道路等级】 和前一个不同时
//                                roaddegree = roadmark[j].roaddegree;
//                                degreestr = roadmark[j].degreestr;
//                                roadpart.Insert(i, roadmark[j]);
//                            }
//                            else
//                            {
//                                roadpart[i - 1].roaddegree = roadmark[j].roaddegree;
//                                roadpart[i - 1].degreestr = roadmark[j].degreestr;
//                            }
//                        }
//                        if (roadmark[j].roadtype != -1)
//                        {
//                            if (roadpart[i - 1].roadtype != roadmark[j].roadtype)
//                            {  // 【路面材质】 和前一个不同时
//                                roadtype = roadmark[j].roadtype;
//                                roadpart.Insert(i, roadmark[j]);
//                            }
//                            else
//                            {
//                                roadpart[i - 1].roadtype = roadmark[j].roadtype;
//                            }
//                        }
//                        if (roadmark[j].roadcross != -1) //新增 如果有路口标记 也分割
//                        {
//                            var diff = Math.Abs(roadpart[i - 1].mile - roadmark[j].mile);
//                            if (diff >= _Setting.SplitPartDistance)
//                            {
//                                roadpart.Insert(i, roadmark[j]);
//                            }
//                        }
//                        if (_Setting.userRoadCondition)
//                        {
//                            if (!string.IsNullOrEmpty(roadmark[j].roadCondition)) //新增 如果有路口标记 也分割
//                            {
//                                roadpart.Insert(i, roadmark[j]);
//                            }
//                        }


//                        j++;
//                    }
//                    roadpart[i].roaddegree = roaddegree;
//                    roadpart[i].degreestr = degreestr;
//                    roadpart[i].roadtype = roadtype;
//                    roadpart[i].isPub = false;
//                }
//            }

//            if (roadpart.Count < 2) return;  // 只有一个roadpart时直接退出？？？？？？？？？

//            // 去除 位置相邻、mile有重复 的部分
//            // 可能的重复原因： 1.同时满足多个可加入条件时，执行多次Insert
//            //                  2.Insert以后，下标未移动，
//            for (int i = 0; i < roadpart.Count - 1; ++i)
//            {
//                if (roadpart[i].mile == roadpart[i + 1].mile)
//                {
//                    roadpart.RemoveAt(i--);
//                }
//            }

//            List<string> degreeinfo = new List<string>();
//            int tdegree = roadpart[0].roaddegree;
//            string tdegreestr = RoadGradeStr[roadpart[0].roaddegree];
//            int tsmile = roadpart[0].mile;
//            int temile = 0;

//            //处理如果该检测路段有多个公路等级
//            foreach (MilePart tpart in roadpart)
//            {
//                temile = tpart.mile;
//                //取第一个路段为参照点，如果不是该路段等级，？？？
//                if (tdegree != tpart.roaddegree)
//                {
//                    degreeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, tdegreestr));
//                    tdegree = tpart.roaddegree;
//                    tdegreestr = RoadGradeStr[tpart.roaddegree];
//                    tsmile = tpart.mile;
//                }
//            }
//            degreeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, tdegreestr));
//            File.WriteAllLines(projectpath + "\\DegreeInfo.txt", degreeinfo.ToArray(), Encoding.UTF8);

//            if (roadpart.Count < 1) return;
//            List<string> typeinfo = new List<string>();
//            int ttype = roadpart[0].roadtype;
//            string ttypestr = _RoadTypeExcelStr[roadpart[0].roadtype];
//            tsmile = roadpart[0].mile;
//            temile = 0;
//            foreach (MilePart tpart in roadpart)
//            {
//                temile = tpart.mile;
//                if (ttype != tpart.roadtype)
//                {
//                    typeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, ttypestr));
//                    ttype = tpart.roadtype;
//                    ttypestr = _RoadTypeExcelStr[tpart.roadtype];
//                    tsmile = tpart.mile;
//                }
//            }
//            typeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, ttypestr));
//            File.WriteAllLines(projectpath + "\\RoadTypeInfo.txt", typeinfo.ToArray(), Encoding.UTF8);

//            for (int i = 0; i < roadpart.Count; ++i)
//            {
//                roadpart[i].dmi = prjinfo.Mile2Dmi(roadpart[i].mile);
//            }

//        }


//        //side = 0, 左右任意一侧判断，即认为有车辙损坏，距路面图像左边距离为0m
//        //side = 1, 左侧判断，认为左轮迹车辙损坏，距路面图像左边距离为0.9m
//        //side = 2, 右侧判断，认为右轮迹车辙损坏，距路面图像左边距离为2.8m
//        public static void GetRutDis(ProjectInfo prjinfo, double[] sval, int[] smile, double[] rutThresh, ref List<Disease> arrdis, int side = 0)
//        {
//            if (sval == null || sval.Length == 0)
//            {
//                return;
//            }

//            if (RoadDiseaseTypes.rutQthresh != 0)
//            {
//                rutThresh[0] = RoadDiseaseTypes.rutQthresh;
//                rutThresh[1] = RoadDiseaseTypes.rutZthresh;
//            }
//            else
//            {
//                rutThresh[0] = RoadDiseaseTypes.rutZthresh;
//            }
//            List<Disease> rutdis = new List<Disease>();
//            int len = sval.Length;
//            int[] disdegree = new int[len];
//            bool bflag = false;
//            for (int i = 0; i < len; ++i)
//            {
//                bflag = false;
//                for (int j = rutThresh.Length - 1; j >= 0; --j)
//                {
//                    if (sval[i] > rutThresh[j])//倒着比较 07和18标准 1为重度 0为轻度  城镇标准0为重度
//                    {
//                        disdegree[i] = j;
//                        bflag = true;
//                        break;
//                    }
//                }
//                if (bflag)
//                {
//                    continue;
//                }
//                disdegree[i] = -1;
//            }

//            string[] degreestr = { "轻", "重" };
//            int oldtype = disdegree[0];
//            int dislen = 0;

//            for (int i = 1; i < len; ++i)
//            {
//                if (oldtype == -1)
//                {
//                    oldtype = disdegree[i];
//                    dislen = 0;
//                    continue;
//                }

//                if (oldtype != disdegree[i])
//                {
//                    dislen++;

//                    Disease tempdis = new Disease();
//                    tempdis.m_mile = smile[i];
//                    tempdis.realwidth = _Setting.RutDisWidth;
//                    tempdis.realheight = dislen;
//                    tempdis.calcwidth = tempdis.realwidth;
//                    tempdis.calcheight = tempdis.realheight;
//                    tempdis.Area = tempdis.calcwidth * tempdis.calcheight;
//                    tempdis.RoadType = "沥青";
//                    if (side == 1)
//                    {
//                        tempdis.rect.X = (int)(0.9 / _RoadConfig.WidthScale);
//                    }
//                    else if (side == 2)
//                    {
//                        tempdis.rect.X = (int)(2.8 / _RoadConfig.WidthScale);
//                    }

//                    if (_Setting.ParmStyle == StandardParmType.CityRoad
//                        || _Setting.ParmStyle == StandardParmType.RuralRoadBeijing
//                        || _Setting.ParmStyle == StandardParmType.CityRoadShanghai
//                        || _Setting.ParmStyle == StandardParmType.RuralRoadLiaoning)
//                    {
//                        tempdis.RoadDisType = "车辙";
//                    }
//                    else
//                    {
//                        tempdis.RoadDisType = "车辙." + degreestr[oldtype];
//                    }
//                    rutdis.Add(tempdis);
//                    dislen = 0;
//                }
//                else
//                {
//                    dislen++;
//                }
//                oldtype = disdegree[i];
//            }

//            arrdis.AddRange(rutdis);
//        }
//        /// <summary>
//        /// 由于国检转换的精度需要 桩号为小数 写下该方法
//        /// </summary>
//        /// <param name="projectpath"></param>
//        /// <param name="prjinfo"></param>
//        /// <param name="xlslen"></param>
//        /// <param name="direction"></param>
//        /// <param name="RoadGradeStr"></param>
//        /// <param name="roadpart"></param>
//        /// <param name="_RoadTypeDict"></param>
//        /// <param name="_RoadGradeDict"></param>
//        public static void GetAllMilePartD(string projectpath, ProjectInfo prjinfo, double xlslen, int direction, string[] RoadGradeStr,
//    ref List<MilePartD> roadpart, Dictionary<string, int> _RoadTypeDict, Dictionary<string, int> _RoadGradeDict)
//        {
//            if (_Setting.needSub)
//            {


//                int oldStartMile = prjinfo._StartMile;
//                int oldEndMile = prjinfo._EndMile;

//                string[] subStr = _Setting.nowSubIndexStr.Split(',');


//                if (subStr.Length > 1)
//                {

//                    for (int t = subStr.Length - 1; t > -1; t -= 2)
//                    {
//                        //上行
//                        int num1 = int.Parse(subStr[t - 1]);
//                        int num2 = int.Parse(subStr[t]);
//                        if (prjinfo._DirectionInt == 1)
//                        {

//                            prjinfo._StartMile = num1 <= num2 ? num1 : num2;
//                            prjinfo._EndMile = num1 <= num2 ? num2 : num1;

//                        }
//                        else
//                        {
//                            prjinfo._StartMile = num1 <= num2 ? num2 : num1;
//                            prjinfo._EndMile = num1 <= num2 ? num1 : num2;

//                        }
//                        //将之前 根据真实起点桩号生成的路段清除掉 加入新的
//                        roadpart.Clear();

//                        MilePartD spart1 = new MilePartD() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade], degreestr = prjinfo._RoadGrade };
//                        roadpart.Add(spart1);
//                        //获取打标的信息
//                        string filename = projectpath + "\\RoadStatuMarkInfo.txt";
//                        List<MilePartD> roadmark = new List<MilePartD>();
//                        if (File.Exists(filename) && !_Setting.banMarkSign)
//                        {
//                            string[] disinfo = File.ReadAllLines(filename, Encoding.UTF8);
//                            foreach (string line in disinfo)
//                            {
//                                MilePartD markinfo = new MilePartD();
//                                if (line.Contains("路面材质"))
//                                {
//                                    string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                                    string roadtype = s[s.Length - 1];

//                                    //重庆农村路有砂石路面病害类型，其他规范没有
//                                    if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadChongqing))
//                                    {
//                                        int dmival = Convert.ToInt32(s[2]);
//                                        markinfo.dmi = dmival;
//                                        markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                        markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                                        //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                                        markinfo.degreestr = prjinfo._RoadGrade;
//                                        markinfo.roadcross = -1;
//                                        roadmark.Add(markinfo);
//                                    }
//                                    //低等级农村公路
//                                    else if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel))
//                                    {
//                                        int dmival = Convert.ToInt32(s[2]);
//                                        markinfo.dmi = dmival;
//                                        markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                        markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                                        //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                                        markinfo.degreestr = prjinfo._RoadGrade;
//                                        markinfo.roadcross = -1;
//                                        roadmark.Add(markinfo);
//                                    }
//                                }
//                                else if (line.Contains("公路等级"))
//                                {
//                                    string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                                    int dmival = Convert.ToInt32(s[2]);
//                                    markinfo.dmi = dmival;
//                                    markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                    //markinfo.roadtype = roadmark[roadmark.Count - 1].roadtype;
//                                    markinfo.roaddegree = _RoadGradeDict[s[s.Length - 1].Replace("主干路次干路", "主干路")];
//                                    markinfo.degreestr = s[s.Length - 1];
//                                    markinfo.roadcross = -1;
//                                    roadmark.Add(markinfo);
//                                }
//                                else if (line.Contains("路面单元"))
//                                {
//                                    string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                                    string roadcross = s[s.Length - 1];

//                                    int dmival = Convert.ToInt32(s[2]);
//                                    markinfo.dmi = dmival;
//                                    markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                    //markinfo.roadtype = roadmark[roadmark.Count - 1].roadtype;
//                                    //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                                    //markinfo.degreestr = roadmark[roadmark.Count - 1].degreestr;
//                                    markinfo.roadcross = 1;
//                                    roadmark.Add(markinfo);
//                                }
//                            }
//                        }

//                        if (direction > 0)//升序
//                        {
//                            roadmark.Sort(delegate (MilePartD x, MilePartD y) { return x.mile.CompareTo(y.mile); });
//                        }
//                        else if (direction < 0)//降序
//                        {
//                            roadmark.Sort(delegate (MilePartD x, MilePartD y) { return y.mile.CompareTo(x.mile); });
//                        }

//                        //分区间
//                        MilePartD spart = new MilePartD();
//                        spart = roadpart[0];
//                        if (roadmark.Count > 0 && roadmark[0].mile == spart.mile)
//                        {
//                            if (roadmark[0].roaddegree != -1)
//                            {
//                                spart.roaddegree = roadmark[0].roaddegree;
//                                spart.degreestr = roadmark[0].degreestr;
//                            }
//                            if (roadmark[0].roadtype != -1)
//                            {
//                                spart.roadtype = roadmark[0].roadtype;
//                            }
//                        }
//                        double curmile = spart.mile;

//                        //上行时 当前桩号=上一个桩号 + 区间长度
//                        //下行时 当前桩号=上一个桩号 - 区间长度

//                        while (direction * (prjinfo._EndMile - curmile) > 0)
//                        {
//                            MilePartD pmile = new MilePartD();
//                            curmile = curmile + direction * xlslen;
//                            //if (direction > 0)
//                            //{
//                            //    curmile = (curmile / xlslen + direction) * xlslen;
//                            //}
//                            //else
//                            //{
//                            //    curmile = ((curmile + xlslen - 1) / xlslen + direction) * xlslen;
//                            //}
//                            if (direction * (prjinfo._EndMile - curmile) < 0)
//                            {
//                                curmile = prjinfo._EndMile;
//                            }
//                            pmile.mile = curmile;
//                            pmile.roadtype = spart.roadtype;
//                            pmile.roaddegree = spart.roaddegree;
//                            pmile.degreestr = spart.degreestr;
//                            roadpart.Add(pmile);
//                        }


//                        //由打标的信息，再将区间隔断
//                        if (roadmark.Count > 0)
//                        {
//                            int roadtype = spart.roadtype;
//                            int roaddegree = spart.roaddegree;
//                            string degreestr = spart.degreestr;
//                            if (roadmark[0].mile == spart.mile)
//                            {
//                                if (roadmark[0].roaddegree != -1)
//                                {
//                                    spart.roaddegree = roadmark[0].roaddegree;
//                                    spart.degreestr = roadmark[0].degreestr;
//                                }
//                                if (roadmark[0].roadtype != -1)
//                                {
//                                    spart.roadtype = roadmark[0].roadtype;
//                                }
//                                roadtype = spart.roadtype;
//                                roaddegree = spart.roaddegree;
//                                degreestr = spart.degreestr;
//                                roadpart[0] = spart;
//                            }

//                            for (int i = 1, j = 0; i < roadpart.Count; i++)
//                            {
//                                while (j < roadmark.Count &&
//                                    ((direction > 0 && roadpart[i - 1].mile <= roadmark[j].mile && roadpart[i].mile > roadmark[j].mile)
//                                    || (direction < 0 && roadpart[i - 1].mile >= roadmark[j].mile && roadpart[i].mile < roadmark[j].mile)))
//                                {
//                                    if (roadmark[j].roaddegree != -1)
//                                    {
//                                        if (roadpart[i - 1].roaddegree != roadmark[j].roaddegree)
//                                        {
//                                            roaddegree = roadmark[j].roaddegree;
//                                            degreestr = roadmark[j].degreestr;
//                                            roadpart.Insert(i, roadmark[j]);
//                                        }
//                                        else
//                                        {
//                                            roadpart[i - 1].roaddegree = roadmark[j].roaddegree;
//                                            roadpart[i - 1].degreestr = roadmark[j].degreestr;
//                                        }
//                                    }
//                                    if (roadmark[j].roadtype != -1)
//                                    {
//                                        if (roadpart[i - 1].roadtype != roadmark[j].roadtype)
//                                        {
//                                            roadtype = roadmark[j].roadtype;
//                                            roadpart.Insert(i, roadmark[j]);
//                                        }
//                                        else
//                                        {
//                                            roadpart[i - 1].roadtype = roadmark[j].roadtype;
//                                        }
//                                    }
//                                    if (roadmark[j].roadcross != -1) //新增 如果有路口标记 也分割
//                                    {
//                                        roadpart.Insert(i, roadmark[j]);
//                                    }
//                                    j++;
//                                }
//                                roadpart[i].roaddegree = roaddegree;
//                                roadpart[i].degreestr = degreestr;
//                                roadpart[i].roadtype = roadtype;
//                            }
//                        }

//                        if (roadpart.Count < 2) return;
//                        for (int i = 0; i < roadpart.Count - 1; ++i)
//                        {
//                            if (roadpart[i].mile == roadpart[i + 1].mile)
//                            {
//                                roadpart.RemoveAt(i--);
//                            }
//                        }

//                        List<string> degreeinfo = new List<string>();
//                        int tdegree = roadpart[0].roaddegree;
//                        string tdegreestr = RoadGradeStr[roadpart[0].roaddegree];
//                        double tsmile = roadpart[0].mile;
//                        double temile = 0;

//                        //处理如果该检测路段有多个公路等级
//                        foreach (MilePartD tpart in roadpart)
//                        {
//                            temile = tpart.mile;
//                            //取第一个路段为参照点，如果不是该路段等级，？？？
//                            if (tdegree != tpart.roaddegree)
//                            {
//                                degreeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, tdegreestr));
//                                tdegree = tpart.roaddegree;
//                                tdegreestr = RoadGradeStr[tpart.roaddegree];
//                                tsmile = tpart.mile;
//                            }
//                        }
//                        degreeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, tdegreestr));
//                        File.WriteAllLines(projectpath + "\\DegreeInfo.txt", degreeinfo.ToArray(), Encoding.UTF8);

//                        if (roadpart.Count < 1) return;
//                        List<string> typeinfo = new List<string>();
//                        int ttype = roadpart[0].roadtype;
//                        string ttypestr = _RoadTypeExcelStr[roadpart[0].roadtype];
//                        tsmile = roadpart[0].mile;
//                        temile = 0;
//                        foreach (MilePartD tpart in roadpart)
//                        {
//                            temile = tpart.mile;
//                            if (ttype != tpart.roadtype)
//                            {
//                                typeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, ttypestr));
//                                ttype = tpart.roadtype;
//                                ttypestr = _RoadTypeExcelStr[tpart.roadtype];
//                                tsmile = tpart.mile;
//                            }
//                        }
//                        typeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, ttypestr));
//                        File.WriteAllLines(projectpath + "\\RoadTypeInfo.txt", typeinfo.ToArray(), Encoding.UTF8);

//                        for (int i = 0; i < roadpart.Count; ++i)
//                        {
//                            roadpart[i].dmi = prjinfo.Mile2Dmi(roadpart[i].mile);
//                        }


//                    }

//                }
//                prjinfo._StartMile = oldStartMile;
//                prjinfo._EndMile = oldEndMile;

//            }


//            else
//            {
//                //获取打标的信息
//                string filename = projectpath + "\\RoadStatuMarkInfo.txt";
//                List<MilePartD> roadmark = new List<MilePartD>();
//                if (File.Exists(filename) && !_Setting.banMarkSign)
//                {
//                    string[] disinfo = File.ReadAllLines(filename, Encoding.UTF8);
//                    foreach (string line in disinfo)
//                    {
//                        MilePartD markinfo = new MilePartD();
//                        if (line.Contains("路面材质"))
//                        {
//                            string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                            string roadtype = s[s.Length - 1];

//                            //重庆农村路有砂石路面病害类型，其他规范没有
//                            if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadChongqing))
//                            {
//                                int dmival = Convert.ToInt32(s[2]);
//                                markinfo.dmi = dmival;
//                                markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                                //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                                markinfo.degreestr = prjinfo._RoadGrade;
//                                markinfo.roadcross = -1;
//                                roadmark.Add(markinfo);
//                            }
//                            //低等级农村公路
//                            else if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel))
//                            {
//                                int dmival = Convert.ToInt32(s[2]);
//                                markinfo.dmi = dmival;
//                                markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                                //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                                markinfo.degreestr = prjinfo._RoadGrade;
//                                markinfo.roadcross = -1;
//                                roadmark.Add(markinfo);
//                            }
//                            else if (roadtype == "沥青" || roadtype == "水泥" || (roadtype == "砂石" && _Setting.ParmStyle == StandardParmType.RuralRoadHunan))
//                            {
//                                int dmival = Convert.ToInt32(s[2]);
//                                markinfo.dmi = dmival;
//                                markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                                markinfo.roadtype = _RoadTypeDict[s[s.Length - 1]];
//                                //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                                markinfo.degreestr = prjinfo._RoadGrade;
//                                markinfo.roadcross = -1;
//                                roadmark.Add(markinfo);
//                            }
//                        }
//                        else if (line.Contains("公路等级"))
//                        {
//                            string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                            int dmival = Convert.ToInt32(s[2]);
//                            markinfo.dmi = dmival;
//                            markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                            //markinfo.roadtype = roadmark[roadmark.Count - 1].roadtype;
//                            markinfo.roaddegree = _RoadGradeDict[s[s.Length - 1].Replace("主干路次干路", "主干路")];
//                            markinfo.degreestr = s[s.Length - 1];
//                            markinfo.roadcross = -1;
//                            roadmark.Add(markinfo);
//                        }
//                        else if (line.Contains("路面单元"))
//                        {
//                            string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                            string roadcross = s[s.Length - 1];

//                            int dmival = Convert.ToInt32(s[2]);
//                            markinfo.dmi = dmival;
//                            markinfo.mile = prjinfo.Dmi2Mile(dmival);
//                            //markinfo.roadtype = roadmark[roadmark.Count - 1].roadtype;
//                            //markinfo.roaddegree = roadmark[roadmark.Count - 1].roaddegree;
//                            //markinfo.degreestr = roadmark[roadmark.Count - 1].degreestr;
//                            markinfo.roadcross = 1;
//                            roadmark.Add(markinfo);
//                        }
//                    }
//                }

//                if (direction > 0)//升序
//                {
//                    roadmark.Sort(delegate (MilePartD x, MilePartD y) { return x.mile.CompareTo(y.mile); });
//                }
//                else if (direction < 0)//降序
//                {
//                    roadmark.Sort(delegate (MilePartD x, MilePartD y) { return y.mile.CompareTo(x.mile); });
//                }
//                MilePartD spart = new MilePartD();
//                spart = roadpart[0];
//                if (roadmark.Count > 0 && roadmark[0].mile == spart.mile)
//                {
//                    if (roadmark[0].roaddegree != -1)
//                    {
//                        spart.roaddegree = roadmark[0].roaddegree;
//                        spart.degreestr = roadmark[0].degreestr;
//                    }
//                    if (roadmark[0].roadtype != -1)
//                    {
//                        spart.roadtype = roadmark[0].roadtype;
//                    }
//                }
//                double curmile = spart.mile;

//                //上行时 当前桩号=上一个桩号 + 区间长度
//                //下行时 当前桩号=上一个桩号 - 区间长度 
//                while (direction * (prjinfo._EndMile - curmile) > 0)
//                {
//                    MilePartD pmile = new MilePartD();
//                    curmile = curmile + direction * xlslen;
//                    /*if (direction > 0)
//                      {
//                          curmile = (curmile / xlslen + direction) * xlslen;
//                      }
//                      else
//                      {
//                          curmile = ((curmile + xlslen - 1) / xlslen + direction) * xlslen;
//                      }*/
//                    if (direction * (prjinfo._EndMile - curmile) < 0)
//                    {
//                        curmile = prjinfo._EndMile;
//                    }
//                    // curmile = (float)Math.Round(curmile, 1);
//                    pmile.mile = curmile;
//                    pmile.roadtype = spart.roadtype;
//                    pmile.roaddegree = spart.roaddegree;
//                    pmile.degreestr = spart.degreestr;
//                    roadpart.Add(pmile);
//                }



//                //由打标的信息，再将区间隔断
//                if (roadmark.Count > 0)
//                {
//                    int roadtype = spart.roadtype;
//                    int roaddegree = spart.roaddegree;
//                    string degreestr = spart.degreestr;
//                    if (roadmark[0].mile == spart.mile)
//                    {
//                        if (roadmark[0].roaddegree != -1)
//                        {
//                            spart.roaddegree = roadmark[0].roaddegree;
//                            spart.degreestr = roadmark[0].degreestr;
//                        }
//                        if (roadmark[0].roadtype != -1)
//                        {
//                            spart.roadtype = roadmark[0].roadtype;
//                        }
//                        roadtype = spart.roadtype;
//                        roaddegree = spart.roaddegree;
//                        degreestr = spart.degreestr;
//                        roadpart[0] = spart;
//                    }

//                    for (int i = 1, j = 0; i < roadpart.Count; i++)
//                    {
//                        while (j < roadmark.Count &&
//                            ((direction > 0 && roadpart[i - 1].mile <= roadmark[j].mile && roadpart[i].mile > roadmark[j].mile)
//                            || (direction < 0 && roadpart[i - 1].mile >= roadmark[j].mile && roadpart[i].mile < roadmark[j].mile)))
//                        {
//                            if (roadmark[j].roaddegree != -1)
//                            {
//                                if (roadpart[i - 1].roaddegree != roadmark[j].roaddegree)
//                                {
//                                    roaddegree = roadmark[j].roaddegree;
//                                    degreestr = roadmark[j].degreestr;
//                                    roadpart.Insert(i, roadmark[j]);
//                                }
//                                else
//                                {
//                                    roadpart[i - 1].roaddegree = roadmark[j].roaddegree;
//                                    roadpart[i - 1].degreestr = roadmark[j].degreestr;
//                                }
//                            }
//                            if (roadmark[j].roadtype != -1)
//                            {
//                                if (roadpart[i - 1].roadtype != roadmark[j].roadtype)
//                                {
//                                    roadtype = roadmark[j].roadtype;
//                                    roadpart.Insert(i, roadmark[j]);
//                                }
//                                else
//                                {
//                                    roadpart[i - 1].roadtype = roadmark[j].roadtype;
//                                }
//                            }
//                            if (roadmark[j].roadcross != -1) //新增 如果有路口标记 也分割
//                            {
//                                roadpart.Insert(i, roadmark[j]);
//                            }
//                            j++;
//                        }
//                        roadpart[i].roaddegree = roaddegree;
//                        roadpart[i].degreestr = degreestr;
//                        roadpart[i].roadtype = roadtype;
//                    }
//                }

//                if (roadpart.Count < 2) return;
//                for (int i = 0; i < roadpart.Count - 1; ++i)
//                {
//                    if (roadpart[i].mile == roadpart[i + 1].mile)
//                    {
//                        roadpart.RemoveAt(i--);
//                    }
//                }

//                List<string> degreeinfo = new List<string>();
//                int tdegree = roadpart[0].roaddegree;
//                string tdegreestr = RoadGradeStr[roadpart[0].roaddegree];
//                double tsmile = roadpart[0].mile;
//                double temile = 0;

//                //处理如果该检测路段有多个公路等级
//                foreach (MilePartD tpart in roadpart)
//                {
//                    temile = tpart.mile;
//                    //取第一个路段为参照点，如果不是该路段等级，？？？
//                    if (tdegree != tpart.roaddegree)
//                    {
//                        degreeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, tdegreestr));
//                        tdegree = tpart.roaddegree;
//                        tdegreestr = RoadGradeStr[tpart.roaddegree];
//                        tsmile = tpart.mile;
//                    }
//                }
//                degreeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, tdegreestr));
//                File.WriteAllLines(projectpath + "\\DegreeInfo.txt", degreeinfo.ToArray(), Encoding.UTF8);

//                if (roadpart.Count < 1) return;
//                List<string> typeinfo = new List<string>();
//                int ttype = roadpart[0].roadtype;
//                string ttypestr = _RoadTypeExcelStr[roadpart[0].roadtype];
//                tsmile = roadpart[0].mile;
//                temile = 0;
//                foreach (MilePartD tpart in roadpart)
//                {
//                    temile = tpart.mile;
//                    if (ttype != tpart.roadtype)
//                    {
//                        typeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, ttypestr));
//                        ttype = tpart.roadtype;
//                        ttypestr = _RoadTypeExcelStr[tpart.roadtype];
//                        tsmile = tpart.mile;
//                    }
//                }
//                typeinfo.Add(string.Format("{0:0.000}-{1:0.000}km{2}", tsmile * 0.001, temile * 0.001, ttypestr));
//                File.WriteAllLines(projectpath + "\\RoadTypeInfo.txt", typeinfo.ToArray(), Encoding.UTF8);

//                for (int i = 0; i < roadpart.Count; ++i)
//                {
//                    roadpart[i].dmi = prjinfo.Mile2Dmi(roadpart[i].mile);
//                }
//            }


//        }

//        //获取工程中车速
//        public static bool GetSpeedMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] val)
//        {
//            if (val != null) val = null;
//            if (!prjinfo._IsIRIMTD) return false;

//            string[] LStrs = null;
//            string[] RStrs = null;
//            string LIRIfrname = string.Format(@"{0}\IRIMTD\DAQ0\Speed_{1}m.txt", prjdir.FullName, 10);
//            string RIRIfrname = string.Format(@"{0}\IRIMTD\DAQ1\Speed_{1}m.txt", prjdir.FullName, 10);
//            if (File.Exists(LIRIfrname))
//            {
//                LStrs = File.ReadAllLines(LIRIfrname);
//                if (LStrs.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    //MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【平整度】计算到一半退出了软件\n请【清除结果——平整度】后重新【计算IRM】!");
//                }
//            }
//            else
//            {
//                //MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度数据!\r\n请检查数据完整性，并重新计算IRM！");
//                return false;
//            }

//            if (prjinfo._IsDIRIMTD)
//            {
//                if (File.Exists(RIRIfrname))
//                {
//                    RStrs = File.ReadAllLines(RIRIfrname);
//                    if (RStrs.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        //MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【平整度】计算到一半退出了软件\n请【清除结果——平整度】后重新【计算IRM】!");
//                    }
//                }
//                else
//                {
//                    //MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度数据!\r\n请检查数据完整性，并重新计算IRM！");
//                    return false;
//                }
//            }

//            const double BaseLen = 10;
//            int len = 0;
//            if (roadpart.Count > 0)
//                len = roadpart.Count - 1;
//            double[] lval = new double[len];
//            double[] rval = new double[len];
//            val = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            double lastvalL = 0, lastvalR = 0;
//            for (int i = 0; i < len; i++)
//            {
//                double suml = 0, sumr = 0;
//                string[] tmtd;
//                int lvalnum = 0, rvalnum = 0;
//                startidx = (int)Math.Round((roadpart[i].dmi - 0.5) / BaseLen);  // dmi模baseLen的结果 超过baseLen一半时，才进1位
//                endidx = (int)Math.Round(roadpart[i + 1].dmi / BaseLen);

//                if (startidx >= endidx)
//                {
//                    if (startidx < LStrs.Length)
//                    {
//                        LStrLine = LStrs[startidx];
//                        tmtd = LStrLine.Split(' ');
//                        if (tmtd.Length < 2)
//                        {

//                            tmtd = LStrLine.Split('\t');
//                        }
//                        if (tmtd[0] != "")
//                        {
//                            lastvalL = double.Parse(tmtd[1]);
//                        }
//                        suml += lastvalL;
//                        ++lvalnum;
//                    }
//                    if (prjinfo._IsDIRIMTD)
//                    {
//                        if (startidx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[startidx];
//                            tmtd = RStrLine.Split(' ');
//                            if (tmtd.Length < 2)
//                            {

//                                tmtd = RStrLine.Split('\t');
//                            }
//                            if (tmtd[0] != "")
//                            {
//                                lastvalR = double.Parse(tmtd[1]);
//                            }
//                            sumr += lastvalR;
//                            ++rvalnum;
//                        }
//                    }
//                }
//                else
//                {
//                    for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tmtd = LStrLine.Split(' ');
//                            if (tmtd.Length < 2)
//                            {
//                                tmtd = LStrLine.Split('\t');
//                            }
//                            if (tmtd[0] != "")
//                            {
//                                lastvalL = double.Parse(tmtd[1]);
//                            }
//                            suml += lastvalL;
//                            ++lvalnum;
//                        }
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                            if (ValStridx < RStrs.Length)
//                            {
//                                RStrLine = RStrs[ValStridx];
//                                tmtd = RStrLine.Split(' ');
//                                if (tmtd.Length < 2)
//                                {
//                                    tmtd = RStrLine.Split('\t');
//                                }
//                                if (tmtd[0] != "")
//                                {
//                                    lastvalR = double.Parse(tmtd[1]);
//                                }
//                                sumr += lastvalR;
//                                ++rvalnum;
//                            }
//                        }
//                    }
//                }

//                if (lvalnum > 0) suml /= lvalnum;
//                if (rvalnum > 0) sumr /= rvalnum;

//                if (lvalnum > 0) lval[i] = suml;
//                else if (rvalnum > 0) lval[i] = sumr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (prjinfo._IsDIRIMTD)
//                {
//                    if (rvalnum > 0) rval[i] = sumr;
//                    else if (lvalnum > 0) rval[i] = suml;
//                    else rval[i] = i > 0 ? rval[i - 1] : 0;

//                    val[i] = (lval[i] + rval[i]) / 2;
//                }
//                else
//                {
//                    val[i] = lval[i];
//                }
//            }

//            if (val != null)
//            {
//                int ttlen = val.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    val[i] = Math.Round(val[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }
//        public static bool GetSpeedMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePartD> roadpart, ref double[] val)
//        {
//            if (val != null) val = null;
//            if (!prjinfo._IsIRIMTD) return false;

//            string[] LStrs = null;
//            string[] RStrs = null;
//            string LIRIfrname = string.Format(@"{0}\IRIMTD\DAQ0\Speed_{1}m.txt", prjdir.FullName, 10);
//            string RIRIfrname = string.Format(@"{0}\IRIMTD\DAQ1\Speed_{1}m.txt", prjdir.FullName, 10);
//            if (File.Exists(LIRIfrname))
//            {
//                LStrs = File.ReadAllLines(LIRIfrname);
//                if (LStrs.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    //MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【平整度】计算到一半退出了软件\n请【清除结果——平整度】后重新【计算IRM】!");
//                }
//            }
//            else
//            {
//                //MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度数据!\r\n请检查数据完整性，并重新计算IRM！");
//                return false;
//            }

//            if (prjinfo._IsDIRIMTD)
//            {
//                if (File.Exists(RIRIfrname))
//                {
//                    RStrs = File.ReadAllLines(RIRIfrname);
//                    if (RStrs.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        //MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【平整度】计算到一半退出了软件\n请【清除结果——平整度】后重新【计算IRM】!");
//                    }
//                }
//                else
//                {
//                    //MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度数据!\r\n请检查数据完整性，并重新计算IRM！");
//                    return false;
//                }
//            }

//            const double BaseLen = 10;
//            int len = 0;
//            if (roadpart.Count > 0)
//                len = roadpart.Count - 1;
//            double[] lval = new double[len];
//            double[] rval = new double[len];
//            val = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            double lastvalL = 0, lastvalR = 0;
//            for (int i = 0; i < len; i++)
//            {
//                double suml = 0, sumr = 0;
//                string[] tmtd;
//                int lvalnum = 0, rvalnum = 0;
//                startidx = (int)Math.Round((roadpart[i].dmi - 0.5) / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi / BaseLen);

//                if (startidx >= endidx)
//                {
//                    if (startidx < LStrs.Length)
//                    {
//                        LStrLine = LStrs[startidx];
//                        tmtd = LStrLine.Split(' ');
//                        if (tmtd[0] != "")
//                        {
//                            lastvalL = double.Parse(tmtd[1]);
//                        }
//                        suml += lastvalL;
//                        ++lvalnum;
//                    }
//                    if (prjinfo._IsDIRIMTD)
//                    {
//                        if (startidx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[startidx];
//                            tmtd = RStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {
//                                lastvalR = double.Parse(tmtd[1]);
//                            }
//                            sumr += lastvalR;
//                            ++rvalnum;
//                        }
//                    }
//                }
//                else
//                {
//                    for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tmtd = LStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {
//                                lastvalL = double.Parse(tmtd[1]);
//                            }
//                            suml += lastvalL;
//                            ++lvalnum;
//                        }
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                            if (ValStridx < RStrs.Length)
//                            {
//                                RStrLine = RStrs[ValStridx];
//                                tmtd = RStrLine.Split(' ');
//                                if (tmtd[0] != "")
//                                {
//                                    lastvalR = double.Parse(tmtd[1]);
//                                }
//                                sumr += lastvalR;
//                                ++rvalnum;
//                            }
//                        }
//                    }
//                }

//                if (lvalnum > 0) suml /= lvalnum;
//                if (rvalnum > 0) sumr /= rvalnum;

//                if (lvalnum > 0) lval[i] = suml;
//                else if (rvalnum > 0) lval[i] = sumr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (prjinfo._IsDIRIMTD)
//                {
//                    if (rvalnum > 0) rval[i] = sumr;
//                    else if (lvalnum > 0) rval[i] = suml;
//                    else rval[i] = i > 0 ? rval[i - 1] : 0;

//                    val[i] = (lval[i] + rval[i]) / 2;
//                }
//                else
//                {
//                    val[i] = lval[i];
//                }
//            }

//            if (val != null)
//            {
//                int ttlen = val.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    val[i] = Math.Round(val[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }

//        public static bool GetIRIHValF(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePartD> roadpart, double disval, int side, ref double[] val)
//        {
//            string fname = string.Format(@"{0}\IRIMTD\DAQ{1}\Resample.txt", prjdir.FullName, side);
//            if (!File.Exists(fname))
//            {
//                //新增简易模式
//                string simpfiedName = fname.Replace("Resample.txt", "DeltaHVal.txt");
//                /*if (GetDeltaHValFromLocal(simpfiedName, ref val))
//                {
//                    return true;
//                }
//                */
//                MessageBox.Show("缺少路面纵断面文件：" + fname);
//                return false;
//            }
//            string[] datastrs = File.ReadAllLines(fname);
//            int fileLen = datastrs.Length;
//            if (fileLen == 0)
//            {
//                return false;
//            }

//            // 步骤1: 模拟LoadData读取原始数据（oridata + 插值异常点）
//            double[] oridata = new double[fileLen];
//            double[] toridata = new double[fileLen];  // 模拟toridata (raw拷贝，用于滤波源)
//            int[] oritime = new int[fileLen];  // 未用，但模拟完整

//            for (int i = 0; i < fileLen; i++)
//            {
//                string[] s = datastrs[i].Split('\t');
//                if (s.Length <= 1)
//                {
//                    s = datastrs[i].Split(' ');  // 模拟LoadData兼容
//                }
//                if (s.Length > 3)
//                {
//                    try
//                    {
//                        oridata[i] = double.Parse(s[2]);
//                        oritime[i] = int.Parse(s[3]);
//                    }
//                    catch
//                    {
//                        if (i > 1)
//                        {
//                            // 模拟LoadData插值（线性外推，保持一致）
//                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
//                            oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];
//                        }
//                        else
//                        {
//                            oridata[i] = (i == 0) ? 0 : oridata[i - 1];  // 边界默认0或前点
//                            oritime[i] = (i == 0) ? 0 : oritime[i - 1];
//                        }
//                    }
//                }
//                else
//                {
//                    if (i > 1)
//                    {
//                        oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
//                        oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];
//                    }
//                    else
//                    {
//                        oridata[i] = 0;
//                        oritime[i] = 0;
//                    }
//                }
//                toridata[i] = oridata[i];  // 拷贝raw到toridata（模拟LoadData toridata = oridata）
//            }

//            // 步骤2: 均值滤波（模拟GenerateIRI_NEW步骤3，5点窗口，边界保持原始）
//            for (int i = 2; i < fileLen - 2; ++i)
//            {
//                oridata[i] = (toridata[i - 2] + toridata[i - 1] + toridata[i] + toridata[i + 1] + toridata[i + 2]) / 5;
//            }
//            // 边界保持原始（oridata[0]/[1]和[fileLen-1]/[fileLen-2]不变，与GenerateIRI_NEW一致）


//            // 步骤3: 采样到指定间隔（disval e.g. 0.1m，与GenerateIRI_NEW ResampleData类似）
//            int len = roadpart.Count - 1;
//            val = new double[len];
//            const double pluselen = 0.05;  // 原始间隔0.05m
//            double baselen = disval;
//            int skipnum = (int)(baselen / pluselen);  // 每disval采样skipnum原始点（模拟qplusenum）

//            for (int i = 0; i < len; ++i)
//            {
//                int index = i * skipnum;
//                if (index < oridata.Length)
//                {
//                    val[i] = oridata[index];
//                }
//                else
//                {
//                    val[i] = oridata[oridata.Length - 1];  // 超出用最后一个滤波点（模拟Resample外推）
//                }
//            }

//            string speedSavefname = string.Format(@"{0}\IRIMTD\DAQ{1}\Speed_10m.txt", prjdir.FullName, side);
//            double[] kparms = null, speedparms = null;
//            double[] bparms = null;
//            string fparmpath = fname.Replace("Resample.txt", "Coeff.dat");
//            int parmnum = 0;
//            if (File.Exists(fparmpath))
//            {
//                string[] parms = File.ReadAllLines(fparmpath);
//                try
//                {
//                    parmnum = int.Parse(parms[0]);
//                    speedparms = new double[parmnum];
//                    kparms = new double[parmnum];
//                    bparms = new double[parmnum];
//                    int idx = 1;
//                    for (int i = 0; i < parmnum; ++i) speedparms[i] = double.Parse(parms[idx++]);
//                    for (int i = 0; i < parmnum; ++i) kparms[i] = double.Parse(parms[idx++]);
//                    for (int i = 0; i < parmnum; ++i) bparms[i] = double.Parse(parms[idx++]);
//                }
//                catch
//                {
//                    MessageBox.Show("读取文件出错，请检查！\r\n" + fparmpath);
//                }
//            }

//            if (File.Exists(speedSavefname) && kparms != null)
//            {
//                // 读取速度文件
//                string[] sppedTexts = File.ReadAllLines(speedSavefname);
//                double[] speedValues = new double[sppedTexts.Length];
//                for (int i = 0; i < sppedTexts.Length; i++)
//                {
//                    string[] parts = sppedTexts[i].Split(' ');
//                    if (parts.Length == 2)
//                    {
//                        speedValues[i] = double.Parse(parts[1]);
//                    }
//                }

//                double[] iriDataSpped = new double[val.Length];
//                double oldSpeedVal = 0;


//                for (int i = 0; i < val.Length; i++)
//                {
//                    //int split = (int)(10.0 / disval);
//                    //int sppedIdx = i / split;
//                    //double speedval = 0;
//                    //if (sppedIdx < speedValues.Length)
//                    //{
//                    //    speedval = speedValues[sppedIdx];
//                    //}
//                    //else
//                    //{
//                    //    speedval = oldSpeedVal;
//                    //}

//                    //oldSpeedVal = speedval;

//                    // double kparm = kparms[parmnum - 1];
//                    //double bparm = bparms[parmnum - 1];
//                    //for (int pi = 0; pi < parmnum; ++pi)
//                    //{
//                    //    if (speedval <= speedparms[pi])
//                    //    {
//                    //        kparm = kparms[pi];
//                    //        bparm = bparms[pi];
//                    //        break;
//                    //    }
//                    //}
//                    double kparm = kparms.Average();
//                    iriDataSpped[i] = val[i] * kparm;
//                }
//                val = iriDataSpped;
//            }



//            // 步骤4: 四舍五入（保持原逻辑）
//            if (val != null)
//            {
//                int ttlen = val.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    val[i] = Math.Round(val[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            return true;
//        }
//        public static bool GetRutMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePartD> roadpart, ref double[] lval, ref double[] rval, ref double[] sval, bool IsShow = false)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (sval != null) sval = null;
//            if (!prjinfo._IsRut) return false;

//            String[] LStrs = null;
//            String[] RStrs = null;
//            string LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
//            if (File.Exists(LRutfrname))
//            {
//                LStrs = File.ReadAllLines(LRutfrname);
//                if (LStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            if (prjinfo._RutMode == 1)
//            {
//                string RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
//                if (File.Exists(RRutfrname))
//                {
//                    RStrs = File.ReadAllLines(RRutfrname);
//                    if (RStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                        }
//                    }
//                }
//                else
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧车辙深度数据!\r\n请检查数据完整性！");
//                    }
//                    return false;
//                }
//            }

//            const double BaseLen = 0.1;
//            int len = 0;
//            if (roadpart.Count > 0)
//                len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            sval = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;

//            for (int i = 0; i < len; i++)
//            {
//                startidx = (int)Math.Round(roadpart[i].dmi * prjinfo._DMIScale / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi * prjinfo._DMIScale / BaseLen);
//                int lvalnum = 0, rvalnum = 0, svalnum = 0;
//                double suml = 0, sumr = 0, sums = 0, ltval = 0, rtval = 0;
//                string[] tval;
//                for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                {
//                    if (prjinfo._RutMode == 1)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tval = LStrLine.Split(',');
//                            ltval = Math.Abs(double.Parse(tval[1])) + _Setting.rutLeftCorrect;
//                            if (double.IsNaN(ltval))
//                                ltval = 0;
//                            suml += ltval;
//                            ++lvalnum;
//                        }
//                        if (ValStridx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[ValStridx];
//                            tval = RStrLine.Split(',');
//                            rtval = Math.Abs(double.Parse(tval[1])) + _Setting.rutRightCorrect;
//                            if (double.IsNaN(rtval))
//                                rtval = 0;
//                            sumr += rtval;

//                            ++rvalnum;
//                        }
//                        if (double.IsNaN(ltval))
//                        {
//                            ltval = 0;
//                        }
//                        if (double.IsNaN(rtval))
//                        {
//                            rtval = 0;
//                        }
//                        sums += Math.Max(ltval, rtval);
//                        ++svalnum;
//                    }
//                    else
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tval = LStrLine.Split(',');
//                            ltval = Math.Abs(double.Parse(tval[1])) + _Setting.rutLeftCorrect;
//                            rtval = Math.Abs(double.Parse(tval[3])) + _Setting.rutRightCorrect;
//                            if (double.IsNaN(ltval))
//                            {
//                                ltval = 0;
//                            }
//                            if (double.IsNaN(rtval))
//                            {
//                                rtval = 0;
//                            }
//                            suml += ltval;
//                            sumr += rtval;
//                            sums += Math.Max(ltval, rtval);
//                            ++lvalnum;
//                            ++rvalnum;
//                            ++svalnum;
//                        }
//                    }
//                }

//                if (lvalnum > 0) suml /= lvalnum;
//                if (rvalnum > 0) sumr /= rvalnum;
//                if (svalnum > 0) sums /= svalnum;

//                if (lvalnum > 0) lval[i] = suml;
//                else if (rvalnum > 0) lval[i] = sumr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (rvalnum > 0) rval[i] = sumr;
//                else if (lvalnum > 0) rval[i] = suml;
//                else rval[i] = i > 0 ? rval[i - 1] : 0;

//                if (svalnum > 0) sval[i] = sums;
//                else sval[i] = i > 0 ? sval[i - 1] : 0;
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (sval != null)
//            {
//                int ttlen = sval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    if (_Setting.czJudgeType == 0)
//                    {
//                        sval[i] = (lval[i] + rval[i]) / 2;
//                    }
//                    else if (_Setting.czJudgeType == 1)
//                    {
//                        sval[i] = Math.Max(lval[i], rval[i]);
//                    }

//                    sval[i] = Math.Round(sval[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }
//        public static bool GetRutMaxVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePartD> roadpart, ref double[] lval, ref double[] rval, ref double[] sval)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (!prjinfo._IsRut) return false;

//            String[] LStrs = null;
//            String[] RStrs = null;
//            string LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
//            if (File.Exists(LRutfrname))
//            {
//                LStrs = File.ReadAllLines(LRutfrname);
//                if (LStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                }
//            }
//            else
//            {
//                if (prjinfo._RutMode == 1)

//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
//                }
//                else
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少车辙深度数据!\r\n请检查数据完整性！");
//                }

//                return false;
//            }

//            if (prjinfo._RutMode == 1)
//            {
//                string RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
//                if (File.Exists(RRutfrname))
//                {
//                    RStrs = File.ReadAllLines(RRutfrname);
//                    if (RStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                    }
//                }
//                else
//                {
//                    if (prjinfo._RutMode == 1)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧车辙深度数据!\r\n请检查数据完整性！");
//                    }
//                    else
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少车辙深度数据!\r\n请检查数据完整性！");
//                    }

//                    return false;
//                }
//            }

//            const double BaseLen = 0.1;
//            int len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            sval = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            for (int i = 0; i < len; i++)
//            {
//                startidx = (int)Math.Round(roadpart[i].dmi * prjinfo._DMIScale / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi * prjinfo._DMIScale / BaseLen);
//                int lvalnum = 0, rvalnum = 0, svalnum = 0;
//                double maxl = 0, maxr = 0, maxs = 0, ltval = 0, rtval = 0;
//                string[] trut;
//                for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                {
//                    if (prjinfo._RutMode == 1)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            trut = LStrLine.Split(',');
//                            ltval = Math.Abs(double.Parse(trut[1]));
//                            maxl = Math.Max(maxl, ltval);
//                            ++lvalnum;
//                        }
//                        if (ValStridx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[ValStridx];
//                            trut = RStrLine.Split(',');
//                            rtval = Math.Abs(double.Parse(trut[1]));
//                            maxr = Math.Max(maxr, rtval);
//                            ++rvalnum;
//                        }
//                        maxs = Math.Max(maxs, Math.Max(ltval, rtval));
//                        ++svalnum;
//                    }
//                    else
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            trut = LStrLine.Split(',');
//                            ltval = Math.Abs(double.Parse(trut[1]));
//                            rtval = Math.Abs(double.Parse(trut[3]));
//                            maxl = Math.Max(maxl, ltval);
//                            maxr = Math.Max(maxr, rtval);
//                            maxs = Math.Max(maxs, Math.Max(ltval, rtval));
//                            ++svalnum;
//                            ++lvalnum;
//                            ++rvalnum;
//                        }
//                    }
//                }

//                if (lvalnum > 0) lval[i] = maxl;
//                else if (rvalnum > 0) lval[i] = maxr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (rvalnum > 0) rval[i] = maxr;
//                else if (lvalnum > 0) rval[i] = maxl;
//                else rval[i] = i > 0 ? rval[i - 1] : 0;

//                if (svalnum > 0) sval[i] = maxs;
//                else sval[i] = i > 0 ? sval[i - 1] : 0;
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (sval != null)
//            {
//                int ttlen = sval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    sval[i] = Math.Round(sval[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }


//        public static bool GetMarkInfo(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref string[] MarkInfo)
//        {
//            int len = roadpart.Count;
//            MarkInfo = null;
//            MarkInfo = new string[len];
//            string markfname = string.Format(@"{0}\RoadStatuMarkInfo.txt", prjdir.FullName);
//            if (!File.Exists(markfname))
//            {
//                return false;
//            }
//            List<MarkInfo> marklist = new List<MarkInfo>();
//            string[] markstrs = File.ReadAllLines(markfname, Encoding.UTF8);
//            foreach (string line in markstrs)
//            {
//                string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
//                if (s.Length > 4)
//                {
//                    int dmival = 0;
//                    try
//                    {
//                        dmival = Convert.ToInt32(s[2].Replace("K", "").Replace("+", "").Replace("k", ""));
//                    }
//                    catch
//                    {
//                        continue;
//                    }
//                    MarkInfo markinfo = new MarkInfo();
//                    markinfo._Mile = prjinfo.Dmi2Mile(dmival);
//                    markinfo._Info = s[4];
//                    markinfo._Type = s[3];
//                    marklist.Add(markinfo);
//                }
//            }
//            if (prjinfo._DirectionInt > 0)//升序
//            {
//                marklist.Sort(delegate (MarkInfo x, MarkInfo y) { return x._Mile.CompareTo(y._Mile); });
//            }
//            else if (prjinfo._DirectionInt < 0)//降序
//            {
//                marklist.Sort(delegate (MarkInfo x, MarkInfo y) { return y._Mile.CompareTo(x._Mile); });
//            }

//            for (int i = 0; i < len - 1; ++i)
//            {
//                foreach (MarkInfo mark in marklist)
//                {
//                    if (prjinfo._DirectionInt > 0)
//                    {
//                        if (_Setting.roadCrossingShow)
//                        {
//                            if (roadpart[i].mile <= mark._Mile && roadpart[i + 1].mile > mark._Mile)
//                            {
//                                MarkInfo[i] = string.Format("{0}K{1}+{2:000} {3} {4}\r\n", MarkInfo[i], mark._Mile / 1000, mark._Mile % 1000, mark._Type, mark._Info);
//                            }
//                        }

//                    }
//                    if (prjinfo._DirectionInt < 0)
//                    {
//                        if (_Setting.roadCrossingShow)
//                        {
//                            if (roadpart[i].mile >= mark._Mile && roadpart[i + 1].mile < mark._Mile)
//                            {
//                                MarkInfo[i] = string.Format("{0}K{1}+{2:000} {3} {4}\r\n", MarkInfo[i], mark._Mile / 1000, mark._Mile % 1000, mark._Type, mark._Info);
//                            }
//                        }

//                    }
//                }
//            }
//            return true;
//        }


//        public static bool GetRutDisVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] sval, ref int[] smile)
//        {
//            sval = null;
//            smile = null;
//            if (!prjinfo._IsRut) return false;

//            String[] LStrs = null;
//            String[] RStrs = null;
//            string LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
//            if (File.Exists(LRutfrname))
//            {
//                LStrs = File.ReadAllLines(LRutfrname);
//                if (LStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                }
//            }
//            else
//            {
//                MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
//                return false;
//            }
//            if (prjinfo._RutMode == 1)
//            {
//                string RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
//                if (File.Exists(RRutfrname))
//                {
//                    RStrs = File.ReadAllLines(RRutfrname);
//                    if (RStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                    }
//                }
//                else
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧车辙深度数据!\r\n请检查数据完整性！");
//                    return false;
//                }
//            }

//            const double BaseLen = 0.1;
//            int len = roadpart.Count - 1;
//            sval = new double[len];
//            smile = new int[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            for (int i = 0; i < len; i++)
//            {
//                if (roadpart[i].roadtype == 0)
//                {
//                    startidx = (int)Math.Round(roadpart[i].dmi * prjinfo._DMIScale / BaseLen);
//                    endidx = (int)Math.Round(roadpart[i + 1].dmi * prjinfo._DMIScale / BaseLen);
//                    int svalnum = 0;
//                    double suml = 0, sumr = 0, sums = 0, ltval = 0, rtval = 0;
//                    string[] tval;
//                    for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                    {
//                        if (prjinfo._RutMode == 1)
//                        {
//                            if (ValStridx < LStrs.Length)
//                            {
//                                LStrLine = LStrs[ValStridx];
//                                tval = LStrLine.Split(',');
//                                ltval = Math.Abs(double.Parse(tval[1]));
//                                suml += ltval;
//                            }
//                            if (ValStridx < RStrs.Length)
//                            {
//                                RStrLine = RStrs[ValStridx];
//                                tval = RStrLine.Split(',');
//                                rtval = Math.Abs(double.Parse(tval[1]));
//                                sumr += rtval;
//                            }
//                            sums += Math.Max(ltval, rtval);
//                            ++svalnum;
//                        }
//                        else
//                        {
//                            if (ValStridx < LStrs.Length)
//                            {
//                                LStrLine = LStrs[ValStridx];
//                                tval = LStrLine.Split(',');
//                                ltval = Math.Abs(double.Parse(tval[1]));
//                                rtval = Math.Abs(double.Parse(tval[3]));
//                                suml += ltval;
//                                sumr += rtval;
//                                sums += Math.Max(ltval, rtval);
//                                ++svalnum;
//                            }
//                        }
//                    }
//                    if (svalnum > 0)
//                    {
//                        sval[i] = sums / svalnum;
//                        smile[i] = roadpart[i].mile;
//                    }
//                    else
//                    {
//                        sval[i] = 0;
//                        smile[i] = roadpart[i].mile;
//                    }
//                }
//                else
//                {
//                    sval[i] = 0;
//                    smile[i] = roadpart[i].mile;
//                }

//                if (smile[i] == 0)
//                {
//                    int ttt = 0;
//                    ++ttt;
//                }
//            }

//            if (sval != null)
//            {
//                int ttlen = sval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    sval[i] = Math.Round(sval[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }


//        //获取工程中所有病害  拉框模式
//        public static void GetAllDis(string projectpath, ProjectInfo prjinfo, int direction, Dictionary<string, int> RoadGradeDict,
//            double[] sval, int[] smile, ref Disease[] arrdis,
//            ref Disease[] arrdisrepair, double[] rutthresh, List<MilePart> mileSelct, double[] svalR = null, bool IsCity = false)
//        {
//            try
//            {
//                int allMileLen = mileSelct.Count - 1;
//                string errlog = projectpath + "\\errlog.txt";
//                List<Disease> disease = new List<Disease>();
//                List<Disease> disrepair = new List<Disease>();

//                string[] ImgMilestr = null;
//                if (File.Exists(projectpath + "\\RoadImg\\Camera0\\Road2Mile.txt"))
//                {
//                    ImgMilestr = File.ReadAllLines(projectpath + "\\RoadImg\\Camera0\\Road2Mile.txt");

//                    int temp = 0;
//                    bool tfalg = true;
//                    foreach (string infostr in ImgMilestr)
//                    {
//                        string[] s = infostr.Split(' ');

//                        //读取工程图像大小，用于计算病害框的真实尺寸，注：经过预处理的图像和原始图像大小不同
//                        if (tfalg)
//                        {
//                            string timgname = string.Format("{0}\\RoadImg\\Camera0{1}", projectpath, s[1]);
//                            if (File.Exists(timgname))
//                            {
//                                using (FileStream fs = new FileStream(timgname, FileMode.Open, FileAccess.Read))
//                                {
//                                    System.Drawing.Image _image = System.Drawing.Image.FromStream(fs);
//                                    _RoadConfig.ImageWidth = _image.Width;
//                                    _RoadConfig.ImageHeight = _image.Height;
//                                    // if (!_Setting.hasCamsetting)
//                                    {
//                                        _RoadConfig.WidthScale = _RoadConfig.RealWidth * 1.0 / _RoadConfig.ImageWidth;
//                                        _RoadConfig.HeightScale = _RoadConfig.RealHeight * 1.0 / _RoadConfig.ImageHeight;
//                                    }
//                                    _image.Dispose();
//                                    _image = null;
//                                }
//                                tfalg = false;
//                            }
//                        }
//                        //  _PartClass.txt
//                        string disfile = string.Format("{0}\\RoadImg\\Camera0{1}.txt", projectpath, s[1]);
//                        temp = s[1].LastIndexOf('\\');
//                        string tname = s[1].Substring(temp + 1);
//                        string tpath = "\\RoadImg\\Camera0" + s[1].Substring(0, temp);
//                        int imgmile = (int)Math.Round(Convert.ToDouble(s[0]));
//                        if (prjinfo._DirectionInt > 0 && (imgmile > mileSelct[allMileLen].mile || imgmile < mileSelct[0].mile)
//                || prjinfo._DirectionInt < 0 && (imgmile < mileSelct[allMileLen].mile || imgmile > mileSelct[0].mile))
//                            continue;

//                        if (File.Exists(disfile))
//                        {
//                            string[] dises = File.ReadAllLines(disfile);
//                            foreach (string dis in dises)
//                            {
//                                try
//                                {

//                                    Disease tdis = new Disease(dis, imgmile);
//                                    if (tdis.Area > 0)
//                                    {
//                                        tdis.imgname = tname;
//                                        tdis.imgpath = tpath;
//#if 辽宁建祥3m

//                                        int split1 = (_RoadConfig.ImageHeight - tdis.rect.Height) / 3;
//                                        int split2 = (_RoadConfig.ImageHeight - tdis.rect.Height) * 2 / 3;
//                                        if (tdis.rect.Y > split1 && tdis.rect.Y < split2)
//                                        {
//                                            tdis.m_mile += direction;
//                                        }
//                                        else if (tdis.rect.Y > split2)
//                                        {
//                                            tdis.m_mile = tdis.m_mile + direction * 2;
//                                        }
//#else
//                                        if (tdis.Area <= _RoadConfig.DetectWidth * 2 * 2 / 3)
//                                        {
//                                            if (tdis.rect.Y > (_RoadConfig.ImageHeight - tdis.rect.Height) / 2)
//                                            {
//                                                tdis.m_mile += direction;
//                                            }
//                                        }

//                                        //if (!tdis.RoadDisType.Contains("破碎板") && !tdis.RoadDisType.Contains("松散")
//                                        //    && !tdis.RoadDisType.Contains("露骨") && !tdis.RoadDisType.Contains("网裂"))
//                                        //{
//                                        //    if (tdis.rect.Y > (_RoadConfig.ImageHeight - tdis.rect.Height) / 2)
//                                        //    {
//                                        //        tdis.m_mile += direction;
//                                        //    }
//                                        //}
//                                        //else
//                                        //{
//                                        //    if (tdis.Area <= _RoadConfig.DetectWidth * 2 * 2 / 3)
//                                        //    {
//                                        //        if (tdis.rect.Y > (_RoadConfig.ImageHeight - tdis.rect.Height) / 2)
//                                        //        {
//                                        //            tdis.m_mile += direction;
//                                        //        }
//                                        //    }
//                                        //    else
//                                        //    {

//                                        //    }
//                                        //}

//#endif
//                                        if (tdis.m_mile >= 0)
//                                        {
//                                            if (_Setting.IsRepair = true && _Setting.ParmStyle == StandardParmType.CityRoad && tdis.RoadDisType == "修补")
//                                            {
//                                                disrepair.Add(tdis);
//                                            }
//                                            else
//                                            {
//                                                disease.Add(tdis);
//                                            }
//                                        }
//                                    }
//                                } 
//                                catch (Exception)
//                                {

//                                    string errval = string.Format("病害导入错误：{0}\r\n", disfile);
//                                    File.AppendAllText(errlog, errval, Encoding.UTF8);

//                                }
//                            }
//                        }
//                    }
//                }
//                if (_Setting.OutRut == 1 || (_Setting.OutRut == 2 && (RoadGradeDict[prjinfo._RoadGrade] > 1)))
//                {
//                    if (IsCity)
//                    {
//                        GlobalExcel.GetRutDis(prjinfo, sval, smile, rutthresh, ref disease, 1);
//                        if (svalR != null)
//                        {
//                            GlobalExcel.GetRutDis(prjinfo, svalR, smile, rutthresh, ref disease, 2);
//                        }
//                    }
//                    else
//                    {
//                        GlobalExcel.GetRutDis(prjinfo, sval, smile, rutthresh, ref disease);
//                    }
//                }

//                // 如果是城镇路，要处理路框差的病害
//                if (_Setting.ParmStyle == StandardParmType.CityRoad || _Setting.ParmStyle == StandardParmType.CityRoadShanghai)
//                {
//                    if (direction > 0)
//                    {
//                        disease.Sort(delegate (Disease x, Disease y) { return x.m_mile.CompareTo(y.m_mile); });
//                    }
//                    else
//                    {
//                        disease.Sort(delegate (Disease x, Disease y) { return y.m_mile.CompareTo(x.m_mile); });
//                    }
//                    //GetLuKuangCha(projectpath, prjinfo, ref disease);
//                }

//                if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel ||
//                    _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
//                {
//                    //农村路出表之前根据规范清除不符合面积规定的 露骨，松散病害
//                    //露骨  损坏面积大于或等于20平米的连续
//                    //松散    损坏面积大于或等于20平米的连续
//                    //符合条件的 病害列表
//                    List<Disease> needDeleteDisease = new List<Disease>();

//                    //key 桩号  value 病害集合
//                    Dictionary<List<int>, List<Disease>> groupDiseases = new Dictionary<List<int>, List<Disease>>();
//                    for (int disIndex = 0; disIndex < disease.Count; disIndex++)
//                    {

//                        //根据桩号连续性对病害进行分组
//                        var curDis = disease[disIndex];
//                        if (curDis.RoadDisType.Contains("露骨") || curDis.RoadDisType.Contains("松散"))
//                        {
//                            if (groupDiseases.Count == 0)
//                            {
//                                List<int> miles = new List<int>() { curDis.m_mile };
//                                groupDiseases.Add(miles, new List<Disease>() { curDis });

//                            }
//                            else
//                            {
//                                bool findSite = false;
//                                //遍历组合 看这个放在哪一个
//                                foreach (var dis in groupDiseases)
//                                {
//                                    //判断当前病害在哪个组合里面
//                                    for (int t = 0; t < dis.Key.Count; t++)
//                                    {
//                                        if (Math.Abs(curDis.m_mile - dis.Key[t]) <= 2)
//                                        {
//                                            dis.Key.Add(curDis.m_mile);
//                                            dis.Value.Add(curDis);
//                                            findSite = true;
//                                            break;
//                                        }
//                                    }
//                                }
//                                if (!findSite)
//                                {
//                                    List<int> miles = new List<int>() { curDis.m_mile };
//                                    groupDiseases.Add(miles, new List<Disease>() { curDis });
//                                }
//                            }

//                        }
//                    }

//                    //对分组进行判断 剔除掉 面积小于20的分组


//                    foreach (var disGroup in groupDiseases)
//                    {

//                        double area = 0;
//                        for (int disIndex = 0; disIndex < disGroup.Value.Count; disIndex++)
//                        {
//                            area += disGroup.Value[disIndex].Area;
//                        }
//                        if (area < 20)
//                        {
//                            //小于20的要清除
//                            needDeleteDisease.AddRange(disGroup.Value);
//                        }
//                        else
//                        {

//                        }

//                    }

//                    //从总的病害列表剔除掉不符合的病害
//                    disease = disease.Except(needDeleteDisease).ToList();


//                }


//                // 如果病害程度都按重度计算
//                if (_Setting.Qufen_dis_degree == 1)
//                {
//                    foreach (Disease dis in disease)
//                    {
//                        dis.RoadDisType = dis.RoadDisType.Replace(".轻", ".重");
//                        dis.RoadDisType = dis.RoadDisType.Replace(".中", ".重");
//                    }
//                }
//                arrdisrepair = disrepair.ToArray();
//                arrdis = disease.ToArray();//转成数组，加快访问速度
//                if (direction > 0)
//                {
//                    Array.Sort(arrdis, delegate (Disease x, Disease y) { return x.m_mile.CompareTo(y.m_mile); });
//                    Array.Sort(arrdisrepair, delegate (Disease x, Disease y) { return x.m_mile.CompareTo(y.m_mile); });
//                }
//                else
//                {
//                    Array.Sort(arrdis, delegate (Disease x, Disease y) { return y.m_mile.CompareTo(x.m_mile); });
//                    Array.Sort(arrdisrepair, delegate (Disease x, Disease y) { return y.m_mile.CompareTo(x.m_mile); });
//                }
//            }
//            catch (System.Exception ex)
//            {
//            }

//        }
//        //根据车辙的数据重新计算路框差病害
//        //public static void GetLuKuangCha(string projectpath, ProjectInfo prjinfo, ref List<Disease> arrdis)
//        //{
//        //    if (prjinfo._IsRut)
//        //    {
//        //        if (!LoadRutData(projectpath, prjinfo))
//        //        {
//        //            return;
//        //        }
//        //    }
//        //    else
//        //    {
//        //        return;
//        //    }

//        //    double thresh = 15;
//        //    double depth = -1;
//        //    if (arrdis.Count > 0)
//        //    {
//        //        Disease olddis = new Disease();
//        //        Disease curdis = new Disease();
//        //        olddis = arrdis[0];
//        //        for (int i = 0; i < arrdis.Count; ++i)
//        //        {
//        //            if (arrdis[i].RoadDisType.Contains("路框差"))
//        //            {
//        //                bool isOne = false; //是否两个框是同一个井盖
//        //                if (i > 0)
//        //                {
//        //                    curdis = arrdis[i];

//        //                    int midxdiff = Math.Abs((olddis.rect.Left + olddis.rect.Right) - (curdis.rect.Left + curdis.rect.Right)) / 2;
//        //                    int ydiff = Math.Abs(_RoadConfig.ImageHeight + curdis.rect.Top - olddis.rect.Bottom);

//        //                    // 判断是不是一个井盖跨了两张照片
//        //                    if ((olddis.imgname != curdis.imgname)   //两张照片的名称不一样
//        //                       && (midxdiff < 100) //两个框的中线线在100个像素以内
//        //                        && (ydiff < 100))  //上一个框底和下一个框的顶距离在100个像素以内
//        //                    {
//        //                        isOne = true;

//        //                        // 如果跨了两张照片，就需要重新计算上一个框的实际长度，重新计算这个框的深度
//        //                        arrdis[i - 1].rect.Height += curdis.rect.Height;
//        //                        arrdis[i - 1].realheight = arrdis[i - 1].rect.Height * _RoadConfig.HeightScale;
//        //                        depth = IsLuKuangCha(projectpath, prjinfo, arrdis[i]);
//        //                        arrdis[i - 1].depth = depth;
//        //                    }
//        //                }
//        //                olddis = arrdis[i];

//        //                if (!isOne)
//        //                {
//        //                    // 如果没有跨两张照片，就直接将当前的框纳入计算
//        //                    depth = IsLuKuangCha(projectpath, prjinfo, arrdis[i]);
//        //                    arrdis[i].depth = depth;
//        //                    if (depth < thresh)
//        //                    {
//        //                        arrdis.RemoveAt(i);
//        //                        --i;
//        //                    }
//        //                }
//        //                else
//        //                {
//        //                    // 如果跨了两张照片，将这个框剔除
//        //                    arrdis.RemoveAt(i);
//        //                    --i;
//        //                }

//        //            }
//        //        }
//        //    }
//        //}
//        //判断dis路框差的深度，如果深度大于15mm，返回深度值，否则返回-1
//        //private static double IsLuKuangCha(string projectpath, ProjectInfo prjinfo, Disease dis)
//        //{
//        //    double depth = -1;
//        //    int offlinenum = -1;
//        //    int linesidx = 0;
//        //    int linenum = 12;
//        //    double rutpscale = 0.0018;
//        //    int spidx = 0;
//        //    int epidx = 0;
//        //    int side = 0;
//        //    double rutdis = 1.21;//双车辙的间距，是1.21m

//        //    string[] strs = (dis.imgpath + "\\" + dis.imgname).Split('_');
//        //    linesidx = int.Parse(strs[strs.Length - 2].Replace("\\", ""));
//        //    linesidx = (int)((linesidx * prjinfo._RoadImgDis + dis.rect.Y * _RoadConfig.HeightScale) * 10 + offlinenum);
//        //    if (linesidx < 0)
//        //        linesidx = 0;

//        //    double cx = (dis.rect.Left + dis.rect.Right) * _RoadConfig.WidthScale / 2;

//        //    if (prjinfo._RutMode == 1)
//        //    {
//        //        if (cx <= _RoadConfig.RealWidth / 2)
//        //        {
//        //            side = 0;
//        //            rutpscale = 0.0018;
//        //            cx = cx + rutdis / 2;
//        //            spidx = (int)((cx - 0.5) / rutpscale);
//        //            epidx = (int)((cx + 0.5) / rutpscale);
//        //        }

//        //        else
//        //        {
//        //            side = 1;
//        //            rutpscale = 0.0018;
//        //            cx = cx - rutdis / 2;
//        //            spidx = (int)((cx - 0.5) / rutpscale);
//        //            epidx = (int)((cx + 0.5) / rutpscale);
//        //        }
//        //    }
//        //    else
//        //    {
//        //        side = 0;
//        //        rutpscale = 0.0018;
//        //        spidx = (int)((cx - 0.5) / rutpscale);
//        //        epidx = (int)((cx + 0.5) / rutpscale);
//        //    }

//        //    depth = GetDepth(prjinfo, linesidx, linenum, side, spidx, epidx, projectpath, dis);
//        //    return depth;
//        //}
//        //private static double GetDepth(ProjectInfo prjinfo, int linesidx, int linenum, int side, int spidx, int epidx, string prjpath, Disease curdis)
//        //{
//        //    string oriRutPath = prjpath + curdis.imgpath + "\\" + curdis.imgname + "." + curdis.m_mile.ToString() + ".CPData";
//        //    if (File.Exists(oriRutPath))
//        //    {
//        //        File.Delete(oriRutPath);
//        //    }

//        //    List<double> depthlist = new List<double>();
//        //    if (prjinfo._RutMode == 1)//双车辙
//        //    {
//        //        if (side == 0)
//        //        {
//        //            ReadRutData(rutfilepaths_L, linesidx, linenum, ref depthlist, rutparm_L._asp, rutparm_L._aep, spidx, epidx, rutparm_L._gslen, oriRutPath);
//        //        }
//        //        else if (side == 1)
//        //        {
//        //            ReadRutData(rutfilepaths_R, linesidx, linenum, ref depthlist, rutparm_R._asp, rutparm_R._aep, spidx, epidx, rutparm_R._gslen, oriRutPath);
//        //        }
//        //    }
//        //    else
//        //    {
//        //        ReadRutData(rutfilepaths_L, linesidx, linenum, ref depthlist, rutparm_L._asp, rutparm_L._cep, spidx, epidx, rutparm_L._gslen, oriRutPath);
//        //    }

//        //    return ComputeDepth(depthlist, prjpath, curdis);
//        //}

//        //private static void ReadRutData(string[] fpaths, int linesidx, int linenum, ref List<double> depthlist, int sthidx, int ethidx, int spidx, int epidx, int ftlen, string oriRutPath)
//        //{
//        //    int filelinenum = 25600;
//        //    int fileidx = (int)(linesidx / filelinenum);
//        //    if (fileidx >= fpaths.Length)
//        //    {
//        //        return;
//        //    }
//        //    int lineidx = (int)(linesidx - fileidx * filelinenum);
//        //    int lineshownum = Math.Min(linenum, filelinenum - lineidx);
//        //    ReadRutData(fpaths[fileidx], lineidx, lineshownum, ref depthlist, sthidx, ethidx, spidx, epidx, ftlen, oriRutPath);
//        //    if (linenum > lineshownum)
//        //    {
//        //        fileidx = fileidx + 1;
//        //        lineidx = 0;
//        //        lineshownum = linenum - lineshownum;
//        //        if (fileidx < fpaths.Length)
//        //        {
//        //            ReadRutData(fpaths[fileidx], lineidx, lineshownum, ref depthlist, sthidx, ethidx, spidx, epidx, ftlen, oriRutPath);
//        //        }
//        //    }
//        //}
////        private static void ReadRutData(string fpath, int lineidx, int lineshownum, ref List<double> depthlist, int sthidx, int ethidx, int spidx, int epidx, int ftlen, string oriRutPath)
////        {
////            FileStream fw = new FileStream(oriRutPath, FileMode.Append);
////            StreamWriter sw = new StreamWriter(fw);

////            string tmpstr = null;

////            int linebytes = rutparm_L._hpixel * rutparm_L._pixsize;
////            using (FileStream frstream = new FileStream(fpath, FileMode.Open))
////            {
////                frstream.Seek(lineidx * linebytes, SeekOrigin.Begin);
////                for (int i = 0; i < lineshownum; ++i)
////                {
////                    if (frstream.Read(rbarr, 0, linebytes) > 0)
////                    {
////                        Buffer.BlockCopy(rbarr, 0, profile, 0, linebytes);
////                        for (int j = 0; j < rutparm_L._hpixel; ++j)
////                        {
////                            profileZ[j] = profile[j] / rutparm_L._scaleval;
////                            profileZtmp[j] = profileZ[j];

////                            if (j >= sthidx && j < ethidx)
////                            {
////                                //string temp = (i * 100).ToString("0.00");
////                                //tmpstr = (j*1.5).ToString("0.00")+ "," + temp+  "," + profileZ[j].ToString("0.00")+"\r\n";
////                                tmpstr = profileZ[j].ToString("0.00\t");
////                                sw.Write(tmpstr);
////                            }
////                        }
////                        sw.Write("\r\n");


////                        depthlist.Add(ComputeLineDepth(profileZ, spidx, epidx, sthidx, ethidx, ftlen, profileZtmp));
////                    }
////                }
////            }

////            sw.Close();
////            fw.Close();
////        }

////        // 计算一个断面的深度
////        private static double ComputeLineDepth(float[] profileZ, int sidx, int eidx, int sthidx, int ethidx, int ftlen, float[] profileZtmp)
////        {
////            double depth = 0;

////            if (sidx < sthidx)
////                sidx = sthidx;
////            if (sidx > ethidx)
////                return depth;

////            if (eidx > ethidx)
////                eidx = ethidx;
////            if (eidx < sthidx)
////                return depth;

////            float k = 1, b = 0;

////            //MyRut.pickline(ref profileZ, sthidx, ethidx, 28);
////            MyRut.MidianAverageFileter(profileZ, sthidx, ethidx, ftlen, ref profileZtmp);
////            MyRut.leastsquare(profileZtmp, sthidx, ethidx, ref k, ref b);
////            MyRut.distanceline(ref profileZtmp, sthidx, ethidx, k, b);

////            depth = MyRut.getmaxmin(profileZtmp, sidx, eidx, Math.Abs(sidx - eidx));
////            return depth;
////        }
////        //获取工程中所有病害 小方格模式
////        public static void GetSmallRectAllDis(string projectpath, ProjectInfo prjinfo, int direction, Dictionary<string, int> RoadGradeDict,
////            double[] sval, int[] smile, ref SmalRectDisease[] arrdis, ref SmalRectDisease[] arrdisrepair, double[] rutthresh, List<MilePart> mileSelct)
////        {
////            int allMileLen = mileSelct.Count - 1;
////            string errlog = projectpath + "\\errlog.txt";
////            List<SmalRectDisease> disease = new List<SmalRectDisease>();
////            List<SmalRectDisease> disrepair = new List<SmalRectDisease>();

////            string[] ImgMilestr = null;
////            if (File.Exists(projectpath + "\\RoadImg\\Camera0\\Road2Mile.txt"))
////            {
////                ImgMilestr = File.ReadAllLines(projectpath + "\\RoadImg\\Camera0\\Road2Mile.txt");
////                int temp = 0;
////                bool tfalg = true;
////                foreach (string infostr in ImgMilestr)
////                {
////                    string[] s = infostr.Split(' ');

////                    //读取工程图像大小，用于计算病害框的真实尺寸，注：经过预处理的图像和原始图像大小不同
////                    if (tfalg)
////                    {
////                        string timgname = string.Format("{0}\\RoadImg\\Camera0{1}", projectpath, s[1]);
////                        if (File.Exists(timgname))
////                        {
////                            using (FileStream fs = new FileStream(timgname, FileMode.Open, FileAccess.Read))
////                            {
////                                System.Drawing.Image _image = System.Drawing.Image.FromStream(fs);
////                                _RoadConfig.ImageWidth = _image.Width;
////                                _RoadConfig.ImageHeight = _image.Height;

////                                {
////                                    _RoadConfig.WidthScale = _RoadConfig.RealWidth * 1.0 / _RoadConfig.ImageWidth;
////                                    _RoadConfig.HeightScale = _RoadConfig.RealHeight * 1.0 / _RoadConfig.ImageHeight;
////                                }
////                                _image.Dispose();
////                                _image = null;
////                            }
////                            tfalg = false;
////                        }
////                    }
////                    //  _PartClass.txt
////                    string disfile = string.Format("{0}\\RoadImg\\Camera0{1}_PartClass.txt", projectpath, s[1]);
////                    temp = s[1].LastIndexOf('\\');
////                    string tname = s[1].Substring(temp + 1);
////                    string tpath = "\\RoadImg\\Camera0" + s[1].Substring(0, temp);

////                    int imgmile = (int)Math.Round(Convert.ToDouble(s[0]));
////                    if (prjinfo._Direction > 0 && (imgmile > mileSelct[allMileLen].mile || imgmile < mileSelct[0].mile)
////                   || prjinfo._Direction < 0 && (imgmile < mileSelct[allMileLen].mile || imgmile > mileSelct[0].mile))
////                        continue;

////                    if (File.Exists(disfile))
////                    {
////                        string[] dises = File.ReadAllLines(disfile);
////                        foreach (string dis in dises)
////                        {
////                            try
////                            {
////                                SmalRectDisease tdis = new SmalRectDisease(dis, imgmile);
////                                if (tdis.isDiseaseOK)
////                                {
////                                    if (tdis.Area > 0)
////                                    {
////                                        tdis.imgname = tname;
////                                        tdis.imgpath = tpath;
////                                        // 如果框在图片中心点下面，则桩号加一（方向为负，则桩号减一）
////#if 辽宁建祥3m
////                                        int splitY1 = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 3;
////                                        int splitY2 = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum * 2 / 3;
////                                        if (tdis.FirstRectNum > splitY1 && tdis.FirstRectNum < splitY2)
////                                        {
////                                            tdis.m_mile += direction;
////                                        }
////                                        else if (tdis.FirstRectNum > splitY2)
////                                        {
////                                            tdis.m_mile = tdis.m_mile + direction * 2;
////                                        }


////#else


////                                        if (tdis.FirstRectNum > (_RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 2))
////                                        {
////                                            tdis.m_mile += direction;
////                                        }
////#endif

////                                        if (_Setting.IsRepair = true && _Setting.ParmStyle == StandardParmType.CityRoad && tdis.RoadDisType == "修补")
////                                        {
////                                            disrepair.Add(tdis);
////                                        }
////                                        else
////                                        {
////                                            disease.Add(tdis);
////                                        }
////                                    }
////                                }
////                            }
////                            catch
////                            {
////                                string errval = string.Format("病害导入错误：{0}\r\n", disfile);
////                                File.AppendAllText(errlog, errval, Encoding.UTF8);
////                            }
////                        }
////                    }
////                }
////            }

////            //自动化检测，车辙系数为-，表示不测车辙
////            if (_Setting.OutRut == 1 || (_Setting.OutRut == 2 && (RoadGradeDict[prjinfo._RoadGrade] > 1)))
////            {
////                GlobalExcel.GetRutDis(prjinfo, sval, smile, rutthresh, ref disease);
////            }


////            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel || _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
////            {
////                //农村路出表之前根据规范清除不符合面积规定的 露骨，松散病害
////                //露骨  损坏面积大于或等于20平米的连续
////                //松散    损坏面积大于或等于20平米的连续
////                //符合条件的 病害列表
////                List<SmalRectDisease> needDeleteDisease = new List<SmalRectDisease>();

////                //key 桩号  value 病害集合
////                Dictionary<List<int>, List<SmalRectDisease>> groupDiseases = new Dictionary<List<int>, List<SmalRectDisease>>();
////                for (int disIndex = 0; disIndex < disease.Count; disIndex++)
////                {

////                    //根据桩号连续性对病害进行分组
////                    var curDis = disease[disIndex];
////                    if (curDis.RoadDisType.Contains("露骨") || curDis.RoadDisType.Contains("松散"))
////                    {
////                        if (groupDiseases.Count == 0)
////                        {
////                            List<int> miles = new List<int>() { curDis.m_mile };
////                            groupDiseases.Add(miles, new List<SmalRectDisease>() { curDis });

////                        }
////                        else
////                        {
////                            bool findSite = false;
////                            //遍历组合 看这个放在哪一个
////                            foreach (var dis in groupDiseases)
////                            {
////                                //判断当前病害在哪个组合里面
////                                for (int t = 0; t < dis.Key.Count; t++)
////                                {
////                                    if (Math.Abs(curDis.m_mile - dis.Key[t]) <= 2)
////                                    {
////                                        dis.Key.Add(curDis.m_mile);
////                                        dis.Value.Add(curDis);
////                                        findSite = true;
////                                        break;
////                                    }
////                                }
////                            }
////                            if (!findSite)
////                            {
////                                List<int> miles = new List<int>() { curDis.m_mile };
////                                groupDiseases.Add(miles, new List<SmalRectDisease>() { curDis });
////                            }
////                        }

////                    }
////                }

////                //对分组进行判断 剔除掉 面积小于20的分组


////                foreach (var disGroup in groupDiseases)
////                {

////                    double area = 0;
////                    for (int disIndex = 0; disIndex < disGroup.Value.Count; disIndex++)
////                    {
////                        area += disGroup.Value[disIndex].Area;
////                    }
////                    if (area < 20)
////                    {
////                        //小于20的要清除
////                        needDeleteDisease.AddRange(disGroup.Value);
////                    }
////                    else
////                    {

////                    }

////                }

////                //从总的病害列表剔除掉不符合的病害
////                disease = disease.Except(needDeleteDisease).ToList();


////            }

////            arrdisrepair = disrepair.ToArray();
////            arrdis = disease.ToArray();//转成数组，加快访问速度
////            if (direction > 0)
////            {
////                Array.Sort(arrdis, delegate (SmalRectDisease x, SmalRectDisease y) { return x.m_mile.CompareTo(y.m_mile); });
////                Array.Sort(arrdisrepair, delegate (SmalRectDisease x, SmalRectDisease y) { return x.m_mile.CompareTo(y.m_mile); });
////            }
////            else
////            {
////                Array.Sort(arrdis, delegate (SmalRectDisease x, SmalRectDisease y) { return y.m_mile.CompareTo(x.m_mile); });
////                Array.Sort(arrdisrepair, delegate (SmalRectDisease x, SmalRectDisease y) { return y.m_mile.CompareTo(x.m_mile); });
////            }
////        }  //获取工程中IRI数值
       
        
        
        
//        public static bool GetIRIMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] lval, ref double[] rval, bool IsShow = false)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (!prjinfo._IsIRIMTD) return false;
//            string[] LStrs = null;
//            string[] RStrs = null;
//            string LIRIfrname = string.Format(@"{0}\IRIMTD\DAQ0\IRI_{1}m.txt", prjdir.FullName, 10);
//            string RIRIfrname = string.Format(@"{0}\IRIMTD\DAQ1\IRI_{1}m.txt", prjdir.FullName, 10);
//            if (File.Exists(LIRIfrname))
//            {
//                LStrs = File.ReadAllLines(LIRIfrname);

//                //LStrs = LStrs.Skip(1).ToArray();

//                if (LStrs.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【平整度】计算到一半退出了软件\n请【清除结果——平整度】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度数据!\r\n请检查数据完整性，并重新计算IRM！");
//                }
//                return false;
//            }

//            if (prjinfo._IsDIRIMTD)
//            {
//                if (File.Exists(RIRIfrname))
//                {
//                    RStrs = File.ReadAllLines(RIRIfrname);
//                    //RStrs = RStrs.Skip(1).ToArray();

//                    if (RStrs.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【平整度】计算到一半退出了软件\n请【清除结果——平整度】后重新【计算IRM】!");
//                        }
//                    }
//                }
//                else
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度数据!\r\n请检查数据完整性，并重新计算IRM！");
//                    }
//                    return false;
//                }
//            }

//            const double BaseLen = 10;
//            int len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            double lastvalL = 0, lastvalR = 0;
//            for (int i = 0; i < len; i++)
//            {
//                double suml = 0, sumr = 0;
//                string[] tmtd;
//                int lvalnum = 0, rvalnum = 0;
//                startidx = (int)Math.Round((roadpart[i].dmi - 0.5) / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi / BaseLen);

//                if (startidx >= endidx)
//                {
//                    if (startidx < LStrs.Length)
//                    {
//                        LStrLine = LStrs[startidx];
//                        tmtd = LStrLine.Split(' ');
//                        if (tmtd[0] != "")
//                        {
//                            lastvalL = double.Parse(tmtd[1]);
//                        }
//                        suml += lastvalL;
//                        ++lvalnum;
//                    }
//                    if (prjinfo._IsDIRIMTD)
//                    {
//                        if (startidx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[startidx];
//                            tmtd = RStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {
//                                lastvalR = double.Parse(tmtd[1]);
//                            }
//                            sumr += lastvalR;
//                            ++rvalnum;
//                        }
//                    }
//                }
//                else
//                {
//                    for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tmtd = LStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {
//                                lastvalL = double.Parse(tmtd[1]);
//                            }
//                            suml += lastvalL;
//                            ++lvalnum;
//                        }
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                            if (ValStridx < RStrs.Length)
//                            {
//                                RStrLine = RStrs[ValStridx];
//                                tmtd = RStrLine.Split(' ');
//                                if (tmtd[0] != "")
//                                {
//                                    lastvalR = double.Parse(tmtd[1]);
//                                }
//                                sumr += lastvalR;
//                                ++rvalnum;
//                            }
//                        }
//                    }
//                }

//                if (lvalnum > 0)
//                {
//                    suml /= lvalnum;
//                }
//                if (rvalnum > 0)
//                {
//                    sumr /= rvalnum;
//                }

//                if (lvalnum > 0)
//                {
//                    lval[i] = suml;
//                }
//                else if (rvalnum > 0)
//                {
//                    lval[i] = sumr;
//                }
//                else if (i > 0)
//                {
//                    lval[i] = lval[i - 1];
//                }
//                else
//                {
//                    lval[i] = 0;
//                }

//                if (prjinfo._IsDIRIMTD)
//                {
//                    if (rvalnum > 0)
//                    {
//                        rval[i] = sumr;
//                    }
//                    else if (lvalnum > 0)
//                    {
//                        rval[i] = suml;
//                    }
//                    else if (i > 0)
//                    {
//                        rval[i] = rval[i - 1];
//                    }
//                    else
//                    {
//                        rval[i] = 0;
//                    }
//                }

//            }
//            if (len >= 2)
//            {
//                if (lval[0] == 0)
//                {
//                    lval[0] = lval[1];
//                }
//                if (lval[len - 1] == 0)
//                {
//                    lval[len - 1] = lval[len - 2];
//                }
//                if (prjinfo._IsDIRIMTD)
//                {
//                    if (rval[0] == 0)
//                    {
//                        rval[0] = rval[1];
//                    }
//                    if (rval[len - 1] == 0)
//                    {
//                        rval[len - 1] = rval[len - 2];
//                    }
//                }
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            return true;
//        }
//        //public static bool GetDeltaHVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, int side, ref double[] val)
//        //{
//        //    string fname = string.Format(@"{0}\IRIMTD\DAQ{1}\Resample.txt", prjdir.FullName, side);
//        //    string fname2 = string.Format(@"{0}\IRIMTD\DAQ{1}\resample.txt", prjdir.FullName, side);
//        //    if (!File.Exists(fname))
//        //    {
//        //        if (!File.Exists(fname2))
//        //        {  //新增简易模式
//        //            string simpfiedName = fname.Replace("Resample.txt", "DeltaHVal.txt");
//        //            if (GetDeltaHValFromLocal(simpfiedName, ref val))
//        //            {
//        //                return true;
//        //            }
//        //            if (!_Setting.isGDIriCalculate)
//        //                MessageBox.Show("缺少路面纵断面文件：" + fname);
//        //            return false;
//        //        }
//        //        else
//        //        {
//        //            fname = fname2;
//        //        }
//        //    }
//        //    string[] datastrs = File.ReadAllLines(fname);

//        //    int len = roadpart.Count - 1;
//        //    val = new double[len];
//        //    //原始断面长度 50mm  Resample.txt文件0.05米一个值

//        //    const double pluselen = 0.05;
//        //    //计算跳车断面长度 100mm
//        //    const double baselen = 0.1;
//        //    int skipnum = (int)(baselen / pluselen);
//        //    string[] tstrs;
//        //    double max = 0, min = 0, hval = 0, oriValue = 0;
//        //    int startidx = 0, endidx = 0, ValStridx = 0;
//        //    //均值滤波 
//        //    #region cwb 20230531

//        //    double[] oriDataD = new double[datastrs.Length]; //原始数据
//        //    double[] nextDataD = new double[datastrs.Length];  //滤波后的数据
//        //    for (int i = 0; i < datastrs.Length; i++)
//        //    {
//        //        string[] lineStrs = datastrs[i].Split('\t');
//        //        try
//        //        {
//        //            oriValue = double.Parse(lineStrs[2]);
//        //            oriDataD[i] = oriValue;
//        //            nextDataD[i] = oriValue;
//        //        }
//        //        catch
//        //        {

//        //            continue;
//        //        }
//        //    }
//        //    for (int i = 2; i < datastrs.Length - 2; ++i)
//        //    {
//        //        oriDataD[i] = (nextDataD[i - 2] + nextDataD[i - 1] + nextDataD[i] + nextDataD[i + 1] + nextDataD[i + 2]) / 5;
//        //    }

//        //    for (int i = 0; i < len; ++i)
//        //    {
//        //        bool HasData = false;
//        //        startidx = Math.Min((int)Math.Round((roadpart[i].dmi) / pluselen), datastrs.Length);
//        //        endidx = Math.Min((int)Math.Round(roadpart[i + 1].dmi / pluselen), datastrs.Length);
//        //        max = -100000; min = 100000; hval = 0;
//        //        for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//        //        {
//        //            HasData = true;
//        //            if (ValStridx % skipnum == 0)
//        //            {
//        //                hval = oriDataD[ValStridx];
//        //                max = Math.Max(hval, max);
//        //                min = Math.Min(hval, min);
//        //            }
//        //        }
//        //        if (HasData)
//        //        {
//        //            val[i] = max - min;
//        //        }
//        //    }
//        //    #endregion
//        //    #region 原始
//        //    //for (int i = 0; i < len; ++i)
//        //    //{
//        //    //    bool HasData = false;
//        //    //    startidx = Math.Min((int)Math.Round((roadpart[i].dmi) / pluselen), datastrs.Length);
//        //    //    endidx = Math.Min((int)Math.Round(roadpart[i + 1].dmi / pluselen), datastrs.Length);
//        //    //    max = -100000; min = 100000; hval = 0;
//        //    //    for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//        //    //    {
//        //    //        HasData = true;
//        //    //        if (ValStridx % skipnum == 0)
//        //    //        {
//        //    //            tstrs = datastrs[ValStridx].Split('\t');
//        //    //            try
//        //    //            {
//        //    //                hval = double.Parse(tstrs[2]);
//        //    //            }
//        //    //            catch (System.Exception ex)
//        //    //            {
//        //    //                continue;
//        //    //            }
//        //    //            max = Math.Max(hval, max);
//        //    //            min = Math.Min(hval, min);
//        //    //        }
//        //    //    }
//        //    //    if (HasData)
//        //    //    {
//        //    //        val[i] = max - min;
//        //    //    }
//        //    //}
//        //    #endregion
//        //    if (val != null)
//        //    {
//        //        int ttlen = val.Length;
//        //        for (int i = 0; i < ttlen; ++i)
//        //        {
//        //            #region 涂工 20230816 人工纠正因子
//        //            string rutInterveneFAactorStr = _Setting.mpdInterveneFAactor;
//        //            string[] strSplitRut = rutInterveneFAactorStr.Split(',');
//        //            double upFact = double.Parse(strSplitRut[0]);
//        //            double downFact = double.Parse(strSplitRut[1]);
//        //            if (val[i] < 50)
//        //            {
//        //                val[i] = val[i] * upFact;
//        //            }
//        //            if (val[i] < 80 && val[i] >= 50)
//        //            {
//        //                val[i] = val[i] * downFact;
//        //            }
//        //            #endregion

//        //            val[i] = Math.Round(val[i], _Setting.sheetRoundingOffNum);
//        //        }
//        //    }

//        //    return true;
//        //}
//        public static bool GetPBVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, List<MilePart> roadpart10,
//            ref int[][] val, double[] thresh, double[] LPBVal, double[] RPBVal, int maxormean, ref double[] DeltaHVal)
//        {
//            if (_Setting.isGDIriCalculate)
//            {
//                return false;
//            }
//            if (val != null) val = null;
//            if (DeltaHVal != null) DeltaHVal = null;


//            if (!prjinfo._IsIRIMTD) return false;



//            if (LPBVal == null)
//            {

//                MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度构造深度数据!\r\n请检查数据完整性！");


//                return false;
//            }

//            if (prjinfo._IsDIRIMTD)
//            {

//                if (RPBVal == null && !_Setting.isGDIriCalculate)
//                {

//                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度构造深度数据!\r\n请检查数据完整性！");


//                    return false;
//                }
//            }


//            int len = roadpart.Count - 1;
//            val = new int[len][];
//            DeltaHVal = new double[len];
//            for (int i = 0; i < len; i++)
//            {
//                val[i] = new int[thresh.Length];
//                for (int j = 0; j < thresh.Length; ++j)
//                {
//                    val[i][j] = 0;
//                }
//            }

//            int startidx = 0, endidx = 0, ValStridx = 0;
//            // for (int i = 0; i < len - 1; i++)
//            //cwb 20220913
//            #region 原始代码
//            //for (int i = 0; i < len - 1; i++)
//            //{
//            //    double ltval = 0, rtval = 0, htval = 0;

//            //    if (prjinfo._Direction > 0)
//            //    {
//            //        for (int tt = startidx; tt < roadpart10.Count - 1; ++tt)
//            //        {
//            //            if (roadpart[i].mile <= roadpart10[tt].mile)
//            //            {
//            //                startidx = tt;
//            //                break;
//            //            }
//            //        }
//            //        for (int tt = endidx; tt < roadpart10.Count - 1; ++tt)
//            //        {
//            //            if (roadpart[i + 1].mile <= roadpart10[tt].mile)
//            //            {
//            //                endidx = tt;
//            //                break;
//            //            }
//            //        }
//            //    }
//            //    else
//            //    {
//            //        for (int tt = startidx; tt < roadpart10.Count - 1; ++tt)
//            //        {
//            //            if (roadpart[i].mile >= roadpart10[tt].mile)
//            //            {
//            //                startidx = tt;
//            //                break;
//            //            }
//            //        }
//            //        for (int tt = endidx; tt < roadpart10.Count - 1; ++tt)
//            //        {
//            //            if (roadpart[i + 1].mile >= roadpart10[tt].mile)
//            //            {
//            //                endidx = tt;
//            //                break;
//            //            }
//            //        }
//            //    }

//            //    for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//            //    {
//            //        ltval = 0; rtval = 0; htval = 0;
//            //        if (ValStridx < LPBVal.Length)
//            //        {
//            //            ltval = LPBVal[ValStridx];
//            //        }
//            //        if (prjinfo._IsDIRIMTD)
//            //        {
//            //            if (ValStridx < RPBVal.Length)
//            //            {
//            //                rtval = RPBVal[ValStridx];
//            //            }
//            //        }
//            //        if (ltval != 0 && rtval != 0)
//            //        {
//            //            if (maxormean == 0)
//            //            {
//            //                htval = Math.Max(ltval, rtval);
//            //            }
//            //            else
//            //            {
//            //                htval = (ltval + rtval) / 2;
//            //            }
//            //        }
//            //        else if (ltval != 0) htval = ltval;
//            //        else if (rtval != 0) htval = rtval;
//            //        for (int j = 0; j < thresh.Length; ++j)
//            //        {
//            //            DeltaHVal[i] = htval;
//            //            if (htval < thresh[j])
//            //            {
//            //                val[i][j]++;
//            //                break;
//            //            }
//            //        }
//            //    }
//            //}
//            #endregion
//            #region 20220913 cwb
//            for (int i = 0; i < len; i++)
//            {
//                double ltval = 0, rtval = 0, htval = 0;

//                if (prjinfo._DirectionInt > 0)
//                {
//                    for (int tt = startidx; tt <= roadpart10.Count - 1; ++tt)//
//                    {
//                        if (roadpart[i].mile <= roadpart10[tt].mile)
//                        {
//                            startidx = tt;
//                            break;
//                        }
//                    }
//                    for (int tt = endidx; tt <= roadpart10.Count - 1; ++tt)//
//                    {
//                        if (roadpart[i + 1].mile <= roadpart10[tt].mile)
//                        {
//                            endidx = tt;
//                            break;
//                        }
//                    }
//                }
//                else
//                {
//                    for (int tt = startidx; tt <= roadpart10.Count - 1; ++tt)//
//                    {
//                        if (roadpart[i].mile >= roadpart10[tt].mile)
//                        {
//                            startidx = tt;
//                            break;
//                        }
//                    }
//                    for (int tt = endidx; tt <= roadpart10.Count - 1; ++tt)//
//                    {
//                        if (roadpart[i + 1].mile >= roadpart10[tt].mile)
//                        {
//                            endidx = tt;
//                            break;
//                        }
//                    }
//                }

//                for (ValStridx = startidx; ValStridx < endidx; ValStridx++)
//                {
//                    ltval = 0; rtval = 0; htval = 0;
//                    if (ValStridx < LPBVal.Length)
//                    {
//                        ltval = LPBVal[ValStridx];
//                    }
//                    if (prjinfo._IsDIRIMTD)
//                    {
//                        if (ValStridx < RPBVal.Length)
//                        {
//                            rtval = RPBVal[ValStridx];
//                        }
//                    }
//                    if (ltval != 0 && rtval != 0)
//                    {
//                        if (maxormean == 0)
//                        {
//                            htval = Math.Max(ltval, rtval);
//                        }
//                        else
//                        {
//                            htval = (ltval + rtval) / 2;
//                        }
//                    }
//                    else if (ltval != 0) htval = ltval;
//                    else if (rtval != 0) htval = rtval;
//                    for (int j = 0; j < thresh.Length; ++j)
//                    {
//                        #region real
//                        DeltaHVal[i] = htval;
//                        if (htval < thresh[j])
//                        {
//                            val[i][j]++;
//                            break;
//                        }
//                        #endregion

//                    }
//                }
//            }
//            #endregion


//            if (LPBVal != null)
//            {
//                int ttlen = LPBVal.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    LPBVal[i] = Math.Round(LPBVal[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (RPBVal != null)
//            {
//                int ttlen = RPBVal.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    RPBVal[i] = Math.Round(RPBVal[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (DeltaHVal != null)
//            {
//                int ttlen = DeltaHVal.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    DeltaHVal[i] = Math.Round(DeltaHVal[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }
//        //获取工程中MTD数值
//        public static bool GetMTDMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] lval, ref double[] rval, ref double[] cval, bool IsShow = false)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (!prjinfo._IsIRIMTD) return false;

//            string[] LMTDsr = null;
//            string[] RMTDsr = null;
//            string[] CMTDsr = null;
//            string LMTDfrname = string.Format(@"{0}\IRIMTD\Laser0\MTD_{1}m.txt", prjdir.FullName, 10);
//            string RMTDfrname = string.Format(@"{0}\IRIMTD\Laser1\MTD_{1}m.txt", prjdir.FullName, 10);
//            string CMTDfrname = string.Format(@"{0}\IRIMTD\Laser2\MTD_{1}m.txt", prjdir.FullName, 10);
//            if (File.Exists(LMTDfrname))
//            {
//                LMTDsr = File.ReadAllLines(LMTDfrname);
//                if (LMTDsr.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【构造深度MTD】计算到一半退出了软件\n请【清除结果——构造深度MTD】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧构造深度数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            if (prjinfo._IsDIRIMTD)
//            {
//                if (File.Exists(RMTDfrname))
//                {
//                    RMTDsr = File.ReadAllLines(RMTDfrname);
//                    if (RMTDsr.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【构造深度MTD】计算到一半退出了软件\n请【清除结果——构造深度MTD】后重新【计算IRM】!");
//                        }
//                    }
//                }
//                else
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧构造深度数据!\r\n请检查数据完整性！");
//                    return false;
//                }

//                if (prjinfo._IsMMTD)
//                {
//                    if (File.Exists(CMTDfrname))
//                    {
//                        CMTDsr = File.ReadAllLines(CMTDfrname);
//                        if (CMTDsr.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                        {
//                            if (!IsShow)
//                            {
//                                MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【构造深度MTD】计算到一半退出了软件\n请【清除结果——构造深度MTD】后重新【计算IRM】!");
//                            }
//                        }
//                    }
//                    else
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n缺少中线构造深度数据!\r\n请检查数据完整性！");
//                        }
//                        return false;
//                    }
//                }
//            }
//            const double BaseLen = 10;
//            int len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            cval = new double[len];
//            string LStrLine, RStrLine, CStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            double lastvalL = 0, lastvalR = 0, lastvalM = 0;
//            for (int i = 0; i < len; i++)
//            {

//                double suml = 0, sumr = 0, sumc = 0;
//                int lvalnum = 0, rvalnum = 0, cvalnum = 0;
//                string[] tmtd;
//                startidx = (int)Math.Round((roadpart[i].dmi - 0.5) / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi / BaseLen);

//                if (startidx >= endidx)
//                {
//                    if (startidx < LMTDsr.Length)
//                    {
//                        LStrLine = LMTDsr[startidx];
//                        tmtd = LStrLine.Split(' ');
//                        if (tmtd[0] != "")
//                        {
//                            if (tmtd.Length > 1)

//                            {
//                                lastvalL = Math.Abs(double.Parse(tmtd[1]));
//                            }
//                            else
//                            {
//                                tmtd = LStrLine.Split('\t');
//                                if (tmtd.Length > 1)
//                                    lastvalL = Math.Abs(double.Parse(tmtd[1]));

//                            }
//                        }
//                        suml += lastvalL;
//                        ++lvalnum;
//                    }
//                    if (prjinfo._IsDIRIMTD)
//                    {
//                        if (startidx < RMTDsr.Length)
//                        {
//                            RStrLine = RMTDsr[startidx];
//                            tmtd = RStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {


//                                if (tmtd.Length > 1)

//                                {
//                                    lastvalR = Math.Abs(double.Parse(tmtd[1]));
//                                }
//                                else
//                                {
//                                    tmtd = RStrLine.Split('\t');
//                                    if (tmtd.Length > 1)
//                                        lastvalR = Math.Abs(double.Parse(tmtd[1]));

//                                }
//                            }
//                            sumr += lastvalR;
//                            ++rvalnum;
//                        }
//                        if (CMTDsr != null && startidx < CMTDsr.Length)
//                        {
//                            CStrLine = CMTDsr[startidx];
//                            tmtd = CStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {


//                                if (tmtd.Length > 1)

//                                {
//                                    lastvalM = Math.Abs(double.Parse(tmtd[1]));
//                                }
//                                else
//                                {
//                                    tmtd = CStrLine.Split('\t');
//                                    if (tmtd.Length > 1)
//                                        lastvalM = Math.Abs(double.Parse(tmtd[1]));

//                                }
//                            }
//                            sumc += lastvalM;
//                            ++cvalnum;
//                        }
//                    }
//                }
//                else
//                {
//                    for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                    {
//                        if (ValStridx < LMTDsr.Length)
//                        {
//                            LStrLine = LMTDsr[ValStridx];
//                            tmtd = LStrLine.Split(' ');
//                            if (tmtd[0] != "")
//                            {

//                                if (tmtd.Length > 1)

//                                {
//                                    lastvalL = Math.Abs(double.Parse(tmtd[1]));

//                                }
//                                else
//                                {
//                                    tmtd = LStrLine.Split('\t');
//                                    if (tmtd.Length > 1)
//                                        lastvalL = Math.Abs(double.Parse(tmtd[1]));


//                                }
//                            }
//                            suml += lastvalL + _Setting.mptLeftCorrect;
//                            ++lvalnum;
//                        }
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                            if (ValStridx < RMTDsr.Length)
//                            {
//                                RStrLine = RMTDsr[ValStridx];
//                                tmtd = RStrLine.Split(' ');
//                                if (tmtd[0] != "")
//                                {

//                                    lastvalR = Math.Abs(double.Parse(tmtd[1])) + _Setting.mptRightCorrect;


//                                    if (tmtd.Length > 1)

//                                    {
//                                        lastvalR = Math.Abs(double.Parse(tmtd[1])) + _Setting.mptRightCorrect;
//                                    }
//                                    else
//                                    {
//                                        tmtd = RStrLine.Split('\t');
//                                        if (tmtd.Length > 1)
//                                        {
//                                            lastvalR = Math.Abs(double.Parse(tmtd[1])) + _Setting.mptRightCorrect;

//                                        }

//                                    }
//                                }
//                                sumr += lastvalR;
//                                ++rvalnum;
//                            }
//                            if (CMTDsr != null && ValStridx < CMTDsr.Length)
//                            {
//                                CStrLine = CMTDsr[ValStridx];
//                                tmtd = CStrLine.Split(' ');
//                                if (tmtd[0] != "")
//                                {
//                                    if (tmtd.Length > 1)
//                                    {
//                                        lastvalM = Math.Abs(double.Parse(tmtd[1])) + _Setting.mptMidCorrect;
//                                    }
//                                }
//                                sumc += lastvalM;
//                                ++cvalnum;
//                            }
//                        }
//                    }
//                }
//                if (lvalnum > 0) suml /= lvalnum;
//                if (rvalnum > 0) sumr /= rvalnum;
//                if (cvalnum > 0) sumc /= cvalnum;

//                if (lvalnum > 0) lval[i] = suml;
//                else if (rvalnum > 0) lval[i] = sumr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (prjinfo._IsDIRIMTD)
//                {
//                    if (rvalnum > 0) rval[i] = sumr;
//                    else if (lvalnum > 0) rval[i] = suml;
//                    else rval[i] = i > 0 ? rval[i - 1] : 0;
//                    if (CMTDsr != null)
//                    {
//                        if (cvalnum > 0) cval[i] = sumc;
//                        else cval[i] = i > 0 ? cval[i - 1] : 0;
//                    }
//                }
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (cval != null)
//            {
//                int ttlen = cval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    cval[i] = Math.Round(cval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            return true;
//        }

//        //获取工程中MPD数值
//        public static bool GetMPDMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] lval, ref double[] rval, ref double[] cval, bool IsShow = false)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (!prjinfo._IsIRIMTD) return false;

//            string[] LMTDsr = null;
//            string[] RMTDsr = null;
//            string[] CMTDsr = null;
//            string LMTDfrname = string.Format(@"{0}\IRIMTD\Laser0\MPD_{1}m.txt", prjdir.FullName, 10);
//            string RMTDfrname = string.Format(@"{0}\IRIMTD\Laser1\MPD_{1}m.txt", prjdir.FullName, 10);
//            string CMTDfrname = string.Format(@"{0}\IRIMTD\Laser2\MPD_{1}m.txt", prjdir.FullName, 10);
//            if (File.Exists(LMTDfrname))
//            {
//                LMTDsr = File.ReadAllLines(LMTDfrname);
//                if (LMTDsr.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【构造深度MPD】计算到一半退出了软件\n请【清除结果——构造深度MPD】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧构造深度数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            if (prjinfo._IsDIRIMTD)
//            {
//                if (File.Exists(RMTDfrname))
//                {
//                    RMTDsr = File.ReadAllLines(RMTDfrname);
//                    if (RMTDsr.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【构造深度MPD】计算到一半退出了软件\n请【清除结果——构造深度MPD】后重新【计算IRM】!");
//                        }
//                    }
//                }
//                else
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧构造深度数据!\r\n请检查数据完整性！");
//                    }
//                    return false;
//                }
//                if (prjinfo._IsMMTD)
//                {
//                    if (File.Exists(CMTDfrname))
//                    {
//                        CMTDsr = File.ReadAllLines(CMTDfrname);
//                        if (CMTDsr.Length * 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                        {
//                            if (!IsShow)
//                            {
//                                MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【构造深度MPD】计算到一半退出了软件\n请【清除结果——构造深度MPD】后重新【计算IRM】!");
//                            }
//                        }
//                    }
//                    else
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n缺少中线构造深度数据!\r\n请检查数据完整性！");
//                        }
//                        return false;
//                    }
//                }
//            }

//            const double BaseLen = 10;
//            int len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            cval = new double[len];
//            string LStrLine, RStrLine, CStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            double lastvalL = 0, lastvalR = 0, lastvalM = 0;
//            for (int i = 0; i < len; i++)
//            {
//                double suml = 0, sumr = 0, sumc = 0;
//                int lvalnum = 0, rvalnum = 0, cvalnum = 0;
//                string[] tmtd;
//                startidx = (int)Math.Round((roadpart[i].dmi - 0.5) / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi / BaseLen);

//                if (startidx >= endidx)
//                {
//                    if (startidx < LMTDsr.Length)
//                    {
//                        LStrLine = LMTDsr[startidx];
//                        tmtd = LStrLine.Split(' ');
//                        if (tmtd[0] != "" && tmtd.Length > 1)
//                        {
//                            lastvalL = Math.Abs(double.Parse(tmtd[1]));
//                        }
//                        suml += lastvalL;
//                        ++lvalnum;
//                    }
//                    if (prjinfo._IsDIRIMTD)
//                    {
//                        if (startidx < RMTDsr.Length)
//                        {
//                            RStrLine = RMTDsr[startidx];
//                            tmtd = RStrLine.Split(' ');
//                            if (tmtd[0] != "" && tmtd.Length > 1)
//                            {
//                                lastvalR = Math.Abs(double.Parse(tmtd[1]));
//                            }
//                            sumr += lastvalR;
//                            ++rvalnum;
//                        }
//                        if (CMTDsr != null && startidx < CMTDsr.Length)
//                        {
//                            CStrLine = CMTDsr[startidx];
//                            tmtd = CStrLine.Split(' ');
//                            if (tmtd[0] != "" && tmtd.Length > 1)
//                            {
//                                lastvalM = Math.Abs(double.Parse(tmtd[1]));
//                            }
//                            sumc += lastvalM;
//                            ++cvalnum;
//                        }
//                    }
//                }
//                else
//                {
//                    for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                    {
//                        if (ValStridx < LMTDsr.Length)
//                        {
//                            LStrLine = LMTDsr[ValStridx];
//                            tmtd = LStrLine.Split(' ');
//                            if (tmtd[0] != "" && tmtd.Length > 1)
//                            {
//                                lastvalL = Math.Abs(double.Parse(tmtd[1]));
//                            }
//                            suml += lastvalL;
//                            ++lvalnum;
//                        }
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                            if (ValStridx < RMTDsr.Length)
//                            {
//                                RStrLine = RMTDsr[ValStridx];
//                                tmtd = RStrLine.Split(' ');
//                                if (tmtd[0] != "" && tmtd.Length > 1)
//                                {
//                                    lastvalR = Math.Abs(double.Parse(tmtd[1]));
//                                }
//                                sumr += lastvalR;
//                                ++rvalnum;
//                            }
//                            if (CMTDsr != null && ValStridx < CMTDsr.Length)
//                            {
//                                CStrLine = CMTDsr[ValStridx];
//                                tmtd = CStrLine.Split(' ');
//                                if (tmtd[0] != "" && tmtd.Length > 1)
//                                {
//                                    lastvalM = Math.Abs(double.Parse(tmtd[1]));
//                                }
//                                sumc += lastvalM;
//                                ++cvalnum;
//                            }
//                        }
//                    }
//                }
//                if (lvalnum > 0) suml /= lvalnum;
//                if (rvalnum > 0) sumr /= rvalnum;
//                if (cvalnum > 0) sumc /= cvalnum;

//                if (lvalnum > 0) lval[i] = suml;
//                else if (rvalnum > 0) lval[i] = sumr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (prjinfo._IsDIRIMTD)
//                {
//                    if (rvalnum > 0) rval[i] = sumr;
//                    else if (lvalnum > 0) rval[i] = suml;
//                    else rval[i] = i > 0 ? rval[i - 1] : 0;
//                    if (CMTDsr != null)
//                    {
//                        if (cvalnum > 0) cval[i] = sumc;
//                        else cval[i] = i > 0 ? cval[i - 1] : 0;
//                    }
//                }
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (cval != null)
//            {
//                int ttlen = cval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    cval[i] = Math.Round(cval[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }   //获取工程中Rut数值
//        public static bool GetRutMeanVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] lval, ref double[] rval, ref double[] sval, bool IsShow = false)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (sval != null) sval = null;
//            if (!prjinfo._IsRut) return false;

//            String[] LStrs = null;
//            String[] RStrs = null;
//            string LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
//            if (File.Exists(LRutfrname))
//            {
//                LStrs = File.ReadAllLines(LRutfrname);
//                if (LStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            if (prjinfo._RutMode == 1)
//            {
//                string RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
//                if (File.Exists(RRutfrname))
//                {
//                    RStrs = File.ReadAllLines(RRutfrname);
//                    if (RStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        if (!IsShow)
//                        {
//                            MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                        }
//                    }
//                }
//                else
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧车辙深度数据!\r\n请检查数据完整性！");
//                    }
//                    return false;
//                }
//            }

//            const double BaseLen = 0.1;
//            int len = 0;
//            if (roadpart.Count > 0)
//                len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            sval = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;

//            for (int i = 0; i < len; i++)
//            {
//                startidx = (int)Math.Round(roadpart[i].dmi * prjinfo._DMIScale / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi * prjinfo._DMIScale / BaseLen);
//                int lvalnum = 0, rvalnum = 0, svalnum = 0;
//                double suml = 0, sumr = 0, sums = 0, ltval = 0, rtval = 0;
//                string[] tval;
//                for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                {


//                    if (prjinfo._RutMode == 1)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tval = LStrLine.Split(',');
//                            if (tval.Length < 1)
//                            {
//                                continue;
//                            }
//                            ltval = Math.Abs(double.Parse(tval[1])) + _Setting.rutLeftCorrect;
//                            if (double.IsNaN(ltval))
//                                ltval = 0;
//                            suml += ltval;
//                            ++lvalnum;
//                        }
//                        if (ValStridx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[ValStridx];
//                            tval = RStrLine.Split(',');
//                            if (tval.Length < 1)
//                            {
//                                continue;
//                            }
//                            rtval = Math.Abs(double.Parse(tval[1])) + _Setting.rutRightCorrect;
//                            if (double.IsNaN(rtval))
//                                rtval = 0;
//                            sumr += rtval;

//                            ++rvalnum;
//                        }
//                        if (double.IsNaN(ltval))
//                        {
//                            ltval = 0;
//                        }
//                        if (double.IsNaN(rtval))
//                        {
//                            rtval = 0;
//                        }
//                        sums += Math.Max(ltval, rtval);
//                        ++svalnum;
//                    }
//                    else
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            tval = LStrLine.Split(',');
//                            if (tval.Length < 4)
//                            {
//                                continue;
//                            }
//                            ltval = Math.Abs(double.Parse(tval[1])) + _Setting.rutLeftCorrect;
//                            rtval = Math.Abs(double.Parse(tval[3])) + _Setting.rutRightCorrect;
//                            if (double.IsNaN(ltval))
//                            {
//                                ltval = 0;
//                            }
//                            if (double.IsNaN(rtval))
//                            {
//                                rtval = 0;
//                            }
//                            suml += ltval;
//                            sumr += rtval;
//                            sums += Math.Max(ltval, rtval);
//                            ++lvalnum;
//                            ++rvalnum;
//                            ++svalnum;
//                        }
//                    }
//                }

//                if (lvalnum > 0) suml /= lvalnum;
//                if (rvalnum > 0) sumr /= rvalnum;
//                if (svalnum > 0) sums /= svalnum;

//                if (lvalnum > 0) lval[i] = suml;
//                else if (rvalnum > 0) lval[i] = sumr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (rvalnum > 0) rval[i] = sumr;
//                else if (lvalnum > 0) rval[i] = suml;
//                else rval[i] = i > 0 ? rval[i - 1] : 0;

//                if (svalnum > 0) sval[i] = sums;
//                else sval[i] = i > 0 ? sval[i - 1] : 0;
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (sval != null)
//            {
//                int ttlen = sval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    if (_Setting.czJudgeType == 0)
//                    {
//                        sval[i] = (lval[i] + rval[i]) / 2;
//                    }
//                    else if (_Setting.czJudgeType == 1)
//                    {
//                        sval[i] = Math.Max(lval[i], rval[i]);
//                    }
//                    sval[i] = Math.Round(sval[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        }

//        public static bool GetRutMaxVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref double[] lval, ref double[] rval, ref double[] sval)
//        {
//            if (lval != null) lval = null;
//            if (rval != null) rval = null;
//            if (!prjinfo._IsRut) return false;

//            String[] LStrs = null;
//            String[] RStrs = null;
//            string LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
//            if (File.Exists(LRutfrname))
//            {
//                LStrs = File.ReadAllLines(LRutfrname);
//                if (LStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                }
//            }
//            else
//            {
//                if (prjinfo._RutMode == 1)

//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
//                }
//                else
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少车辙深度数据!\r\n请检查数据完整性！");
//                }

//                return false;
//            }

//            if (prjinfo._RutMode == 1)
//            {
//                string RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
//                if (File.Exists(RRutfrname))
//                {
//                    RStrs = File.ReadAllLines(RRutfrname);
//                    if (RStrs.Length / 10 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【车辙】计算到一半退出了软件\n请【清除结果——车辙】后重新【计算IRM】!");
//                    }
//                }
//                else
//                {
//                    if (prjinfo._RutMode == 1)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧车辙深度数据!\r\n请检查数据完整性！");
//                    }
//                    else
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n缺少车辙深度数据!\r\n请检查数据完整性！");
//                    }

//                    return false;
//                }
//            }

//            const double BaseLen = 0.1;
//            int len = roadpart.Count - 1;
//            lval = new double[len];
//            rval = new double[len];
//            sval = new double[len];
//            string LStrLine, RStrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            for (int i = 0; i < len; i++)
//            {
//                startidx = (int)Math.Round(roadpart[i].dmi * prjinfo._DMIScale / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi * prjinfo._DMIScale / BaseLen);
//                int lvalnum = 0, rvalnum = 0, svalnum = 0;
//                double maxl = 0, maxr = 0, maxs = 0, ltval = 0, rtval = 0;
//                string[] trut;
//                for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                {
//                    if (prjinfo._RutMode == 1)
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            trut = LStrLine.Split(',');
//                            ltval = Math.Abs(double.Parse(trut[1]));
//                            maxl = Math.Max(maxl, ltval);
//                            ++lvalnum;
//                        }
//                        if (ValStridx < RStrs.Length)
//                        {
//                            RStrLine = RStrs[ValStridx];
//                            trut = RStrLine.Split(',');
//                            rtval = Math.Abs(double.Parse(trut[1]));
//                            maxr = Math.Max(maxr, rtval);
//                            ++rvalnum;
//                        }
//                        maxs = Math.Max(maxs, Math.Max(ltval, rtval));
//                        ++svalnum;
//                    }
//                    else
//                    {
//                        if (ValStridx < LStrs.Length)
//                        {
//                            LStrLine = LStrs[ValStridx];
//                            trut = LStrLine.Split(',');
//                            ltval = Math.Abs(double.Parse(trut[1]));
//                            rtval = Math.Abs(double.Parse(trut[3]));
//                            maxl = Math.Max(maxl, ltval);
//                            maxr = Math.Max(maxr, rtval);
//                            maxs = Math.Max(maxs, Math.Max(ltval, rtval));
//                            ++svalnum;
//                            ++lvalnum;
//                            ++rvalnum;
//                        }
//                    }
//                }

//                if (lvalnum > 0) lval[i] = maxl;
//                else if (rvalnum > 0) lval[i] = maxr;
//                else lval[i] = i > 0 ? lval[i - 1] : 0;

//                if (rvalnum > 0) rval[i] = maxr;
//                else if (lvalnum > 0) rval[i] = maxl;
//                else rval[i] = i > 0 ? rval[i - 1] : 0;

//                if (svalnum > 0) sval[i] = maxs;
//                else sval[i] = i > 0 ? sval[i - 1] : 0;
//            }

//            if (lval != null)
//            {
//                int ttlen = lval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    lval[i] = Math.Round(lval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (rval != null)
//            {
//                int ttlen = rval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    rval[i] = Math.Round(rval[i], _Setting.sheetRoundingOffNum);
//                }
//            }

//            if (sval != null)
//            {
//                int ttlen = sval.Length;
//                for (int i = 0; i < ttlen; ++i)
//                {
//                    sval[i] = Math.Round(sval[i], _Setting.sheetRoundingOffNum);
//                }
//            }
//            return true;
//        } //获取工程中几何线形数值
//        public static bool GetGeoAligVal(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
//            ref double[] curvatureVal, ref double[] crossSlopeVal, ref double[] heightSlopeVal, bool IsShow = false)
//        {
//            if (curvatureVal != null) curvatureVal = null;
//            if (crossSlopeVal != null) crossSlopeVal = null;
//            if (heightSlopeVal != null) heightSlopeVal = null;

//            if (!prjinfo._IsRut) return false;
//            if (prjinfo._GeoAlig != 1) return false;

//            String[] curvatureStrs = null;
//            String[] crossSlopeStrs = null;
//            String[] heightSlopeStrs = null;

//            string curvaturefname = string.Format(@"{0}\camera0\imu.hon.Curvature", prjdir.FullName);
//            if (File.Exists(curvaturefname))
//            {
//                curvatureStrs = File.ReadAllLines(curvaturefname);
//                if (curvatureStrs.Length / 2 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【几何线形】计算到一半退出了软件\n请【清除结果——几何线形】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少曲率数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            string crossSlopefname = string.Format(@"{0}\camera0\imu.hon.CrossSlope", prjdir.FullName);
//            if (File.Exists(crossSlopefname))
//            {
//                crossSlopeStrs = File.ReadAllLines(crossSlopefname);
//                if (crossSlopeStrs.Length / 2 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【几何线形】计算到一半退出了软件\n请【清除结果——几何线形】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少横坡数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            string heightSlopefname = string.Format(@"{0}\camera0\imu.hon.HeightSlope", prjdir.FullName);
//            if (File.Exists(heightSlopefname))
//            {
//                heightSlopeStrs = File.ReadAllLines(heightSlopefname);
//                if (heightSlopeStrs.Length / 2 < prjinfo._EndDmi - prjinfo._EndDmi / 5)
//                {
//                    if (!IsShow)
//                    {
//                        MessageBox.Show(prjdir.FullName + "\r\n上次【计算IRM】中【几何线形】计算到一半退出了软件\n请【清除结果——几何线形】后重新【计算IRM】!");
//                    }
//                }
//            }
//            else
//            {
//                if (!IsShow)
//                {
//                    MessageBox.Show(prjdir.FullName + "\r\n缺少纵坡数据!\r\n请检查数据完整性！");
//                }
//                return false;
//            }

//            const double BaseLen = 0.5;
//            int len = roadpart.Count - 1;
//            curvatureVal = new double[len];
//            crossSlopeVal = new double[len];
//            heightSlopeVal = new double[len];

//            double tval = 0;
//            string StrLine;
//            int startidx = 0, endidx = 0, ValStridx = 0;
//            for (int i = 0; i < len; i++)
//            {
//                startidx = (int)Math.Round(roadpart[i].dmi * prjinfo._DMIScale / BaseLen);
//                endidx = (int)Math.Round(roadpart[i + 1].dmi * prjinfo._DMIScale / BaseLen);

//                int curvatureValNum = 0, crossSlopeValNum = 0, heightSlopeValNum = 0;
//                double sumCurvatureVal = 0, sumCrossSlopeVal = 0, sumHeightSlopeVal = 0;

//                string[] tvalstrs;
//                for (ValStridx = startidx; ValStridx < endidx; ++ValStridx)
//                {
//                    if (ValStridx < curvatureStrs.Length)
//                    {
//                        StrLine = curvatureStrs[ValStridx];
//                        tvalstrs = StrLine.Split(',');
//                        tval = Math.Abs(double.Parse(tvalstrs[1]));
//                        sumCurvatureVal += tval;
//                        ++curvatureValNum;
//                    }
//                    if (ValStridx < crossSlopeStrs.Length)
//                    {
//                        StrLine = crossSlopeStrs[ValStridx];
//                        tvalstrs = StrLine.Split(',');
//                        tval = Math.Abs(double.Parse(tvalstrs[1]));
//                        sumCrossSlopeVal += tval;
//                        ++crossSlopeValNum;
//                    }
//                    if (ValStridx < heightSlopeStrs.Length)
//                    {
//                        StrLine = heightSlopeStrs[ValStridx];
//                        tvalstrs = StrLine.Split(',');
//                        tval = Math.Abs(double.Parse(tvalstrs[1]));
//                        sumHeightSlopeVal += tval;
//                        ++heightSlopeValNum;
//                    }
//                }

//                if (curvatureValNum > 0)
//                    curvatureVal[i] = sumCurvatureVal / curvatureValNum;

//                if (crossSlopeValNum > 0)
//                    crossSlopeVal[i] = sumCrossSlopeVal * 100 / crossSlopeValNum;

//                if (heightSlopeValNum > 0)
//                    heightSlopeVal[i] = sumHeightSlopeVal * 100 / heightSlopeValNum;
//            }
//            return true;
//        }//获取工程中GPS信息
//        public static bool GetGPSInfo(ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, ref ExcelGPS[] GPSInfo)
//        {
//            string gpsfname = string.Format(@"{0}\GPS2Mile.txt", prjdir.FullName);
//            if (!File.Exists(gpsfname))
//            {
//                MessageBox.Show("不存在GPS2Mile.txt，请进行GPS桩号匹配操作！");
//                return false;
//            }

//            string[] gpsinfostrs = null;
//            ExcelGPS[] tempinfos = null;
//            if (File.Exists(prjdir.FullName + "\\GPS2Mile.txt"))
//            {
//                gpsinfostrs = File.ReadAllLines(prjdir.FullName + "\\GPS2Mile.txt");
//                tempinfos = new ExcelGPS[gpsinfostrs.Length];
//                for (int i = 0; i < gpsinfostrs.Length; ++i)
//                {
//                    tempinfos[i] = new ExcelGPS(gpsinfostrs[i]);
//                }
//            }

//            int len = roadpart.Count;
//            int gi = 0;
//            GPSInfo = new ExcelGPS[len];
//            for (int i = 0; i < len; i++)//i区间索引，j病害索引
//            {
//                for (; gi < tempinfos.Length; ++gi)
//                {
//                    if (prjinfo._DirectionInt > 0)
//                    {
//                        if (tempinfos[gi]._mile >= roadpart[i].mile)
//                        {
//                            break;
//                        }
//                    }
//                    else
//                    {
//                        if (tempinfos[gi]._mile <= roadpart[i].mile)
//                        {
//                            break;
//                        }
//                    }
//                }
//                if (gi < tempinfos.Length)
//                {
//                    GPSInfo[i] = tempinfos[gi];
//                }
//                else
//                {
//                    GPSInfo[i] = tempinfos[tempinfos.Length - 1];
//                }
//            }
//            return true;
//        }

//    }
//}
