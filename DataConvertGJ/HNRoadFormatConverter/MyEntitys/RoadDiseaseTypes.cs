using Farmework.Other;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace HNRoadFormatConverter.MyEntitys
{
    public class RoadDiseaseTypes
    {
        static XRSetting _Setting = XRSetting.GetInstance();

        public static RoadDiseaseType[][] roaddis;
        public static Dictionary<string, int>[] DiseaseTypeDict;

        /// <summary>
        /// {"沥青", 0},{"水泥", 1},{"砂石", 2}
        /// </summary>
        public static Dictionary<string, int> roadtypedict
            = new Dictionary<string, int> { { "沥青", 0 }, { "水泥", 1 }, { "砂石", 2 } };

        public static Dictionary<string, string>[] DisRut;
        public static int rutidx = 0, rutQthresh = 0, rutZthresh = 0;
        /// <summary>
        /// 读取配置文件病害类型
        /// </summary>
        public static void LoadRoadDisParm()
        {
            roaddis = new RoadDiseaseType[roadtypedict.Count][];
            List<RoadDiseaseType>[] dislist = new List<RoadDiseaseType>[roadtypedict.Count];
            DiseaseTypeDict = new Dictionary<string, int>[roadtypedict.Count];
            for (int i = 0; i < roadtypedict.Count; i++)
            {
                DiseaseTypeDict[i] = new Dictionary<string, int>();
                dislist[i] = new List<RoadDiseaseType>();

            }
            DisRut = new Dictionary<string, string>[1];
            DisRut[0] = new Dictionary<string, string>();

            String fn = System.Windows.Forms.Application.ExecutablePath;
            fn = fn.Remove(fn.LastIndexOf("\\")) + "\\ParaVal.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(fn);

            foreach (XmlNode rootchildnode in doc.DocumentElement.ChildNodes)
            {
                if (_Setting.ParmStyle != StandardParmType.DegreeRoad2001)
                {
                    if (rootchildnode.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
                    {
                        foreach (XmlNode node in rootchildnode.ChildNodes)
                        {
                            if (node.Name.Contains("路面病害类型"))
                            {
                                int idx = 0;
                                string rt = node.Name.Replace("路面病害类型", "");
                                foreach (XmlNode dis in node.ChildNodes)
                                {
                                    dislist[roadtypedict[rt]].Add(new RoadDiseaseType(rt, dis.Name, null,
                                        ((XmlElement)dis).GetAttribute("权重"),
                                        ((XmlElement)dis).GetAttribute("影响宽度"),
                                        ((XmlElement)dis).GetAttribute("有效长度"),
                                        ((XmlElement)dis).GetAttribute("有效面积"),
                                        ((XmlElement)dis).GetAttribute("面积公式"),
                                        ((XmlElement)dis).GetAttribute("显示"),
                                        ((XmlElement)dis).GetAttribute("快捷键")));
                                    DiseaseTypeDict[roadtypedict[rt]].Add(string.Format("{0}.{1}", rt, dis.Name), idx++);
                                }
                                roaddis[roadtypedict[rt]] = dislist[roadtypedict[rt]].ToArray();
                            }
                            if (node.Name.Contains("路面车辙病害程度阈值"))
                            {
                                DisRut[0].Add(((XmlElement)node).GetAttribute("轻度索引"), ((XmlElement)node).GetAttribute("阈值"));
                            }
                        }
                    }
                }
                else
                {
                    if (rootchildnode.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
                    {
                        foreach (XmlNode node in rootchildnode.ChildNodes)
                        {
                            if (node.Name.Contains("沥青路面病害类型"))
                            {
                                int idx = 0;
                                string rt = node.Name.Replace("路面病害类型", "");
                                foreach (XmlNode dis in node.ChildNodes)
                                {
                                    dislist[roadtypedict[rt]].Add(new RoadDiseaseType(rt, dis.Name, null,
                                        ((XmlElement)dis).GetAttribute("权重"),
                                        ((XmlElement)dis).GetAttribute("影响宽度"),
                                        ((XmlElement)dis).GetAttribute("有效长度"),
                                        ((XmlElement)dis).GetAttribute("有效面积"),
                                        ((XmlElement)dis).GetAttribute("面积公式"),
                                        ((XmlElement)dis).GetAttribute("显示"),
                                        ((XmlElement)dis).GetAttribute("快捷键")));
                                    DiseaseTypeDict[roadtypedict[rt]].Add(string.Format("{0}.{1}", rt, dis.Name), idx++);
                                }
                                roaddis[roadtypedict[rt]] = dislist[roadtypedict[rt]].ToArray();
                            }
                            else if (node.Name.Contains("水泥路面病害类型"))
                            {
                                int idx = 0;
                                string rt = node.Name.Replace("路面病害类型", "");
                                foreach (XmlNode dis in node.ChildNodes)
                                {
                                    dislist[roadtypedict[rt]].Add(new RoadDiseaseType(rt, dis.Name,
                                          double.Parse(((XmlElement)dis).GetAttribute("A系数")),
                                          double.Parse(((XmlElement)dis).GetAttribute("B系数")),
                                           ((XmlElement)dis).GetAttribute("显示"),
                                           ((XmlElement)dis).GetAttribute("快捷键")));
                                    DiseaseTypeDict[roadtypedict[rt]].Add(string.Format("{0}.{1}", rt, dis.Name), idx++);
                                }
                                roaddis[roadtypedict[rt]] = dislist[roadtypedict[rt]].ToArray();
                            }
                        }
                    }
                }
            }
            if (DisRut[0].Count > 0)
            {
                rutidx = Convert.ToInt16(DisRut[0].Keys.ElementAt(0));//轻度索引
                string thresh = DisRut[0].Values.ElementAt(0); //阈值 10和15
                string[] temp = null;
                temp = thresh.Split(' ');
                if (temp.Length > 1)
                {
                    rutQthresh = int.Parse(temp[0]);
                    rutZthresh = int.Parse(temp[1]);
                }
                else
                {
                    rutZthresh = int.Parse(temp[0]);
                }
            }
        }

        public static void LoadAutoDectRoadDisParm()
        {
            roaddis = new RoadDiseaseType[roadtypedict.Count][];
            List<RoadDiseaseType>[] dislist = new List<RoadDiseaseType>[roadtypedict.Count];
            DiseaseTypeDict = new Dictionary<string, int>[roadtypedict.Count];
            for (int i = 0; i < roadtypedict.Count; i++)
            {
                DiseaseTypeDict[i] = new Dictionary<string, int>();
                dislist[i] = new List<RoadDiseaseType>();
            }
            DisRut = new Dictionary<string, string>[1];
            DisRut[0] = new Dictionary<string, string>();

            String fn = System.Windows.Forms.Application.ExecutablePath;
            fn = fn.Remove(fn.LastIndexOf("\\")) + "\\AutoParaVal.xml";
            XmlDocument doc = new XmlDocument();
            try { doc.Load(fn); }
            catch (Exception) { return; }

            foreach (XmlNode rootchildnode in doc.DocumentElement.ChildNodes)
            {
                if (rootchildnode.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
                {
                    foreach (XmlNode node in rootchildnode.ChildNodes)
                    {
                        if (node.Name.Contains("路面病害类型"))
                        {
                            int idx = 0;
                            string rt = node.Name.Replace("路面病害类型", "");
                            foreach (XmlNode dis in node.ChildNodes)
                            {
                                dislist[roadtypedict[rt]].Add(new RoadDiseaseType(rt, dis.Name,
                                    ((XmlElement)dis).GetAttribute("权重"),
                                    ((XmlElement)dis).GetAttribute("显示"),
                                    ((XmlElement)dis).GetAttribute("快捷键")));
                                DiseaseTypeDict[roadtypedict[rt]].Add(string.Format("{0}.{1}", rt, dis.Name), idx++);
                            }
                            roaddis[roadtypedict[rt]] = dislist[roadtypedict[rt]].ToArray();
                        }
                        if (node.Name.Contains("路面车辙病害程度阈值"))
                        {
                            DisRut[0].Add(((XmlElement)node).GetAttribute("轻度索引"), ((XmlElement)node).GetAttribute("阈值"));

                        }
                    }
                }
            }
            if (DisRut[0].Count > 0)
            {
                rutidx = Convert.ToInt16(DisRut[0].Keys.ElementAt(0));//轻度索引
                string thresh = DisRut[0].Values.ElementAt(0); //阈值 10和15
                string[] temp = null;
                temp = thresh.Split(' ');
                if (temp.Length > 1)
                {
                    rutQthresh = int.Parse(temp[0]);
                    rutZthresh = int.Parse(temp[1]);
                }
                else
                {
                    rutZthresh = int.Parse(temp[0]);
                }
            }

        }

        /// <summary>
        /// 区间段内病害面积总和清零
        /// </summary>
        public static void Clear()
        {
            for (int a = 0; a < roadtypedict.Count; ++a)
            {
                if (roaddis[a] != null)
                {
                    for (int i = 0; i < roaddis[a].Length; i++)
                    {
                        roaddis[a][i].totalarea = 0;
                        roaddis[a][i].platenum = 0;
                        roaddis[a][i].totallength = 0;
                        roaddis[a][i].count = 0;
                    }
                }
            }
        }
    }
}
