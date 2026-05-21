using Framework.DBHelper;
using HNDtos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XRDataProcess
{
    public class PropertyDataManger
    {
        private PropertyDataManger() { }
        private static List<RowRoad> SelectRoads { get; set; }
        public static List<HeFeiEntity> AllDatas { get; set; }
        public static PropertyDataManger Instance { get; }
        public static List<RowRoad> AllGetRoads { get; set; }

        public static ProjectInfo cPorject = null;
        public static PropertyDataManger GetInstance(ProjectInfo project) {

            cPorject = project;
            if (AllDatas == null)
            {
                 string dbPath = AppDomain.CurrentDomain.BaseDirectory + "\\PropertyInfo.db";

                string connectStr = $" Data Source={dbPath};";
                var sqlScope = new FrmSqlSugerTestScope<HeFeiEntity>(connectStr, SqlSugar.DbType.Sqlite);
                AllDatas = sqlScope.LoadSysAdmin();
            }

            List<HeFeiEntity> datas = AllDatas.Where(
                t => project._RoadCode.Trim().Equals(t.RoadNum.Trim()) 
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
                rowRoad.RoadType = data.RoadType;
                SelectRoads.Add(rowRoad);
            }
            AllGetRoads = new List<RowRoad>(SelectRoads);
            return Instance;
        }
        public static void updateInfo(ref List<MilePart> roadpart, Dictionary<string, int> _roadTypeDic, Dictionary<string, int> _RoadGradeDict)
        { 
            roadpart.Clear();  
            if (cPorject._Direction != 1)
            { 
                roadpart.Reverse(); 
            }  
            for (int i = 0; i < AllGetRoads.Count; i++)
            {
                if (i == AllGetRoads.Count-1)
                {
                    var curMark = AllGetRoads[i];
                    MilePart prePart = new MilePart();
                   prePart.mile = curMark.StartMile;
                   prePart.unit = curMark.Unit;
                   prePart.roadtype = _roadTypeDic[curMark.RoadType];
                   prePart.width = curMark.RoadWid;
                   prePart.isPub = curMark.IsPub;
                    prePart.roaddegree = curMark.Grad;
                    roadpart.Add(prePart);

                    MilePart nextPart = new MilePart();
                    nextPart.mile = curMark.EndMile;
                    nextPart.unit = curMark.Unit;
                    nextPart.roadtype = _roadTypeDic[curMark.RoadType];
                    nextPart.width = curMark.RoadWid;
                    nextPart.isPub = curMark.IsPub;
                    nextPart.roaddegree = curMark.Grad;
                    roadpart.Add(nextPart);

                }
                else
                {
                    var curMark = AllGetRoads[i];
                    MilePart curPart = new MilePart();
                    curPart.mile = curMark.StartMile;
                    curPart.unit = curMark.Unit;
                    curPart.roadtype = _roadTypeDic[curMark.RoadType];
                    curPart.width = curMark.RoadWid;
                    curPart.isPub = curMark.IsPub;
                    curPart.roaddegree = curMark.Grad;
                    roadpart.Add(curPart);
                } 
            }
 
            //处理后  反转回来  
            if (cPorject._Direction != 1)
            {
                roadpart.Reverse();
            }
        }
    }
}
