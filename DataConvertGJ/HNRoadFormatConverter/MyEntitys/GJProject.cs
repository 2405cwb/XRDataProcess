using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HNRoadFormatConverter.MyEntitys
{
    public class GJProject
    {

        public string GjProjectName { get;  }
        public string GjDirPath { get; }

        public GJProject(string path)
        {
               DirectoryInfo di = new DirectoryInfo(path);

                GjProjectName = di.Name;
            GjDirPath = path;

            if (GjProjectName[GjProjectName.Length-1] == 'A')
            {
                Line = 0;
            }
            else
            {
                Line = 1;
            }



                IriDirPath = Path.Combine(path, "IRI");
                //获得工程起始桩号
                
            //获得RiFile地址
                RiFileDirPath = Path.Combine(path, "RIFile");
           
                DrDirPath = Path.Combine(path,"DR");

                RoadPicute = Path.Combine(path, "Images");

                // GetProjectMile();
        }

        public bool InitRoadPictureFileList()
        {

            DirectoryInfo directoryInfo = new DirectoryInfo(RoadPicute);
            if (!Directory.Exists(RoadPicute))
            {
                MessageBox.Show($"没有检测到{RoadPicute}文件夹！");
                return false;
            }

         RoadPictures= directoryInfo.GetFiles("*.jpg", searchOption: SearchOption.AllDirectories).Select(t=>t.FullName).ToList();


            return true;
        }

        public (int,int) GetPcitureSize()
        {
            if (RoadPictures.Count> 0 )
            {
                var first = RoadPictures.First();


            return     GetImageDimensions(first);
            }
            MessageBox.Show($"{GjProjectName}未找到路面图像，或图像损坏请检查！");
            return (0,0);
        }
        public static (int Width, int Height) GetImageDimensions(string imagePath)
        {
            using (var image = Image.FromFile(imagePath))
            {
                return (image.Width, image.Height);
            }
        }

        public  List<string> RoadPictures { get; private set; }

        public double Smile { get; private set; }

        public double Emile { get; private set; }
        public string DrDirPath { get; }


        public string RoadPicute { get; private set; }

        public string IriDirPath { get; }
        public string RiFileDirPath { get; }

        /// <summary>
        /// 0 上行
        /// 1 下行
        /// </summary>
        public int Line = 0;

        private void GetProjectMile()
        {
            DirectoryInfo riFileDir = new DirectoryInfo(DrDirPath);

            FileInfo[] riFiles = riFileDir.GetFiles("*-DR-*");

            List<double> mileValue = new List<double>();

            foreach (FileInfo riFile in riFiles)
            {
                string[] nameSplit = riFile.Name.Split('-');
                if (nameSplit.Length > 1)
                {
                    mileValue.Add(double.Parse(nameSplit[2]));
                    mileValue.Add(double.Parse(nameSplit[3]));
                }
            }


            mileValue.Sort();

            if (Line==0)
            {
                Smile = mileValue[0];
                Emile = mileValue[mileValue.Count - 1];
            }
            else
            {
                Smile = mileValue[mileValue.Count - 1];
                Emile = mileValue[0];
            }

        }

        public List<string> getLpFileText()
        {

            //读取RiFile文件
            DirectoryInfo riFileDir = new DirectoryInfo(RiFileDirPath);
            FileInfo[] riFiles = riFileDir.GetFiles("*-LP-*");

            if (riFiles.Length <= 0)
            {
                MessageBox.Show($"{GjProjectName}工程缺少LP文件，无法进行核验平整度工作,请检查!");
                return new List<string>();
            }

            FileInfo riFile = riFiles[0];
            string[] sdata = File.ReadAllLines(riFile.FullName);
            sdata = sdata.Skip(1).ToArray(); // 移除第一行并重新赋值 
            return sdata.ToList();

        }

        public List<string> calculateIriValue(List<string> datas,int space)
        {
            List<double> iriLeft  = IRM_Algorithm.WorkBankIRIAlgo_withSpeed(datas, 0, space, 0.1);
            List<double> iriRight  = IRM_Algorithm.WorkBankIRIAlgo_withSpeed(datas, 1, space, 0.1);

            List<string> result = new List<string>();

            for (int i = 0; i < iriLeft.Count; i++)
            {
                result.Add($"{i * space},{iriLeft[i]},{iriRight[i]}");
            }

            return result;
        }

        public void CheckIirValue(string outPath,double disVal)
        {

            //读取RiFile文件
            DirectoryInfo riFileDir = new DirectoryInfo(RiFileDirPath);

            FileInfo[] riFiles =   riFileDir.GetFiles("*-LP-*") ;

            if (riFiles.Length<=0)
            {
                MessageBox.Show($"{GjProjectName}工程缺少LP文件，无法进行核验平整度工作,请检查!");
                return;
            }

            FileInfo riFile = riFiles[0];
           

             List<double>iriLeft =   IRM_Algorithm.WorkBankIRIAlgo_withSpeed(riFile.FullName, outPath,0, 10,disVal);
            List<double> iriRight = IRM_Algorithm.WorkBankIRIAlgo_withSpeed(riFile.FullName, outPath, 1, 10,disVal);

            List<string> Iri = new List<string>();

            for (int i = 0; i < iriLeft.Count; i++)
            {
               Iri.Add($"{i * 10},{iriLeft[i].ToString("f2")},{iriRight[i].ToString("f2")}");
            }

       
          

            //拼接输出路径
            string outFilePath = Path.Combine(outPath, $"{GjProjectName}_IRI_{10}m.txt");

            // 确保文件夹存在
            Directory.CreateDirectory(Path.GetDirectoryName(outFilePath)); // 自动创建所有缺失的目录

            File.WriteAllLines(outFilePath, Iri);
             


            //保存文件

        }



         

    }
}
