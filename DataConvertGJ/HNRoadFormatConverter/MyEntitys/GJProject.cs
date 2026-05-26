using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
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
        private List<string> _iriMileTexts;
        private List<double> _iriMileKms;
        private double? _iriStartKm;
        private double? _iriStepKm;

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

            return BuildIriResult(iriLeft, iriRight, space, false);
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

            List<string> Iri = BuildIriResult(iriLeft, iriRight, 10, true);

       
          

            //拼接输出路径
            string outFilePath = Path.Combine(outPath, $"{GjProjectName}_IRI_{10}m.txt");

            // 确保文件夹存在
            Directory.CreateDirectory(Path.GetDirectoryName(outFilePath)); // 自动创建所有缺失的目录

            File.WriteAllLines(outFilePath, Iri);
             


            //保存文件

        }

        private List<string> BuildIriResult(List<double> iriLeft, List<double> iriRight, int space, bool formatValue)
        {
            EnsureIriMileInfo();

            int len = _iriMileTexts != null && _iriMileTexts.Count > 0
                ? _iriMileTexts.Count
                : iriLeft.Count;

            List<string> result = new List<string>();
            for (int i = 0; i < len; i++)
            {
                string mileText = GetIriMileText(i, space);
                double left = GetReportIriValue(iriLeft, i, space);
                double right = GetReportIriValue(iriRight, i, space);

                if (formatValue)
                {
                    result.Add($"{mileText},{left.ToString("f2")},{right.ToString("f2")}");
                }
                else
                {
                    result.Add($"{mileText},{left},{right}");
                }
            }

            return result;
        }

        private double GetReportIriValue(List<double> rawIriValues, int mileIndex, int space)
        {
            if (rawIriValues == null || rawIriValues.Count == 0)
            {
                return 0;
            }

            if (_iriMileKms == null || _iriMileKms.Count == 0 || mileIndex >= _iriMileKms.Count)
            {
                return mileIndex < rawIriValues.Count ? rawIriValues[mileIndex] : rawIriValues.Last();
            }

            double startDmi = GetDmiFromStart(mileIndex);
            double endDmi = mileIndex + 1 < _iriMileKms.Count
                ? GetDmiFromStart(mileIndex + 1)
                : startDmi + space;

            int startidx = (int)Math.Round((startDmi - 0.5) / space);
            int endidx = (int)Math.Round(endDmi / space);
            startidx = Math.Max(0, startidx);
            endidx = Math.Max(0, endidx);

            if (startidx >= endidx)
            {
                return startidx < rawIriValues.Count ? rawIriValues[startidx] : rawIriValues.Last();
            }

            double sum = 0;
            int count = 0;
            for (int i = startidx; i < endidx && i < rawIriValues.Count; i++)
            {
                sum += rawIriValues[i];
                count++;
            }

            if (count > 0)
            {
                return sum / count;
            }

            return startidx < rawIriValues.Count ? rawIriValues[startidx] : rawIriValues.Last();
        }

        private double GetDmiFromStart(int mileIndex)
        {
            return Math.Round(Math.Abs((_iriMileKms[mileIndex] - _iriMileKms[0]) * 1000.0), 1);
        }

        private string GetIriMileText(int index, int space)
        {
            EnsureIriMileInfo();

            if (_iriMileTexts != null && index < _iriMileTexts.Count)
            {
                return _iriMileTexts[index];
            }

            if (_iriStartKm.HasValue)
            {
                double stepKm = _iriStepKm ?? GetDefaultStepKm(space);
                return FormatKm(_iriStartKm.Value + stepKm * index);
            }

            return (index * space).ToString(CultureInfo.InvariantCulture);
        }

        private void EnsureIriMileInfo()
        {
            if (_iriMileTexts != null)
            {
                return;
            }

            _iriMileTexts = new List<string>();
            _iriMileKms = new List<double>();

            List<IriMileFile> iriFiles = GetIriMileFiles();
            if (iriFiles.Count == 0)
            {
                return;
            }

            foreach (IriMileFile iriFile in iriFiles)
            {
                _iriMileTexts.AddRange(iriFile.MileTexts);
                _iriMileKms.AddRange(iriFile.MileKms);
            }

            if (_iriMileTexts.Count > 0 && TryParseKm(_iriMileTexts[0], out double startKm))
            {
                _iriStartKm = startKm;
            }

            if (_iriMileTexts.Count > 1
                && TryParseKm(_iriMileTexts[0], out double firstKm)
                && TryParseKm(_iriMileTexts[1], out double secondKm))
            {
                _iriStepKm = secondKm - firstKm;
            }
        }

        private List<IriMileFile> GetIriMileFiles()
        {
            List<IriMileFile> iriMileFiles = new List<IriMileFile>();
            if (!Directory.Exists(IriDirPath))
            {
                return iriMileFiles;
            }

            DirectoryInfo iriDir = new DirectoryInfo(IriDirPath);
            FileInfo[] files = iriDir.GetFiles("*IRI*.csv", SearchOption.AllDirectories)
                .Union(iriDir.GetFiles("*IRI*.txt", SearchOption.AllDirectories))
                .OrderByDescending(t => t.Name.IndexOf(GjProjectName, StringComparison.OrdinalIgnoreCase) >= 0)
                .ThenByDescending(t => t.LastWriteTime)
                .ToArray();

            foreach (FileInfo file in files)
            {
                IriMileFile iriMileFile = ReadIriMileFile(file);
                if (iriMileFile.MileTexts.Count > 0)
                {
                    iriMileFiles.Add(iriMileFile);
                }
            }

            int direction = GetIriDirection(iriMileFiles);
            if (direction < 0)
            {
                iriMileFiles = iriMileFiles.OrderByDescending(t => t.StartKm).ToList();
            }
            else
            {
                iriMileFiles = iriMileFiles.OrderBy(t => t.StartKm).ToList();
            }

            return iriMileFiles;
        }

        private static IriMileFile ReadIriMileFile(FileInfo file)
        {
            IriMileFile iriMileFile = new IriMileFile();
            foreach (string line in File.ReadLines(file.FullName, Encoding.Default))
            {
                if (TryGetFirstColumnKm(line, out double km, out string mileText))
                {
                    if (iriMileFile.MileTexts.Count == 0)
                    {
                        iriMileFile.StartKm = km;
                    }
                    iriMileFile.EndKm = km;
                    iriMileFile.MileTexts.Add(mileText);
                    iriMileFile.MileKms.Add(km);
                }
            }

            return iriMileFile;
        }

        private int GetIriDirection(List<IriMileFile> iriMileFiles)
        {
            IriMileFile file = iriMileFiles.FirstOrDefault(t => t.MileTexts.Count > 1);
            if (file != null && Math.Abs(file.EndKm - file.StartKm) > 0.000001)
            {
                return file.EndKm > file.StartKm ? 1 : -1;
            }

            return Line == 1 ? -1 : 1;
        }

        private static bool TryGetFirstColumnKm(string line, out double km, out string mileText)
        {
            km = 0;
            mileText = null;

            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            string[] parts = line.Split(new[] { ',', '\t', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return false;
            }

            string first = parts[0].Trim();
            if (!TryParseKm(first, out km))
            {
                return false;
            }

            mileText = first;
            return true;
        }

        private static bool TryParseKm(string text, out double km)
        {
            km = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            text = text.Trim().Trim('"');
            if (text.StartsWith("K", StringComparison.OrdinalIgnoreCase) && text.Contains("+"))
            {
                string[] parts = text.Substring(1).Split('+');
                if (parts.Length == 2
                    && double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double kPart)
                    && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double mPart))
                {
                    km = kPart + mPart / 1000.0;
                    return true;
                }
            }

            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out km)
                || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out km);
        }

        private double GetDefaultStepKm(int space)
        {
            return (Line == 1 ? -1 : 1) * space / 1000.0;
        }

        private static string FormatKm(double km)
        {
            return km.ToString("0.0000", CultureInfo.InvariantCulture);
        }

        private class IriMileFile
        {
            public double StartKm { get; set; }
            public double EndKm { get; set; }
            public List<string> MileTexts { get; } = new List<string>();
            public List<double> MileKms { get; } = new List<double>();
        }



         

    }
}
