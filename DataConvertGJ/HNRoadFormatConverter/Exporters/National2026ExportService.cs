using Farmework.Other.enumTools;
using HNRoadFormatConverter.Entitys;
using HNRoadFormatConverter.MyEntitys;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace HNRoadFormatConverter.Exporters
{
    /// <summary>
    /// 2026年农养国省道路况检测数据提交格式的专用导出规则。
    /// 其它地区规范不要放进这里，避免 Form1 继续堆叠各地分支。
    /// </summary>
    public static class National2026ExportService
    {
        private const int RoadImagePackageSize = 5000;
        private const string DefaultPictureSuffix = ".jpg";

        public static string BuildExportDataPath(ProjectInfo project, string targetBasePath)
        {
            string directionName = project._Direction == "A" ? "上行" : "下行";
            string roadNum = string.IsNullOrWhiteSpace(project.RoadNum)
                ? project._RoadNum
                : project.RoadNum;
            string cityName = (project._City ?? string.Empty).Replace("市", string.Empty);
            string roadCode = TrimDirectionSuffix(project.ConvertProName);

            // [年月日时分秒] 保留给客户填写；多个工程可共用同一个 EXPORTDATA。
            string rootName = $"{cityName}+省检+{roadCode}-{directionName}-{roadNum}车道-[年月日时分秒]";
            return Path.Combine(targetBasePath, rootName, "EXPORTDATA");
        }

        public static bool ExportMetricFiles(
            ProjectInfo project,
            string exportDataPath,
            ICollection<string> handledFiles,
            out string errorMessage)
        {
            errorMessage = null;
            if (project.ConvertPath == null || !project.ConvertPath.Exists)
            {
                errorMessage = $"{project._DataDir.FullName}缺少国检转换中间文件，请检查ConverSource。";
                return false;
            }

            Directory.CreateDirectory(exportDataPath);

            IEnumerable<FileInfo> iriFiles = project.ConvertPath
                .GetFiles("*IRI*.csv", SearchOption.AllDirectories)
                .Union(project.ConvertPath.GetFiles("*IRI*.txt", SearchOption.AllDirectories));

            IEnumerable<FileInfo> gpsFiles = project.ConvertPath
                .GetFiles("*GPS*.txt", SearchOption.AllDirectories)
                .Union(project.ConvertPath.GetFiles("*GPS*.csv", SearchOption.AllDirectories));

            foreach (FileInfo file in iriFiles.Concat(gpsFiles))
            {
                string targetPath = Path.Combine(exportDataPath, file.Name);
                File.Copy(file.FullName, targetPath, true);
                handledFiles?.Add(targetPath);
            }

            return true;
        }

        public static National2026PictureExportTask CreatePictureTask(
            ProjectInfo project,
            string exportDataPath,
            bool isRoadImage)
        {
            string targetDirectory = isRoadImage
                ? Path.Combine(exportDataPath, "Images", project._DataDate, project.ConvertProName, "0")
                : Path.Combine(exportDataPath, "前方图像", project._DataDate, project.ConvertProName);

            Directory.CreateDirectory(targetDirectory);

            List<PicAndMile> pictures = project.GetPicAndMiles(
                isRoadImage,
                CityModelItem.农养国省道路况检测数据提交格式_2026年);

            ApplyPictureNames(pictures, project.ConvertProName);

            return new National2026PictureExportTask(
                new DirectoryInfo(targetDirectory),
                pictures,
                isRoadImage ? National2026ImageKind.RoadImage : National2026ImageKind.FrontImage);
        }

        public static void ExportPictures(
            National2026PictureExportTask task,
            IProgress<int> progress,
            string pictureSuffix = DefaultPictureSuffix)
        {
            bool isFrontImage = task.ImageKind == National2026ImageKind.FrontImage;
            int existingRoadImageCount = 0;

            if (!isFrontImage && Directory.Exists(task.TargetDirectory.FullName))
            {
                existingRoadImageCount = Directory
                    .GetFiles(task.TargetDirectory.FullName, "*.jpg", SearchOption.AllDirectories)
                    .Length;
            }

            for (int i = 0; i < task.Pictures.Count; i++)
            {
                PicAndMile picture = task.Pictures[i];
                string targetDirectory = task.TargetDirectory.FullName;

                if (!isFrontImage)
                {
                    string secondLevelDirName = ((existingRoadImageCount + i) / RoadImagePackageSize).ToString("0");
                    targetDirectory = Path.Combine(task.TargetDirectory.FullName, secondLevelDirName);
                }

                Directory.CreateDirectory(targetDirectory);

                string targetPath = Path.Combine(targetDirectory, picture.ResultPicName + pictureSuffix);
                if (!File.Exists(picture.PicPath))
                {
                    continue;
                }

                if (isFrontImage)
                {
                    CopyJpegWithWatermark(picture.PicPath, targetPath, picture.ResultPicName);
                }
                else
                {
                    File.Copy(picture.PicPath, targetPath, true);
                }

                progress?.Report(1);
            }
        }

        private static void ApplyPictureNames(List<PicAndMile> pictures, string projectName)
        {
            for (int i = 0; i < pictures.Count; i++)
            {
                PicAndMile updated = pictures[i];
                string beforeCalibrationMile = FormatMile(updated.BeforeCalibrationMile);
                string afterCalibrationMile = FormatMile(updated.AfterCalibrationMile);
                updated.updateResultPicName(projectName + "-" + beforeCalibrationMile + "-" + afterCalibrationMile);
                pictures[i] = updated;
            }
        }

        private static string FormatMile(int value)
        {
            int kilometer = value / 1000;
            int meter = Math.Abs(value % 1000);
            return $"{kilometer.ToString().PadLeft(3, '0')}+{meter.ToString().PadLeft(3, '0')}000";
        }

        private static string TrimDirectionSuffix(string projectName)
        {
            if (string.IsNullOrWhiteSpace(projectName) || projectName.Length <= 1)
            {
                return projectName ?? string.Empty;
            }

            char suffix = projectName[projectName.Length - 1];
            return suffix == 'A' || suffix == 'B'
                ? projectName.Substring(0, projectName.Length - 1)
                : projectName;
        }

        private static void CopyJpegWithWatermark(string sourcePath, string targetPath, string watermarkText)
        {
            using (Image sourceImage = Image.FromFile(sourcePath))
            using (Bitmap bitmap = new Bitmap(sourceImage.Width, sourceImage.Height, PixelFormat.Format24bppRgb))
            {
                bitmap.SetResolution(sourceImage.HorizontalResolution, sourceImage.VerticalResolution);

                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.DrawImage(sourceImage, 0, 0, sourceImage.Width, sourceImage.Height);
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

                    float fontSize = Math.Max(18f, sourceImage.Width / 70f);
                    using (Font font = new Font("Microsoft YaHei", fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        float padding = fontSize * 0.5f;
                        using (Brush textBrush = new SolidBrush(Color.Red))
                        {
                            graphics.DrawString(
                                watermarkText,
                                font,
                                textBrush,
                                padding,
                                padding);
                        }
                    }
                }

                ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                    .FirstOrDefault(codec => codec.MimeType == "image/jpeg");

                if (jpegCodec == null)
                {
                    bitmap.Save(targetPath, ImageFormat.Jpeg);
                    return;
                }

                using (EncoderParameters encoderParams = new EncoderParameters(1))
                {
                    encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);
                    bitmap.Save(targetPath, jpegCodec, encoderParams);
                }
            }
        }
    }

    public enum National2026ImageKind
    {
        RoadImage,
        FrontImage
    }

    public sealed class National2026PictureExportTask
    {
        public National2026PictureExportTask(
            DirectoryInfo targetDirectory,
            List<PicAndMile> pictures,
            National2026ImageKind imageKind)
        {
            TargetDirectory = targetDirectory;
            Pictures = pictures ?? new List<PicAndMile>();
            ImageKind = imageKind;
        }

        public DirectoryInfo TargetDirectory { get; }

        public List<PicAndMile> Pictures { get; }

        public National2026ImageKind ImageKind { get; }

        public int Count => Pictures.Count;
    }
}
