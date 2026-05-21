using HNRoadFormatConverter.Entitys;
using HNRoadFormatConverter.MyEntitys;
using System;
using System.Collections.Generic;
using System.IO;

namespace HNRoadFormatConverter.Exporters
{
    /// <summary>
    /// 各规范共用的图片复制与分包逻辑。
    /// 这里只放“如何复制图片、如何按批次建目录”的通用行为；
    /// 具体规范的命名规则仍由调用方在 PicAndMile.ResultPicName 中提前准备好。
    /// </summary>
    public static class PictureExportService
    {
        public const int DefaultBatchSize = 5000;

        public static List<List<PicAndMile>> SplitByBatchSize(
            List<PicAndMile> pictures,
            int batchSize = DefaultBatchSize)
        {
            List<List<PicAndMile>> batches = new List<List<PicAndMile>>();
            if (pictures == null || pictures.Count == 0)
            {
                return batches;
            }

            if (batchSize <= 0)
            {
                batchSize = DefaultBatchSize;
            }

            for (int start = 0; start < pictures.Count; start += batchSize)
            {
                int count = Math.Min(batchSize, pictures.Count - start);
                batches.Add(pictures.GetRange(start, count));
            }

            return batches;
        }

        public static void ExportStandardBatches(
            DirectoryInfo targetRoot,
            List<List<PicAndMile>> batches,
            CityModelItem standard,
            string suffix,
            IProgress<int> progress)
        {
            if (targetRoot == null || batches == null)
            {
                return;
            }

            for (int batchIndex = 0; batchIndex < batches.Count; batchIndex++)
            {
                string directoryName = GetStandardBatchDirectoryName(standard, batchIndex);
                ExportBatch(targetRoot, batches[batchIndex], directoryName, suffix, progress);
            }
        }

        public static void ExportHunanBatches(
            DirectoryInfo targetRoot,
            List<List<PicAndMile>> batches)
        {
            if (targetRoot == null || batches == null)
            {
                return;
            }

            for (int batchIndex = 0; batchIndex < batches.Count; batchIndex++)
            {
                string targetDirectory = Path.Combine(targetRoot.FullName, "0", batchIndex.ToString());
                Directory.CreateDirectory(targetDirectory);

                foreach (PicAndMile picture in batches[batchIndex])
                {
                    CopyPicture(picture.PicPath, Path.Combine(targetDirectory, picture.ResultPicName + ".jpg"));
                }
            }
        }

        private static string GetStandardBatchDirectoryName(CityModelItem standard, int batchIndex)
        {
            switch (standard)
            {
                case CityModelItem.湖南省单位一定制:
                    return batchIndex.ToString("0");
                default:
                    return (batchIndex + 1).ToString("00");
            }
        }

        private static void ExportBatch(
            DirectoryInfo targetRoot,
            List<PicAndMile> pictures,
            string directoryName,
            string suffix,
            IProgress<int> progress)
        {
            string targetDirectory = Path.Combine(targetRoot.FullName, directoryName);
            Directory.CreateDirectory(targetDirectory);

            foreach (PicAndMile picture in pictures)
            {
                string targetPath = Path.Combine(targetDirectory, picture.ResultPicName + suffix);
                if (CopyPicture(picture.PicPath, targetPath))
                {
                    progress?.Report(1);
                }
            }
        }

        private static bool CopyPicture(string sourcePath, string targetPath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            {
                return false;
            }

            File.Copy(sourcePath, targetPath, true);
            return true;
        }
    }
}
