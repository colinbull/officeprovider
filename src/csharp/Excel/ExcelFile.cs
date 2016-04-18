namespace Office.Excel
{
    using System;
    using System.IO;
    using System.Threading;

    public class ExcelFile : IDisposable
    {
        public string OriginalPath { get; set; }

        public string FilePath { get; private set; }

        public bool DeleteOriginal { get; set; }

        public ExcelFile(string originalPath, string path, bool deleteOriginal)
        {
            this.OriginalPath = originalPath;
            this.FilePath = path;
            this.DeleteOriginal = deleteOriginal;
        }

        public ExcelFile(string originalPath, string path)
        {
            this.OriginalPath = originalPath;
            this.FilePath = path;
            this.DeleteOriginal = false;
        }

        public void Dispose()
        {
            File.Delete(this.FilePath);

            if (DeleteOriginal)
            {
                File.Delete(OriginalPath);
            }
        }

        public void Commit(string targetPath)
        {
            File.Copy(this.FilePath, targetPath, true);
            this.Dispose();
        }

        public void Rollback()
        {
            this.Dispose();
        }

        public static ExcelFile Create(string originalDocPath, bool deleteOriginal)
        {
            var tempPath = Path.GetTempFileName();
            var xlsxTemp = Path.ChangeExtension(tempPath, "xlsx");

            if (tempPath != xlsxTemp)
            {
                File.Delete(tempPath);
            }

            File.Copy(originalDocPath, xlsxTemp, true);

            return new ExcelFile(originalDocPath, xlsxTemp, deleteOriginal);
        }

        public static ExcelFile Create(Stream stream, string path)
        {
            using (var fs = File.Open(path, FileMode.OpenOrCreate))
            {
                var buffer = new byte[4096];
                var read = stream.Read(buffer, 0, buffer.Length);

                while (read > 0)
                {
                    fs.Write(buffer, 0, read);
                    read = stream.Read(buffer, 0, buffer.Length);
                }
            }

            return Create(path, false);
        }
    }
}