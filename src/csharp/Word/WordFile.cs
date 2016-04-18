namespace Office.Word
{
    using System;
    using System.IO;

    public class WordFile : IDisposable
    {
        public string OriginalPath { get; set; }

        public string FilePath { get; private set; }

        public bool DeleteOriginal { get; set; }

        public WordFile(string originalPath, string path, bool deleteOriginal)
        {
            this.OriginalPath = originalPath;
            this.FilePath = path;
            this.DeleteOriginal = deleteOriginal;
        }

        public WordFile(string originalPath, string path)
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
            if (File.Exists(targetPath))
            {
                File.Delete(targetPath);
            }

            File.Copy(this.FilePath, targetPath);
            this.Dispose();
        }

        public void Rollback()
        {
            this.Dispose();
        }

        public static WordFile Create(string originalDocPath, bool deleteOriginal)
        {
            var tempPath = Path.GetTempFileName();
            var docxTemp = Path.ChangeExtension(tempPath, "docx");

            if (tempPath != docxTemp)
            {
                File.Delete(tempPath);
            }

            File.Copy(originalDocPath, docxTemp, true);

            return new WordFile(originalDocPath, docxTemp, deleteOriginal);
        }

        public static WordFile Create(Stream stream, string path)
        {
            using (var fs = File.Open(path, FileMode.OpenOrCreate))
            {
                var buffer = new byte[4096];
                var read = stream.Read(buffer, 0, buffer.Length);

                while (read > 0)
                {
                    fs.Write(buffer, 0, 4096);
                    read = stream.Read(buffer, 0, buffer.Length);
                }
            }

            return Create(path, true);
        }
    }
}