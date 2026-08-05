using System.IO;

namespace PrintLayoutAddin.Core
{
    public class TitleBlockInfo
    {
        public string Path { get; set; }
    }

    public static class TitleBlockReader
    {
        public static TitleBlockInfo Load(string dwgPath)
        {
            if (string.IsNullOrEmpty(dwgPath) || !File.Exists(dwgPath))
                throw new FileNotFoundException("Không tìm thấy file khung tên: " + dwgPath);
            return new TitleBlockInfo { Path = dwgPath };
        }
    }
}
