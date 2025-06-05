using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileOperations
{
    internal class Program
    {
        static void Main(string[] args)
        {
            FileActivates.Instance.createFile("test.txt");
            FileActivates.Instance.createDirectory("test");
            FileActivates.Instance.CopyDirectory("test", "test_1t");
            FileActivates.Instance.CopyFile("test.txt", "test_1t.txt");
            string strfile = FileActivates.Instance.GetFileName("test.txt");
            string getFileWithoutExtension = FileActivates.Instance.GetFileNameWithoutExtension("test.txt");
            string strExtension = FileActivates.Instance.GetFileExtension("test.txt");
            bool blFileExist = FileActivates.Instance.FileExists("test.txt");


        }
    }
}
