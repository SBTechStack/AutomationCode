using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileOperations
{
    public class FileActivates : IFile, IDirectory
    {

        private static FileActivates instance = null;

        public static FileActivates Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new FileActivates();
                }
                return instance;
            }
        }

        private string DefaultRootPath = Environment.CurrentDirectory;
        public FileActivates() { }

        public bool IgnoreExistingFiles
        {
            get;
            set;
        }
        public bool OverwriteExistingFiles
        {
            get;
            set;
        }

        public string Filename { get; set; }
        public string FilePath { get; set; }
        public string RootPath { get; set; }
        public string Directoryname { get; set; }
        public string DirectoryPath { get; set; }

        public void ActivateFile(string rootPath)
        {
            DefaultRootPath = rootPath;
        }
        public void createFile(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty.");

            FilePath = System.IO.Path.Combine(DefaultRootPath, filename);
            if (!System.IO.File.Exists(DefaultRootPath))
                System.IO.File.Create(FilePath);
            else
                throw new InvalidOperationException("File already exists at the specified path.");
        }
        public void createDirectory(string directoryname)
        {
            if (string.IsNullOrEmpty(directoryname))
                throw new ArgumentException("Directory name cannot be null or empty.");

            DirectoryPath = System.IO.Path.Combine(DefaultRootPath, directoryname);
            if (!System.IO.Directory.Exists(DirectoryPath))
                System.IO.Directory.CreateDirectory(DirectoryPath);
            else
                throw new InvalidOperationException("Directory already exists at the specified path.");
        }
        public void deleteFile(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty.");

            FilePath = System.IO.Path.Combine(DefaultRootPath, filename);
            if (System.IO.File.Exists(FilePath))
                System.IO.File.Delete(FilePath);
            else
                throw new FileNotFoundException("File not found at the specified path.");
        }
        public void deleteDirectory(string directoryname)
        {
            if (string.IsNullOrEmpty(directoryname))
                throw new ArgumentException("Directory name cannot be null or empty.");

            DirectoryPath = System.IO.Path.Combine(DefaultRootPath, directoryname);
            if (System.IO.Directory.Exists(DirectoryPath))
                System.IO.Directory.Delete(DirectoryPath, true);
            else
                throw new DirectoryNotFoundException("Directory not found at the specified path.");
        }
        public void moveFile(string sourceFilename, string destinationFilename)
        {
            if (string.IsNullOrEmpty(sourceFilename) || string.IsNullOrEmpty(destinationFilename))
                throw new ArgumentException("Source and destination filenames cannot be null or empty.");

            string sourcePath = System.IO.Path.Combine(DefaultRootPath, sourceFilename);
            string destinationPath = System.IO.Path.Combine(DefaultRootPath, destinationFilename);

            if (System.IO.File.Exists(sourcePath))
                System.IO.File.Move(sourcePath, destinationPath);
            else
                throw new FileNotFoundException("Source file not found at the specified path.");
        }
        public void moveDirectory(string sourceDirectoryname, string destinationDirectoryname)
        {
            if (string.IsNullOrEmpty(sourceDirectoryname) || string.IsNullOrEmpty(destinationDirectoryname))
                throw new ArgumentException("Source and destination directory names cannot be null or empty.");

            string sourcePath = System.IO.Path.Combine(DefaultRootPath, sourceDirectoryname);
            string destinationPath = System.IO.Path.Combine(DefaultRootPath, destinationDirectoryname);

            if (System.IO.Directory.Exists(sourcePath))
                System.IO.Directory.Move(sourcePath, destinationPath);
            else
                throw new DirectoryNotFoundException("Source directory not found at the specified path.");
        }
        public void ActivateDirectory(string rootPath)
        {
            DefaultRootPath = rootPath;
        }
        public string GetFilePath(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty.");

            return System.IO.Path.Combine(DefaultRootPath, filename);
        }
        public string GetDirectoryPath(string directoryname)
        {
            if (string.IsNullOrEmpty(directoryname))
                throw new ArgumentException("Directory name cannot be null or empty.");

            return System.IO.Path.Combine(DefaultRootPath, directoryname);
        }
        public string GetFileExtension(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty.");

            string extension = System.IO.Path.GetExtension(filename);
            if (!string.IsNullOrEmpty(extension))
                Console.WriteLine($"File extension: {extension}");
            else
                Console.WriteLine("No file extension found.");

            return extension;

        }
        public string GetFileNameWithoutExtension(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty.");

            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(filename);
            if (!string.IsNullOrEmpty(fileNameWithoutExtension))
                Console.WriteLine($"File name without extension: {fileNameWithoutExtension}");
            else
                Console.WriteLine("No file name found without extension.");

            return fileNameWithoutExtension;
        }
        public string GetDirectoryName(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath))
                throw new ArgumentException("Directory path cannot be null or empty.");

            string directoryName = System.IO.Path.GetFileName(directoryPath);
            if (!string.IsNullOrEmpty(directoryName))
                Console.WriteLine($"Directory name: {directoryName}");
            else
                Console.WriteLine("No directory name found.");

            return directoryName;
        }
        public string GetFileName(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.");

            string fileName = System.IO.Path.GetFileName(filePath);
            if (!string.IsNullOrEmpty(fileName))
                Console.WriteLine($"File name: {fileName}");
            else
                Console.WriteLine("No file name found.");

            return fileName;
        }
        public string GetFileNameWithoutDirectory(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.");

            string fileNameWithoutDirectory = System.IO.Path.GetFileName(filePath);
            if (!string.IsNullOrEmpty(fileNameWithoutDirectory))
                Console.WriteLine($"File name without directory: {fileNameWithoutDirectory}");
            else
                Console.WriteLine("No file name found without directory.");

            return fileNameWithoutDirectory;
        }                
        public void CopyFile(string sourceFilename, string destinationFilename, bool overwrite = false)
        {
            if (string.IsNullOrEmpty(sourceFilename) || string.IsNullOrEmpty(destinationFilename))
                throw new ArgumentException("Source and destination filenames cannot be null or empty.");

            string sourcePath = System.IO.Path.Combine(DefaultRootPath, sourceFilename);
            string destinationPath = System.IO.Path.Combine(DefaultRootPath, destinationFilename);

            if (System.IO.File.Exists(sourcePath))
            {
                if (!IgnoreExistingFiles || !System.IO.File.Exists(destinationPath))
                    System.IO.File.Copy(sourcePath, destinationPath, overwrite);
            }
            else
            {
                throw new FileNotFoundException("Source file not found at the specified path.");
            }
        }
        public  void CopyDirectory  (string sourceDirectoryname, string destinationDirectoryname, bool recursive = false)
        {
            // Check if the target directory exists, if not, create it.
            if (System.IO.Directory.Exists(destinationDirectoryname) == false) System.IO.Directory.CreateDirectory(destinationDirectoryname);

            // Copy each file into it’s new directory.
            foreach (string fi in System.IO.Directory.GetFiles(sourceDirectoryname))
            {
                string targetFileName = System.IO.Path.Combine(destinationDirectoryname, System.IO.Path.GetFileName(fi));

                if (!IgnoreExistingFiles || !System.IO.File.Exists(targetFileName))
                    System.IO.File.Copy(fi, targetFileName, OverwriteExistingFiles);
            }

            if (recursive)
            {
                // Copy each subdirectory using recursion.
                foreach (string SubDir in System.IO.Directory.GetDirectories(sourceDirectoryname))
                    CopyDirectory(SubDir, System.IO.Path.Combine(destinationDirectoryname, System.IO.Path.GetFileName(SubDir)), recursive);
            }
        } 
        public bool FileExists(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty.");

            string filePath = System.IO.Path.Combine(DefaultRootPath, filename);
            return System.IO.File.Exists(filePath);
        }
    }


    public interface IFile
    {
        string Filename { get; set; }
        string FilePath { get; set; }
        string RootPath { get; set; }
    }
    public interface IDirectory
    {
        string Directoryname { get; set; }
        string DirectoryPath { get; set; }
    }
}
