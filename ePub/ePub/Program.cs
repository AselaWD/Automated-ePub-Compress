using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Ionic.Zip;
using Ionic.Zlib;
using System.Reflection;
using System.Diagnostics;

namespace ePub
{
    class Program
    {
        static void Main(string[] args)
        {


            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(Current_Assembly_Resolve);


            int errCount = 0;
            string userName = Environment.UserName;
            string cleanPath = @"F:\eBookWorkflow\ePUB_Clean\";
            string SourcePath = @"Z:\Driver_Tools\Z\eBookWorkflow\ePUB_Clean\lib";
            string xsdSourceFile = "AppSettings.xsd";
            string xsdSourcePath = @"Z:\Driver_Tools\Z\eBookWorkflow\ePUB_Clean\" + xsdSourceFile;
            string exeSourcePath = @"Z:\Driver_Tools\Z\eBookWorkflow\ePUB_Clean\ePub_Clean.exe";

            string path = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\";
            string path1 = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            //string zipPath = @"mimetype.zip";
            //string newFile = @"mimetype";

            Console.WriteLine("-----------------------------");
            Console.WriteLine("Automated ePub Archiver 2.0v");
            Console.WriteLine("-----------------------------");

            string root = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Name;
            string[] files = Directory.GetFiles(path, "*.epub", SearchOption.TopDirectoryOnly);
            string[] unwantedExtensions = { }; // you can extend it  


            FileStream stream = null;

            try
            {
                stream = new FileStream(root + ".epub", FileMode.Open, FileAccess.Read, FileShare.None); //file is opened
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2147024894)
                {
                    //
                }
                else
                {
                    //ePub already Opened
                    errCount++;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[ERROR]: " + root + ".epub is already opened in another program.");
                    Console.ResetColor();
                }
                
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            if (errCount == 0) {
                using (var zip = new ZipFile())
                {
                    //Staus
                    DateTime st = DateTime.Now;
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("[STATUS]: Compression Started! [" + st + "]");
                    Console.ResetColor();

                    zip.EmitTimesInWindowsFormatWhenSaving = false;
                    if (File.Exists("mimetype"))
                        zip.AddFile("mimetype", "").CompressionLevel = CompressionLevel.None;
                    else
                    {
                        errCount++;
                        //META-INF not found!
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("[ERROR]: mimetype is not found!");
                        Console.ResetColor();
                    }

                    zip.CompressionLevel = CompressionLevel.BestCompression;


                    if (Directory.Exists(@"OEBPS"))
                        zip.AddDirectory(@"OEBPS", "OEBPS");
                    else if (Directory.Exists(@"OPS"))
                        zip.AddDirectory(@"OPS", "OPS");
                    else if (Directory.Exists(@"Ops"))
                        zip.AddDirectory(@"Ops", "Ops");
                    else
                    {
                        errCount++;

                        //OEBPS Not found
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("[ERROR]: OEBPS or OPS directory is not found!");
                        Console.ResetColor();

                        //MessageBox.Show("OEBPS or OPS directory is not found!", "Folder Not Found!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }


                    if (Directory.Exists(@"META-INF"))
                        zip.AddDirectory(@"META-INF", "META-INF");
                    else
                    {
                        errCount++;
                        //META-INF not found!
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("[ERROR]: META-INF directory is not found!");
                        Console.ResetColor();

                        //MessageBox.Show("META-INF directory is not found!", "Folder Not Found!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    //Save ePub
                    zip.Save(path + root + ".epubARC");
                    DateTime et = DateTime.Now;
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("[STATUS]: Compression Ended! [" + et + "]");

                    Console.WriteLine("[STATUS]: Total Error Count: [" + errCount + "]");

                    Console.ResetColor();

                    if (errCount > 0)
                    {

                        //End Console
                        Console.WriteLine("------------ END ------------");
                        //Message Error
                        System.IO.File.Delete(path + root + ".epubARC");
                        MessageBox.Show("Due to reported error(s), ePub is not generated.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Environment.Exit(0);
                    }
                    else
                    {
                        try
                        {
                            System.IO.File.Move(path + root + ".epubARC", path + root + ".epub");
                        }
                        catch (Exception ex)
                        {
                            if (ex.HResult == -2147024713)
                            {
                                try
                                {
                                    System.IO.File.Delete(path + root + ".epub");
                                }
                                catch (Exception ex1)
                                {
                                    if (ex1.HResult == -2147024864)
                                    {
                                        //ePub already Opened
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("[ERROR]: " + root + ".epub is already opened in another program.");
                                        Console.ResetColor();
                                    }
                                }

                                try
                                {
                                    System.IO.File.Move(path + root + ".epubARC", path + root + ".epub");
                                }
                                catch (Exception ex2)
                                {
                                    if (ex2.HResult == -2147024713)
                                    {

                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("[WARNING]: " + root + ".epubARC is already exists.");
                                        Console.ResetColor();

                                        System.IO.File.Delete(path + root + ".epubARC");
                                    }
                                }

                            }

                        }


                        //Message Completed
                        MessageBox.Show(root + ".epub generated successfully.", "Completed!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Blue;
                    Console.Write("[QUESTION]: Are you sure want to validate this ePub? [y/n]: ");
                    Console.ResetColor();


                    ConsoleKeyInfo cki = Console.ReadKey();

                    if (cki.Key.ToString() == "y" || cki.Key.ToString() == "Y" || cki.Key == ConsoleKey.Enter)
                    {
                        DateTime vt = DateTime.Now;
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[STATUS]: Validation: [YES]");
                        Console.WriteLine("[STATUS]: Validation Started! [" + vt + "]");
                        Console.ResetColor();

                        try
                        {

                            System.IO.File.Copy(path + root + ".epub", cleanPath + userName + "\\" + root + ".epub", true);
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("[STATUS]: Copied " + root + ".epub" + " to Clean Directory.");
                            Console.ResetColor();

                        }
                        catch (Exception ex)
                        {
                            if (ex.HResult == -2147024893)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(@"[ERROR]: F:\ partition not found.");
                                Console.ResetColor();
                            }

                            if (ex.HResult == -2147024893)
                            {
                                try
                                {
                                    Directory.CreateDirectory(cleanPath + userName);
                                    System.IO.File.Copy(path + root + ".epub", cleanPath + userName + "\\" + root + ".epub", true);
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("[STATUS]: Copied " + root + ".epub" + " to Clean Directory.");
                                    Console.ResetColor();
                                }
                                catch (Exception ex1)
                                {
                                    if (ex1.HResult == -2147024893)
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine(@"[ERROR]: F:\ partition not found.");
                                        Console.ResetColor();
                                    }
                                }
                            }
                        }

                        try
                        {
                            DirectoryCopy(SourcePath, cleanPath + "\\lib", true);
                        }
                        catch (Exception ex)
                        {
                            if (ex.HResult == -2147024893)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(@"[ERROR]: " + cleanPath + "\\lib" + " not found.");
                                Console.ResetColor();
                            }
                        }

                        try
                        {
                            System.IO.File.Copy(xsdSourcePath, cleanPath + xsdSourceFile, true);
                        }
                        catch (Exception ex)
                        {
                            if (ex.HResult == -2147024893)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(@"[ERROR]: " + cleanPath + xsdSourceFile + " not found.");
                                Console.ResetColor();
                            }
                        }

                        try
                        {
                            System.IO.File.Copy(exeSourcePath, cleanPath + "ePub_Clean.exe", true);
                        }
                        catch (Exception ex)
                        {
                            if (ex.HResult == -2147024893)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(@"[ERROR]: " + cleanPath + "ePub_Clean.exe" + " not found.");
                                Console.ResetColor();
                            }
                        }

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[STATUS]: Connected to the ePub Clean server.");
                        Console.Write("[STATUS]: Waiting for ePub_Clean.exe.");
                        Console.Write(".");
                        Console.Write(".");
                        Console.Write(".");
                        Console.WriteLine();
                        Console.ResetColor();

                        try
                        {
                            Process p = new Process();
                            p.StartInfo.FileName = cleanPath + "ePub_Clean.exe";
                            p.StartInfo.WorkingDirectory = cleanPath;
                            p.Start();

                            if (p.HasExited)
                            {
                                if (File.Exists(cleanPath + userName + @"\ePUB_Cleanup_Processing.html"))
                                {
                                    Process.Start(cleanPath + userName + @"\ePUB_Cleanup_Processing.html");
                                }
                               
                                if (File.Exists(cleanPath + userName + @"\ePUB_Cleanup_Processing.log"))
                                {
                                    Process.Start("notepad.exe", cleanPath + userName + @"\ePUB_Cleanup_Processing.log");
                                }

                             
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex.HResult == -2147467259)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(@"[ERROR]: " + cleanPath + "ePub_Clean.exe" + " not found.");
                                Console.ResetColor();
                            }
                        }

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[STATUS]: ePub_Clean.exe is running.");
                        Console.ResetColor();


                        //End Console
                        Console.WriteLine("------------ END ------------");

                        //MessageBox.Show("[STATUS]: Validation: [YES]", "Completed!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("[STATUS]: Validation: [NO]");
                        Console.ResetColor();

                        //End Console
                        Console.WriteLine("------------ END ------------");

                        //MessageBox.Show("[STATUS]: Validation: [NO]", "Completed!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }


                }
            }
            else
            {
                //End Console
                Console.WriteLine("------------ END ------------");
                //Message Error
                System.IO.File.Delete(path + root + ".epubARC");
                MessageBox.Show("Due to reported error(s), ePub is not generated.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }

        }

        static Assembly Current_Assembly_Resolve(object sender, ResolveEventArgs args)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("EmbedAssembly.IonicZip.dll"))
            {
                byte[] assemblyData = new byte[stream.Length];
                stream.Read(assemblyData, 0, assemblyData.Length);
                return Assembly.Load(assemblyData);
            }
        }
        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                try
                {
                    Directory.CreateDirectory(destDirName);
                }
                catch(Exception ex)
                {
                    if(ex.HResult == -2147024893)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(@"[ERROR]: "+ destDirName + " not found.");
                        Console.ResetColor();
                    }
                }
                
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);

                try
                {
                    file.CopyTo(temppath,true);   
                }
                catch(Exception ex)
                {
                    if(ex.HResult == -2147024893)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(@"[ERROR]: " + temppath + " not found.");
                        Console.ResetColor();
                    }
                    
                }

            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}
