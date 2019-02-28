using System;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Threading;

namespace HttpDataServerProject1
{
    class Program
    {
        private static Boolean errorParsing = false;
        private static Boolean doKill = false;
        private static Boolean doCopy = false;
        private static Boolean doStart = false;
        private static Int32 port = 80;
        private static String password = String.Empty;
        static void Main(String[] args)
        {
            foreach (String arg in args)
            {
                if (arg.Length < 2)
                {
                    errorParsing = true;
                }
                else
                {
                    switch (arg.Substring(0, 2))
                    {
                        case "-k":
                        case "/k":
                            doKill = true;
                            break;
                        case "-c":
                        case "/c":
                            doCopy = true;
                            break;
                        case "-s":
                        case "/s":
                            doStart = true;
                            break;
                        case "-p":
                        case "/p":
                            if (!Int32.TryParse(arg.Substring(2, arg.Length - 2), out port))
                            {
                                errorParsing = true;
                            }
                            break;
                        case "-w":
                        case "/w":
                            if (arg.Length > 2)
                            {
                                password = arg.Substring(2);
                            }
                            break;
                        default:
                            errorParsing = true;
                            break;
                    }
                }
            }

            if (errorParsing)
            {
                Console.WriteLine(@"usage: programName[ (-|/)k)?( (-|/)c)?( (-|/)s)?( (-|/)p\d+)?");
            }
            else
            {
                if (doKill)
                {
                    KillMyProcesses();
                    Thread.Sleep(500);
                }

                if (doCopy)
                {
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService 1c7");
                    //CopyDeployDirToExecDir(@"C:\Lnetpub\DataService 1c on port 11014");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService Sql on port 11012");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService Utilities on port 11009");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\Sssp"); // 11008
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService Mail on port 11007");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService Fs on port 11005");
                    //CopyDeployDirToExecDir(@"C:\Lnetpub\DataService 1c on port 11004");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService Excel on port 11003");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\DataService Sql on port 11002");
                    CopyDeployDirToExecDir(@"C:\Lnetpub\Proxy on port 80");
                }

                if (doStart)
                {
                    StartMyProcesses(args);
                }
            }
            Console.ReadKey();
        }
        private static void KillMyProcesses()
        {
            //ConsoleWriteProcesses();

            KillProcesses("MsServer");
            KillProcesses("HttpDataServerProject2");
            KillProcesses("HttpDataServerProject3");
            //KillProcesses("HttpDataServerProject4");
            KillProcesses("HttpDataServerProject5");
            KillProcesses("HttpDataServerProject7");
            KillProcesses("HttpDataServerProject8");
            KillProcesses("HttpDataServices"); // Project9
            KillProcesses("Project12_SqlServer");
            //KillProcesses("HttpDataServerProject14_1c7");
            KillProcesses("Project_1c7");
            KillProcesses("1cv7s");
            KillProcesses("Excel");
            KillProcesses("Word");

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            //ConsoleWriteProcesses();
        }
        private static void KillProcesses(String name)
        {
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps)
            {
                if (p.ProcessName.ToUpper() == name.ToUpper())
                {
                    p.Kill();
                    Console.WriteLine("Process {0,5:#####} {1:s} has been killed.", p.Id, p.ProcessName);
                }
            }
        }
        private static void CopyDeployDirToExecDir(String deployPath)
        {
            String dPath = Path.Combine(deployPath);
            DirectoryInfo ddi = new DirectoryInfo(dPath);
            String ePath = Path.Combine(ddi.Parent.FullName, ddi.Name + ".exec");
            //Console.WriteLine(ePath);
            DirectoryInfo edi = new DirectoryInfo(ePath);
            if (!edi.Exists)
            {
                edi = Directory.CreateDirectory(ePath);
            }
            else
            {
                //if (ddi.LastWriteTime > edi.LastWriteTime)
                {
                    edi.Delete(true);
                    edi = Directory.CreateDirectory(ePath);
                }
            }
            DirectoryInfo[] temp = ddi.GetDirectories("Application Files");
            if (temp.Length > 0)
            {
                DirectoryInfo sdi = temp[0];
                temp = sdi.GetDirectories();
                if (temp.Length > 0)
                {
                    // ищем последнюю директорию
                    sdi = null;
                    String fullName = "";
                    foreach (DirectoryInfo di in temp)
                    {
                        if (String.CompareOrdinal(fullName, di.FullName) < 0)
                        {
                            sdi = di;
                            fullName = di.FullName;
                        }
                    }
                    Console.WriteLine(sdi.FullName);
                    // все файлы копируем в edi
                    DirectoryCopyWithRemoveDeploySufix(sdi.FullName, edi.FullName, true);
                }
            }
        }
        private static void DirectoryCopyWithRemoveDeploySufix(string sourceDirName, string destDirName, bool copySubDirs)
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
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                String fileName = file.Name;
                if (fileName.EndsWith(".deploy"))
                {
                    fileName = fileName.Substring(0, fileName.Length - 7);
                }
                string temppath = Path.Combine(destDirName, fileName);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopyWithRemoveDeploySufix(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
        private static void StartMyProcesses(String[] args)
        {
            System.Diagnostics.Process pi = new System.Diagnostics.Process();
            pi.StartInfo.Domain = "sibdomain.ru";
            pi.StartInfo.UserName = "sokolov";
            pi.StartInfo.Password = ReadPassword(password);
            pi.StartInfo.UseShellExecute = false;
            String StartInfoArguments = String.Join(" ", args);

            Console.WriteLine(pi.StartInfo.Arguments);

            pi.StartInfo.Arguments = @"1cd=""\\SRV-TS2\dbase_1c$\Фармацея Фарм-Сиб"" 1cn=""Соколов_Евгенй_клиент_2"" 1cp=""yNFxfrvqxP"" port=11004";
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService 1c7.exec\Project_1c7.exe";
            pi.Start();
            Thread.Sleep(100);

            pi.StartInfo.Arguments = @"1cd=""\\SRV-TS2\dbase_1c$\ФК_Гарза"" 1cn=""Соколов_Евгенй_клиент_2"" 1cp=""yNFxfrvqxP"" port=11014";
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService 1c7.exec\Project_1c7.exe";
            pi.Start();
            Thread.Sleep(100);

            pi.StartInfo.Arguments = StartInfoArguments;

            //pi.StartInfo.FileName = @"C:\Lnetpub\DataService 1c on port 11014.exec\HttpDataServerProject14_1c7.exe";
            //pi.Start();
            //Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService Sql on port 11012.exec\Project12_SqlServer.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService Utilities on port 11009.exec\HttpDataServices.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\Sssp.exec\HttpDataServerProject8.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService Mail on port 11007.exec\HttpDataServerProject7.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService Fs on port 11005.exec\HttpDataServerProject5.exe";
            pi.Start();
            Thread.Sleep(100);
            //pi.StartInfo.FileName = @"C:\Lnetpub\DataService 1c on port 11004.exec\HttpDataServerProject4.exe";
            //pi.Start();
            //Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService Excel on port 11003.exec\HttpDataServerProject3.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\DataService Sql on port 11002.exec\HttpDataServerProject2.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.FileName = @"C:\Lnetpub\Proxy on port 80.exec\MsServer.exe";
            pi.Start();
            Thread.Sleep(100);
            pi.StartInfo.Arguments += " -p81";
            pi.StartInfo.FileName = @"C:\Lnetpub\Proxy on port 80.exec\MsServer.exe";
            pi.Start();
            Thread.Sleep(100);
        }
        private static void ConsoleWriteProcesses()
        {
            Process[] ps = Process.GetProcesses();
            //Console.WriteLine("Begin -----------------------------------------------------------");
            foreach (Process p in ps)
            {
                Console.WriteLine("{0,5:#####} {1:s}", p.Id, p.ProcessName);
            }
            //Console.WriteLine("End   -----------------------------------------------------------");
        }
        public static SecureString ReadPassword(string password)
        {
            SecureString secPass = new SecureString();
            for (int i = 0; i < password.Length; i++)
                secPass.AppendChar(password[i]);
            return secPass;
        }
    }
}
