using System;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;
using System.Collections.Generic;
using System.Net.NetworkInformation;
using System.Net;

namespace AzureFileWatcher
{
    public class Program
    {
        #region Propriety
        //スリープ時間
        static string TIMETOCOPY = ConfigurationManager.AppSettings["TimeToCopy"].ToString();
        //Azureファイルのパス
        static string AZUREFILE = ConfigurationManager.AppSettings["AzureFile"];
        //送り先パス
        static string DESTINYFILE = ConfigurationManager.AppSettings["DestinyFile"];

        static string sourceDir;
        static StringBuilder strLogMessge = new StringBuilder();
        static DateTime dateTime;
        //Configで入力した待機時間の変換
        static int hour = Convert.ToInt32(TIMETOCOPY.Substring(0, 2));
        static int min = Convert.ToInt32(TIMETOCOPY.Substring(3, 2));
        static int seg = Convert.ToInt32(TIMETOCOPY.Substring(6, 2));
        static TimeSpan time = new TimeSpan(hour, min, seg);

        static private System.Timers.Timer timeWatcher;
        static Queue<string> queueFullPath = new Queue<string>();
        static Queue<string> queueNamePaht = new Queue<string>();
        static int countQueue = 0;

        #endregion

        static void Main(string[] args)
        {
            //多重起動の確認
            if (!DoubleProcess())
                return;
            //PAｔｈの確認
            if (!CheckPath())
                return;

            //監視ツールを立ち上げる
            FileSystemWatcher watcher = new FileSystemWatcher();
            //特定のファイルを監視するかフィルターをつける。　この場合「全部」
            watcher.Filter = "*.*";
            //監視するパス
            watcher.Path = AZUREFILE + "\\";
            //監視するもの。　これでトリガー発生が決まる
            watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            //トリガーを発生させる
            watcher.EnableRaisingEvents = true;
            //サブジレクトリーも監視するか
            watcher.IncludeSubdirectories = true;
            //イベントの作成
            watcher.Created += new FileSystemEventHandler(CreateFile);
            watcher.Deleted += new FileSystemEventHandler(DeleteFile);
            watcher.Renamed += new RenamedEventHandler(RenameFile);
            //変わるまで待機する
            watcher.WaitForChanged(WatcherChangeTypes.Renamed, 5000);

            //アプリの自動シャットダウンを防ぐ
            new AutoResetEvent(false).WaitOne();
        }

        public static void CreateFile(object sender, FileSystemEventArgs e)
        {
            try
            {
                //ログ
                WriteLog(">Copy File-", "CreateFile");

                //１００ファイルコピーしたらコンソール画面をクルーンアップする
                if (countQueue >= 100)
                {
                    //コンソール画面消去
                    Console.Clear();
                    //カウンターゼロ
                    countQueue = 0;
                }

                if (!string.IsNullOrEmpty(sourceDir))
                {
                    //フォルダーすべてを移動する/コピーする                  
                    DirectoryCopy(e.FullPath, Path.Combine(DESTINYFILE, e.Name), true);

                    //元のパスをきれいにする
                    sourceDir = string.Empty;
                }
                //ファイルの種類をゲット
                var extension = Path.GetExtension(e.Name);

                //種類がなかったらフォルダーである

                if (string.IsNullOrEmpty(extension))
                { //フォルダーの作成                        
                    Directory.CreateDirectory(Path.Combine(DESTINYFILE, e.Name));
                    string[] dir = Directory.GetFiles(e.FullPath);
                    if (dir.Length > 0)
                    {
                        foreach (string url in dir)
                        {
                            //キューにパスを入れる
                            queueFullPath.Enqueue(url);
                            //キューにパスの名前を入れる
                            queueNamePaht.Enqueue(Path.Combine(e.Name, Path.GetFileName(url)));
                            //タイマー開始
                            RunTimeWatcher();
                        }
                    }
                }
                else
                {
                    if (!Directory.Exists(Path.GetDirectoryName(Path.Combine(DESTINYFILE, e.Name))))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(Path.Combine(DESTINYFILE, e.Name)));
                    }
                    //キューにパスを入れる
                    queueFullPath.Enqueue(e.FullPath);
                    //キューにパスの名前を入れる
                    queueNamePaht.Enqueue(e.Name);
                    //タイマー開始
                    RunTimeWatcher();
                }
            }
            catch (Exception ex)
            {
                WriteLog("******\n\nERROR\n\n******", $"Error - DeleteFile {e.Name}", ex, true);
            }
        }

        public static void RunTimeWatcher()
        {
            timeWatcher = new System.Timers.Timer();

            timeWatcher.AutoReset = false;               // 1回しか呼ばない場合はfalse
            timeWatcher.Interval = time.TotalMilliseconds;                // Intervalの設定単位はミリ秒
            timeWatcher.Elapsed += timeWatcher_Elapsed;       // タイマイベント処理(時間経過後の処理)を登録
            timeWatcher.Enabled = true;                  // <-- これを呼ばないとタイマは開始しません

            //現在時刻
            dateTime = DateTime.Now;
            //ログ
            WriteLog($"---Copy file will start in..... {dateTime.AddTicks(time.Ticks).ToLongTimeString()}  Config Time / Interval : {TIMETOCOPY} ", "CreateFile");
        }
        private static void timeWatcher_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
          
            //パスをキューから取り出す
            var fullPath = queueFullPath.Dequeue();
            //ファイル名をキューから取り出す
            var namePath = queueNamePaht.Dequeue();
            try
            {   //ログ
                WriteLog($"---Copy START.......FileName : {namePath} ", "CreateFile");
                //ファイルのコピー開始　
                File.Copy(fullPath, Path.Combine(DESTINYFILE, namePath), true);
                WriteLog($"---Finalizing.......FileName : {namePath}", $"CreateFile {namePath}", null, true);
                //コンソール画面を消すためのカウンター
                countQueue++;

                if (queueNamePaht.Count == 0)
                    Console.WriteLine("\n Wating File........");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //エラーのためパスをキューに戻す
                queueFullPath.Enqueue(fullPath);
                //エラーのためパス名をキューに戻す
                queueNamePaht.Enqueue(namePath);

                WriteLog("******\n\nERROR\n\n******", $"Error-CreateFile- {namePath}", ex, true);
                //接続の確認
                ReconnectAzureFile();
            }
        }

        public static void DeleteFile(object sender, FileSystemEventArgs e)
        {
            //現在時刻をGET
            dateTime = DateTime.Now;
            try
            {
                //ファイルの送り先
                string destFile = Path.Combine(DESTINYFILE, e.Name);
                //ファイルの種類取得
                FileAttributes attr = File.GetAttributes(Path.Combine(DESTINYFILE, e.Name));

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    //ディレクトリー取得
                    string[] diretorio = Directory.GetDirectories(destFile);
                    //ファイル取得
                    string[] sorce = Directory.GetFiles(destFile);
                    //ファイルが一個以上/ディレクトリ―が一個以上ある場合は一斉移動するメソッド
                    if (sorce.Length > 0 || diretorio.Length > 0)
                    {
                        sourceDir = destFile;
                        //削除
                        DeleteFile(destFile);
                    }
                    else
                    {
                        //削除
                        DeleteFile(destFile);
                    }
                }
                else
                {
                    //削除
                    DeleteFile(destFile);
                }
            }
            catch (Exception ex)
            {
                WriteLog("******\n\nERROR\n\n******", $"Error - DeleteFile {e.Name}", ex, true);
            }
        }

        public static void RenameFile(object sender, RenamedEventArgs e)
        {
            dateTime = DateTime.Now;
            // Thread.Sleep(TIMETOCOPY);            
            try
            {
                //種類取得
                FileAttributes attr = File.GetAttributes(Path.Combine(DESTINYFILE, e.OldName));
                //ディレクトリーの場合
                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {//ディレクトリーの名前変更
                    if (Directory.Exists(Path.Combine(DESTINYFILE, e.OldName)))
                    {
                        Directory.Move(Path.Combine(DESTINYFILE, e.OldName), Path.Combine(DESTINYFILE, e.Name));
                    }
                }
                else
                //ファイルの場合
                if (File.Exists(Path.Combine(DESTINYFILE, e.OldName)))
                {
                    //ファイルの名前変更
                    File.Move(Path.Combine(DESTINYFILE, e.OldName), Path.Combine(DESTINYFILE, e.Name));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                WriteLog("******\n\nERROR\n\n****** ", "RenameFle", ex, true);
                //接続の確認
                ReconnectAzureFile();
            }
        }

        static void DeleteFile(string destFile)
        {
            try
            {
                FileAttributes attr = File.GetAttributes(destFile);

                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {//ディレクトリー削除
                    if (Directory.Exists(destFile))
                        Directory.Delete(destFile, true);
                }//ファイル削除
                else if (File.Exists(destFile))
                    File.Delete(destFile);
            }
            catch { }
        }

        static bool CheckPath()
        {
            bool ok = true;
            Console.WriteLine("\n Checking Azure File Drive............");
            
            if(!PortInUse(445))
            {
                Console.WriteLine("\n Checking Azure File Drive............");
                ok = false;
            }

            try
            {
                int result = ConnectionAzureFile();
                if (result != NO_ERROR)
                {
                    Console.WriteLine("\n  Cant accesss the azure file Share. \n Check the config file and chek port 445 is open to access file.");
                    Console.WriteLine("\n Press Enter to Exit.");
                    Console.ReadKey();
                  
                    return false;
                }


                if (!Directory.Exists(DESTINYFILE))
                {
                    WriteLog("Destiny File Path is INVALID.", "CreateFile");
                    WriteLog("Creating File.... ", "CreateFile", null, true);
                    Directory.CreateDirectory(DESTINYFILE);
                }
            }
            catch (Exception ex)
            {
                WriteLog("Error!! Check the Config File.", "CheckPath", ex, true);

                ok = false;
            }

            if (!ok)
            {
                Console.WriteLine("\n Press Enter to Exit.");
                Console.ReadKey();
            }
            else
            {
                Console.Clear();
                Console.WriteLine("\n Wating File........");
            }
            return ok;
        }

        public static bool PortInUse(int port)
        {
            bool inUse = false;

            IPGlobalProperties ipProperties = IPGlobalProperties.GetIPGlobalProperties();
            IPEndPoint[] ipEndPoints = ipProperties.GetActiveTcpListeners();


            foreach (IPEndPoint endPoint in ipEndPoints)
            {
                if (endPoint.Port == port)
                {
                    inUse = true;
                    break;
                }
            }


            return inUse;
        }

        static private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(sourceDirName);
                DirectoryInfo[] dirs = dir.GetDirectories();

                // If the source directory does not exist, throw an exception.
                if (!dir.Exists)
                {
                    throw new DirectoryNotFoundException(
                        "Source directory does not exist or could not be found: "
                        + sourceDirName);
                }

                // If the destination directory does not exist, create it.
                if (!Directory.Exists(destDirName))
                {
                    Directory.CreateDirectory(destDirName);
                }


                // Get the file contents of the directory to copy.
                FileInfo[] files = dir.GetFiles();

                foreach (FileInfo file in files)
                {
                    // Create the path to the new copy of the file.
                    string temppath = Path.Combine(destDirName, file.Name);

                    // Copy the file.
                    file.CopyTo(temppath, false);
                }

                // If copySubDirs is true, copy the subdirectories.
                if (copySubDirs)
                {

                    foreach (DirectoryInfo subdir in dirs)
                    {
                        // Create the subdirectory.
                        string temppath = Path.Combine(destDirName, subdir.Name);

                        // Copy the subdirectories.
                        DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                    }
                }
            }
            catch
            {

            }
        }

        #region console

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;


        private static void WriteLog(string msg, string method, Exception ex = null, bool writeToLog = false)
        {
            dateTime = DateTime.Now;
            string message = $"[{dateTime.ToString("yyyy/MM/dd HH:mm:ss")}] {msg}";
            if (ex != null)
                message += $"\n\n{ex.ToString()}";

            Console.WriteLine($"{message}");
            strLogMessge.Append($"\n\n {message}");

            if (writeToLog)
            {
                WriteTraceLog(strLogMessge.ToString(), method, ex);
                strLogMessge.Clear();
            }
        }
        public static void ShowConsoleWindow()
        {
            var handle = GetConsoleWindow();

            if (handle == IntPtr.Zero)
            {
                AllocConsole();
            }
            else
            {
                ShowWindow(handle, SW_SHOW);
            }
        }

        public static void HideConsoleWindow()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
        }
        #endregion

        #region  Verificate process
        private static bool DoubleProcess()
        {
            bool ok = true;
            //Mutex名を決める（必ずアプリケーション固有の文字列に変更すること！）
            string mutexName = "AzureFileWatch";

            bool hasHandle = false;
            //Mutexオブジェクトを作成する
            System.Threading.Mutex mutex = new System.Threading.Mutex(true, mutexName, out hasHandle);

            try
            {

                //ミューテックスを得られたか調べる
                if (hasHandle == false)
                {
                    //得られなかった場合は、すでに起動していると判断して終了
                    Console.WriteLine("多重起動はできません。");
                    Console.ReadKey();
                    ok = false;
                }
            }
            finally
            {
                if (hasHandle)
                {
                    //ミューテックスを解放する
                    mutex.ReleaseMutex();
                }
                mutex.Close();
            }
            return ok;
        }
        #endregion

        #region LogFile

        /// <summary>
        /// ログ出力
        /// </summary>
        /// <param name="msg">メッセージ</param>
        /// <param name="ex">Exception(無指定の場合はメッセージのみ出力)</param>
        /// <remarks></remarks>
        /// 
        private static void WriteTraceLog(String msg, string Title, Exception ex)
        {
            try
            {
                // ログフォルダ名作成
                String logFolder = System.AppDomain.CurrentDomain.BaseDirectory + "Log";

                // ログフォルダ名作成
                if (!Directory.Exists(logFolder))
                    System.IO.Directory.CreateDirectory(logFolder);
                dateTime = DateTime.Now;
                var data = dateTime.ToString("yyyy-MM-dd-HH-mm-ss");
                // ログファイル名作成
                String logFile = logFolder + $"\\{data}---{Title.Replace("\\", "")}.log";

                System.IO.StreamWriter sw = null;
                try
                {
                    sw = new StreamWriter(logFile, true, System.Text.Encoding.Default); //.GetEncoding(932)));
                    sw.WriteLine(msg);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                finally
                {
                    if (sw != null)
                        sw.Close();
                }
            }
            catch { }
        }
        #endregion]

        #region Create Init Config -- Create RegistryKey and Shortcut File
        private static void CreateRegistryKey(string shortCut)
        {
            Microsoft.Win32.RegistryKey regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
            string strExecutableName = "";
            string strPath = shortCut;  // Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, strExecutableName);           
            regkey.SetValue(strExecutableName, strPath);
            regkey.Close();
            //https://k-sugi.sakura.ne.jp/it_synthesis/windows/visual-c/4418/
        }

        private static void StartupConfig()
        {

            string shortcutPath = System.IO.Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory), @"MyApp.lnk");
            // ショートカットのリンク先(起動するプログラムのパス)
            string targetPath = System.AppDomain.CurrentDomain.BaseDirectory;

            // WshShellを作成
            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
            // ショートカットのパスを指定して、WshShortcutを作成
            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutPath);
            // ①リンク先
            shortcut.TargetPath = targetPath;
            // ②引数
            shortcut.Arguments = "/a /b /c";
            // ③作業フォルダ
            shortcut.WorkingDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            // ④実行時の大きさ 1が通常、3が最大化、7が最小化
            shortcut.WindowStyle = 1;
            // ⑤コメント
            shortcut.Description = "テストのアプリケーション";
            // ⑥アイコンのパス 自分のEXEファイルのインデックス0のアイコン
            shortcut.IconLocation = System.AppDomain.CurrentDomain.BaseDirectory + ",0";

            // ショートカットを作成
            shortcut.Save();

            // 後始末
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shortcut);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shell);
        }
        #endregion

        #region ファイルの認証/接続

        //接続切断するWin32 API を宣言

        [DllImport("mpr.dll", EntryPoint = "WNetCancelConnection2", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]

        private static extern int WNetCancelConnection2(string lpName, Int32 dwFlags, bool fForce);

        //認証情報を使って接続するWin32 API宣言

        [DllImport("mpr.dll", EntryPoint = "WNetAddConnection2", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]

        private static extern int WNetAddConnection2(ref NETRESOURCE lpNetResource, string lpPassword, string lpUsername, Int32 dwFlags);

        //WNetAddConnection2に渡す接続の詳細情報の構造体

        [StructLayout(LayoutKind.Sequential)]
        internal struct NETRESOURCE
        {
            public int dwScope;//列挙の範囲
            public int dwType;//リソースタイプ
            public int dwDisplayType;//表示オブジェクト
            public int dwUsage;//リソースの使用方法

            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpLocalName;//ローカルデバイス名。使わないならNULL。

            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpRemoteName;//リモートネットワーク名。使わないならNULL

            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpComment;//ネットワーク内の提供者に提供された文字列

            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpProvider;//リソースを所有しているプロバイダ名

          
        }

        #region Errors

        private const int NO_ERROR = 0;

        private const int ERROR_ACCESS_DENIED = 5;
        private const int ERROR_ALREADY_ASSIGNED = 85;
        private const int ERROR_BAD_DEVICE = 1200;
        private const int ERROR_BAD_NET_NAME = 67;
        private const int ERROR_BAD_PROVIDER = 1204;
        private const int ERROR_CANCELLED = 1223;
        private const int ERROR_EXTENDED_ERROR = 1208;
        private const int ERROR_INVALID_ADDRESS = 487;
        private const int ERROR_INVALID_PARAMETER = 87;
        private const int ERROR_INVALID_PASSWORD = 1216;
        private const int ERROR_MORE_DATA = 234;
        private const int ERROR_NO_MORE_ITEMS = 259;
        private const int ERROR_NO_NET_OR_BAD_PATH = 1203;
        private const int ERROR_NO_NETWORK = 1222;

        private const int ERROR_BAD_PROFILE = 1206;
        private const int ERROR_CANNOT_OPEN_PROFILE = 1205;
        private const int ERROR_DEVICE_IN_USE = 2404;
        private const int ERROR_NOT_CONNECTED = 2250;
        private const int ERROR_OPEN_FILES = 2401;
        private const int ERRO_BAD_NETPATH = 53;

        #endregion
        //認証に失敗すると UnauthorizedAccessException が発生する(Exceptionでキャッチしてもいいかも)
        private static int ConnectionAzureFile()
        {
            NETRESOURCE netResource = new NETRESOURCE();

            netResource.dwScope = 0; //列挙の範囲
            netResource.dwType = 1;  //リソースタイプ
            netResource.dwDisplayType = 0;//表示オブジェクト
            netResource.dwUsage = 0;//リソースの使用方法
            // ネットワークドライブにする場合は"z:"などドライブレター設定  
            netResource.lpLocalName = ConfigurationManager.AppSettings["DriveRemoteName"];//ローカルデバイス名。使わないならNULL。
            //share name
            netResource.lpRemoteName = ConfigurationManager.AppSettings["AzureServer"];//リモートネットワーク名。使わないならNULL
            netResource.lpProvider = "";
            netResource.lpComment ="";//ネットワーク内の提供者に提供された文字列
              //パスワード
        string password = ConfigurationManager.AppSettings["Password"];
            //ユーザーID
            string userId = ConfigurationManager.AppSettings["UserID"];
            int ret = 0;

            try
            {
                //既に接続してる場合があるので一旦切断する
                ret = WNetCancelConnection2(netResource.lpLocalName, 0, true);  //netResource.lpRemoteName, 0, true);

                //if (ret != NO_ERROR)
                // return ret;

                //共有フォルダに認証情報を使って接続
                ret = WNetAddConnection2(ref netResource, password, userId, 0);

                return ret;
            }
            catch (Exception ex)
            {
                return ERROR_CANCELLED;
            }
        }

        private static void ReconnectAzureFile()
        {
            int count = 0;
            while (ConnectionAzureFile()!= NO_ERROR)
            {
                Console.WriteLine($"\n Tryng to connect azure share file.....{ ConfigurationManager.AppSettings["AzureServer"]}  Try count : {count}");
                Thread.SpinWait(10000);
                count++;
           }
        }

        public static bool IsDriveMapped()
        {
            string[] DriveList = Environment.GetLogicalDrives();
            for (int i = 0; i < DriveList.Length; i++)
            {
                var test = ConfigurationManager.AppSettings["DriveRemoteName"];
                if (ConfigurationManager.AppSettings["DriveRemoteName"] + "\\" == DriveList[i].ToString())
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }
}
