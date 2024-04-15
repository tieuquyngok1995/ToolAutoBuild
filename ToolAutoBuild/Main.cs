using AutoRunBuild.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoRunBuild
{
    public partial class Main : Form
    {
        // Set time 
        private static System.Timers.Timer timer;
        // Path folder SVN
        private string pathSVN;
        // Path folder Src
        private string pathSrc;
        // List Item
        List<ItemModel> listItemResource = new List<ItemModel>();
        List<ItemModel> listMessageResources = new List<ItemModel>();

        /// <summary>
        /// Main
        /// </summary>
        public Main()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Event load form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Main_Load(object sender, EventArgs e)
        {
            try
            {
                // get path svn folder in setting
                pathSVN = Properties.Settings.Default.pathSVN;
                // get path src folder in setting
                pathSrc = Properties.Settings.Default.pathSrc;

                autoRun();

                txtPathSVN.Text = !string.IsNullOrEmpty(pathSVN) ? pathSVN : string.Empty;
                txtPathSrc.Text = !string.IsNullOrEmpty(pathSrc) ? pathSrc : string.Empty;

                btnRun.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error during processing function [Main Load].\r\nError detail: " + ex.Message, "Error Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Event click open dialog select folder SVN
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenFolderSVN_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (!string.IsNullOrEmpty(pathSVN)) fbd.SelectedPath = pathSVN;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPathSVN.Text = fbd.SelectedPath;
                    pathSVN = fbd.SelectedPath;

                    Properties.Settings.Default.pathSVN = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error during processing function [Main Load].\r\nError detail: " + ex.Message, "Error Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Event click open dialog select folder src
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenFolderSrc_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (!string.IsNullOrEmpty(pathSrc)) fbd.SelectedPath = pathSrc;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPathSrc.Text = fbd.SelectedPath;
                    pathSrc = fbd.SelectedPath;

                    Properties.Settings.Default.pathSrc = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error during processing function [Main Load].\r\nError detail: " + ex.Message, "Error Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Event click button run
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                btnRun.Enabled = false;
                progressBar.Value = 0;

                // Clear logs
                txtLogs.Text = string.Empty;

                // Run command update svn
                if (runCommandUpdateSVN())
                {
                    txtLogs.Text += "Update SVN and Git successfully.";
                    progressBar.Value = 14;
                }
                else
                {
                    txtLogs.Text += "Update SVN and Git failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                if (readFileExcel(0))
                {
                    txtLogs.Text += "\r\nRead file Excel ItemResource successfully.";
                    progressBar.Value = 28;
                }
                else
                {
                    txtLogs.Text += "\r\nRead file Excel ItemResource failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                if (readAndEditFile(0))
                {
                    txtLogs.Text += "\r\nEdit file ItemResource successfully.";
                    progressBar.Value = 42;
                }
                else
                {
                    txtLogs.Text += "\r\nEdit file ItemResource failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                if (readFileExcel(1))
                {
                    txtLogs.Text += "\r\nRead file Excel MessageResources successfully.";
                    progressBar.Value = 56;
                }
                else
                {
                    txtLogs.Text += "\r\nRead file Excel MessageResources failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                if (readAndEditFile(1))
                {
                    txtLogs.Text += "\r\nEdit file MessageResources successfully.";
                    progressBar.Value = 70;
                }
                else
                {
                    txtLogs.Text += "\r\nEdit file MessageResources failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                if (runCommandBuildViewModel())
                {
                    txtLogs.Text += "\r\nRun command build ViewModel successfully.";
                    progressBar.Value = 84;
                }
                else
                {
                    txtLogs.Text += "\r\nRun command build ViewModel failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                if (runCommandCommitGit())
                {
                    txtLogs.Text += "\r\nCommit file to Git successfully.";
                    progressBar.Value = 100;
                }
                else
                {
                    txtLogs.Text += "\r\nCommit file to Git failed.";
                    progressBar.Value = 0;
                    btnRun.Enabled = true;
                    return;
                }

                txtLogs.Text += "\r\nThe processing was completed at " + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + ".";
                btnRun.Enabled = true;
            }
            catch (Exception ex)
            {
                txtLogs.Text += "\r\nDetailed error: " + ex.Message;
            }
        }

        #region Function
        /// <summary>
        /// Auto run build 
        /// </summary>
        private void autoRun()
        {
            timer = new System.Timers.Timer(); // Create a new timer instance
            timer.Elapsed += new ElapsedEventHandler(btnRun_Click); // Hook up the Elapsed event for the timer.
            timer.AutoReset = true; // Instruct the timer to restart every time the Elapsed event has been called                
            timer.SynchronizingObject = this; // Synchronize the timer with this form UI (IMPORTANT)
            timer.Interval = 3600000; // Set the interval to 1 second (1000 milliseconds)
            timer.Enabled = true; // Start the timer
        }

        /// <summary>
        /// Run command line update svn 
        /// </summary>
        /// <returns></returns>
        private bool runCommandUpdateSVN()
        {
            try
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = "/C cd /d \"" + pathSrc + "\"&& git checkout AutoBuild&& git pull&& svn update \"" + pathSVN + "\"&& exit";
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                process.Close();
                return true;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Read file exel 
        /// </summary>
        /// <returns></returns>
        private bool readFileExcel(short mode)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            try
            {
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;
                listItemResource = new List<ItemModel>();
                listMessageResources = new List<ItemModel>();
                string pathFileExcel = pathSVN;
                pathFileExcel += mode == 0 ? "/項目リソース.xlsx" : "/資料_メッセージ一覧_MessageResources.xlsx";

                if (!File.Exists(pathFileExcel)) return false;

                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(pathFileExcel);
                xlWorkbook.RefreshAll();
                Excel._Worksheet xlWorksheet = mode == 0 ? xlWorkbook.Sheets[2] : xlWorkbook.Sheets[4];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rows = xlRange.Rows.Count;

                //specify the range 
                if (mode == 0)
                {
                    for (int i = 2; i <= rows; i++)
                    {
                        if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value != null && xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value != null)
                        {
                            listItemResource.Add(new ItemModel(xlRange.Cells[i, 2].Value.ToString(), xlRange.Cells[i, 3].Value.ToString()));
                        }
                    }
                }
                else
                {
                    for (int i = 3; i <= rows; i++)
                    {
                        if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value != null && xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value != null)
                        {
                            string comment = xlRange.Cells[i, 4].Value == null ? string.Empty : xlRange.Cells[i, 4].Value.ToString();
                            listMessageResources.Add(new ItemModel(xlRange.Cells[i, 2].Value.ToString(), xlRange.Cells[i, 3].Value.ToString(), comment));
                        }
                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close(0);
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
                return true;
            }
            catch (Exception ex)
            {
                if (xlWorkbook != null) xlWorkbook.Close(0);
                if (xlApp != null) xlApp.Quit();
                throw ex;
            }
        }

        /// <summary>
        /// Read and Edit file
        /// </summary>
        /// <returns></returns>
        private bool readAndEditFile(short mode)
        {
            try
            {
                List<ItemModel> listTmpItemResource = new List<ItemModel>(listItemResource);
                List<ItemModel> listTmpMessageResources = new List<ItemModel>(listMessageResources);

                string pathFileServer = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources";
                pathFileServer += mode == 0 ? "/ItemResources.resx" : "/MessageResources.resx";

                string pathFileDesignerServer = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources";
                pathFileDesignerServer += mode == 0 ? "/ItemResources.Designer.cs" : "/MessageResources.designer.cs";

                string pathFileClient = pathSrc + "/src/server/angularapp/projects/app-generated/view-models/src/lib/resources";
                pathFileClient += mode == 0 ? "/item-resources.service.ts" : "/message-resources.service.ts";

                if (!File.Exists(pathFileServer)) return false;

                using (StreamReader file = new StreamReader(pathFileServer))
                {
                    string line;
                    while ((line = file.ReadLine()) != null)
                    {
                        if (line.Contains("<data name=\""))
                        {
                            string[] lines = line.Split('"');
                            if (mode == 0)
                            {
                                listTmpItemResource.RemoveAll(obj => obj.Key.Equals(lines[1].Trim()));
                            }
                            else
                            {
                                listTmpMessageResources.RemoveAll(obj => obj.Key.Equals(lines[1].Trim()));
                            }
                        }
                    }

                    file.Close();
                }

                // Edit file Item Resource
                if (listTmpItemResource.Count > 0)
                {
                    // Edit file server
                    string text = removeLastLineBlank(createTextFileItemResourcesServer(listTmpItemResource));
                    if (!string.IsNullOrEmpty(text))
                    {
                        string readText = string.Empty;
                        using (StreamReader streamReader = new StreamReader(pathFileServer, Encoding.GetEncoding("shift-jis")))
                        {
                            readText = streamReader.ReadToEnd();
                        }
                        readText = readText.Replace("</root>", text + "\r\n</root>");
                        File.WriteAllText(pathFileServer, readText);
                    }

                    // Edit file desgin server
                    text = removeLastLineBlank(createTextItemResourcesDesignerServer(listTmpItemResource));
                    if (!string.IsNullOrEmpty(text))
                    {
                        string readTextDesigner = string.Empty;
                        using (StreamReader streamReader = new StreamReader(pathFileDesignerServer, Encoding.GetEncoding("shift-jis")))
                        {
                            readTextDesigner = streamReader.ReadToEnd();
                        }
                        readTextDesigner = readTextDesigner.Remove(readTextDesigner.Length - 6) + text + "\r\n    }\r\n}\r\n";
                        File.WriteAllText(pathFileDesignerServer, readTextDesigner);
                    }

                    // Edit file client
                    text = removeLastLineBlank(createTextFileItemResourcesClient(listTmpItemResource));
                    if (!string.IsNullOrEmpty(text))
                    {
                        string readTextClient = string.Empty;
                        using (StreamReader streamReader = new StreamReader(pathFileClient, Encoding.GetEncoding("shift-jis")))
                        {
                            readTextClient = streamReader.ReadToEnd();
                        }
                        readTextClient = readTextClient.Remove(readTextClient.Length - 6) + text + "\r\n\r\n}";
                        File.WriteAllText(pathFileClient, readTextClient);
                    }
                }

                // Edit file Message Resource
                if (listTmpMessageResources.Count > 0)
                {
                    // Edit file server
                    string text = removeLastLineBlank(createTextFileMessageResourcesServer(listTmpMessageResources));

                    if (!string.IsNullOrEmpty(text))
                    {
                        string readText = string.Empty;
                        using (StreamReader streamReader = new StreamReader(pathFileServer, Encoding.GetEncoding("shift-jis")))
                        {
                            readText = streamReader.ReadToEnd();
                        }
                        readText = readText.Replace("</root>", text + "\r\n</root>");
                        File.WriteAllText(pathFileServer, readText);
                    }

                    // Edit file desgin server
                    text = removeLastLineBlank(createTextFileMessageResourcesDesignerServer(listTmpMessageResources));
                    if (!string.IsNullOrEmpty(text))
                    {
                        string readTextDesigner = string.Empty;
                        using (StreamReader streamReader = new StreamReader(pathFileDesignerServer, Encoding.GetEncoding("shift-jis")))
                        {
                            readTextDesigner = streamReader.ReadToEnd();
                        }
                        readTextDesigner = readTextDesigner.Remove(readTextDesigner.Length - 6) + text + "\r\n    }\r\n}\r\n";
                        File.WriteAllText(pathFileDesignerServer, readTextDesigner);
                    }

                    // Edit file client
                    text = removeLastLineBlank(createTextFileMessageResourcesClient(listTmpMessageResources));
                    if (!string.IsNullOrEmpty(text))
                    {
                        string readTextClient = string.Empty;
                        using (StreamReader streamReader = new StreamReader(pathFileClient, Encoding.GetEncoding("shift-jis")))
                        {
                            readTextClient = streamReader.ReadToEnd();
                        }
                        readTextClient = readTextClient.Remove(readTextClient.Length - 6) + text + "\r\n\r\n}";
                        File.WriteAllText(pathFileClient, readTextClient);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Run command build viewmodel
        /// </summary>
        /// <returns></returns>
        private bool runCommandBuildViewModel()
        {
            try
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = "/C cd /d \"" + pathSrc + "/src/server/angularapp\"&& ng build @app-generated/view-models&& exit";
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();

                return true;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Run command commit git
        /// </summary>
        /// <returns></returns>
        private bool runCommandCommitGit()
        {
            try
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = "/C cd /d \"" + pathSrc + "\"&& git commit -a -m \"Auto edit and commit file ItemResources and MessageResources\"&& git push && exit";
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                process.Close();
                return true;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Ceate data file ItemResources in server
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createTextFileItemResourcesServer(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "  <data name=\"{0}\" xml:space=\"preserve\">\r\n    <value>{1}</value>\r\n    <comment>{1}</comment>\r\n  </data>\r\n";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Key, item.Value);
            }
            return result;
        }

        /// <summary>
        /// Ceate data file ItemResources Designer in server
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createTextItemResourcesDesignerServer(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "\r\n        /// <summary>\r\n        ///   Looks up a localized string similar to {0}.\r\n        /// </summary>\r\n        public static string {1} {{\r\n            get {{\r\n                return ResourceManager.GetString(\"{1}\", resourceCulture);\r\n            }}\r\n        }}";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Value, item.Key);
            }
            return result;
        }

        /// <summary>
        /// Ceate data file MessageResources in server
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createTextFileMessageResourcesServer(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "  <data name=\"{0}\" xml:space=\"preserve\">\r\n    <value>{1}</value>\r\n    <comment>{2}</comment>\r\n  </data>\r\n";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Key, item.Value.Replace("\n", "\r\n"), item.Comment);
            }
            return result;
        }

        /// <summary>
        /// Ceate data file MessageResources Designer in server
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createTextFileMessageResourcesDesignerServer(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "\r\n        /// <summary>\r\n        ///   Looks up a localized string similar to {0}.\r\n        /// </summary>\r\n        public static string {1} {{\r\n            get {{\r\n                return ResourceManager.GetString(\"{1}\", resourceCulture);\r\n            }}\r\n        }}";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Value.Replace("\n", "\r\n"), item.Key);
            }
            return result;
        }

        /// <summary>
        /// Ceate data file ItemResources in client
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createTextFileItemResourcesClient(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "\r\n  /** {0} */\r\n  readonly {1} = `{0}`;\r\n";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Value, item.Key);
            }
            return result;
        }

        /// <summary>
        /// Ceate data file MessageResources in client
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createTextFileMessageResourcesClient(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "\r\n  /** {0} */\r\n  readonly {1} = `{2}`;\r\n";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Comment, item.Key, item.Value.Replace("\n", "\r\n"));
            }
            return result;
        }

        /// <summary>
        /// Remove new line in text
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string removeLastLineBlank(string str)
        {
            int lastIndex = str.LastIndexOf("\r\n");
            if (lastIndex != -1 && lastIndex == str.Length - 2)
            {
                str = str.Substring(0, lastIndex);
                str = removeLastLineBlank(str);
            }
            return str;
        }
        #endregion
    }
}