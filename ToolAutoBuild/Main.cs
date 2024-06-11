using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using System.Windows.Forms;
using ToolAutoBuild.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace ToolAutoBuild
{
    public partial class Main : Form
    {
        // Set time 
        private static System.Timers.Timer timer;
        // Branh name
        private readonly string branchName = "Develop";
        // Path folder SVN
        private string pathSVN;
        // Path folder Src
        private string pathSrc;
        // Total Data
        private int totalItemResource = 0;
        private int totalMessageResources = 0;
        // check edit   
        private bool isEditItemResource = false;
        private bool isEditMessageResources = false;
        // List Item
        List<ItemModel> listItemResourceInExcel = new List<ItemModel>();
        List<ItemModel> listItemResourceInFile = new List<ItemModel>();
        List<ItemModel> listMessageResourcesInExcel = new List<ItemModel>();
        List<ItemModel> listMessageResourcesFile = new List<ItemModel>();

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
                isEditItemResource = false;
                isEditMessageResources = false;

                btnRun.Enabled = false;
                progressBar.Value = 0;

                // Clear logs
                txtLogs.Text = string.Empty;

                // Clear excel process
                System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (System.Diagnostics.Process p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        p.Kill();
                    }
                }

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

                if (!isEditItemResource && !isEditMessageResources)
                {
                    txtLogs.Text += "\r\nThere is no change, ending the processing at " + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + ".";
                    progressBar.Value = 100;
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

                // clear data
                listItemResourceInExcel.Clear();
                listItemResourceInFile.Clear();
                listMessageResourcesInExcel.Clear();
                listMessageResourcesFile.Clear();
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
                //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.Arguments = "/C cd /d \"" + pathSrc + "\"&& git checkout " + branchName + "&& git reset --hard && git pull&& svn update \"" + pathSVN + "\"&& exit";
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
                listItemResourceInExcel = new List<ItemModel>();
                listMessageResourcesInExcel = new List<ItemModel>();
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
                            listItemResourceInExcel.Add(new ItemModel(xlRange.Cells[i, 2].Value.ToString(), xlRange.Cells[i, 3].Value.ToString()));
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
                            listMessageResourcesInExcel.Add(new ItemModel(xlRange.Cells[i, 2].Value.ToString(), xlRange.Cells[i, 3].Value.ToString(), comment));
                        }
                    }
                }

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close(SaveChanges: false);
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

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
                bool isEdit = false;
                string line = string.Empty;
                string key = string.Empty;
                string value = string.Empty;

                listItemResourceInFile = new List<ItemModel>();
                listMessageResourcesFile = new List<ItemModel>();

                List<ItemModel> listTmpItemResource = new List<ItemModel>(listItemResourceInExcel);
                List<ItemModel> listTmpMessageResources = new List<ItemModel>(listMessageResourcesInExcel);

                string pathFileServer = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources";
                pathFileServer += mode == 0 ? "/ItemResources.resx" : "/MessageResources.resx";

                string pathFileDesignerServer = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources";
                pathFileDesignerServer += mode == 0 ? "/ItemResources.Designer.cs" : "/MessageResources.designer.cs";

                string pathFileClient = pathSrc + "/src/server/angularapp/projects/app-generated/view-models/src/lib/resources";
                pathFileClient += mode == 0 ? "/item-resources.service.ts" : "/message-resources.service.ts";

                if (!File.Exists(pathFileServer)) return false;

                using (StreamReader file = new StreamReader(pathFileServer))
                {
                    while ((line = file.ReadLine()) != null)
                    {
                        if (line.Trim().Equals("-->")) isEdit = true;
                        if (!isEdit) continue;

                        string[] lines = line.Split('"');
                        if (line.Contains("<data name=\""))
                        {
                            key = lines.Length > 0 ? lines[1] : string.Empty;
                        }
                        else if (line.Contains("<value>") && !string.IsNullOrEmpty(key))
                        {
                            value += line.Replace("    <value>", "").Replace("</value>", "");
                        }
                        else if (line.Contains("</value>") && !string.IsNullOrEmpty(key))
                        {
                            string val = line.Replace("    <value>", "").Replace("</value>", "");
                            value += string.IsNullOrEmpty(val) ? string.Empty : "\r\n" + val;
                        }
                        else if ((line.Contains("<comment>") || line.Contains("</data>")) && !string.IsNullOrEmpty(key))
                        {
                            string comment = line.Replace("<comment>", "").Replace("</comment>", "").Replace("</data>", "").Trim();
                            comment = string.IsNullOrEmpty(comment) ? value : comment;
                            if (mode == 0)
                            {
                                ItemModel itemModel = listTmpItemResource.FirstOrDefault(obj => obj.Key.Equals(key));
                                if (itemModel != null)
                                {
                                    value = itemModel.Value;
                                }
                                listItemResourceInFile.Add(new ItemModel(key, value, comment == "NULL" ? value : comment));
                                listTmpItemResource.RemoveAll(obj => obj.Key.Equals(key) && obj.Value.Equals(value));
                            }
                            else
                            {
                                ItemModel itemModel = listTmpMessageResources.FirstOrDefault(obj => obj.Key.Equals(key));
                                if (itemModel != null)
                                {
                                    value = itemModel.Value;
                                }

                                int index = value.IndexOf("　　※");
                                string val = value.Contains("<br>") ? value.Replace("<br>", "\r\n") : value;
                                if (index != -1)
                                {
                                    val = val.Remove(index);
                                }

                                listMessageResourcesFile.Add(new ItemModel(key, val, comment));
                                listTmpMessageResources.RemoveAll(obj => obj.Key.Equals(key) && obj.Value.Equals(value));
                            }

                            key = string.Empty;
                            value = string.Empty;
                        }
                        else if (!line.Contains("</value>") && !line.Contains("<comment>") && !string.IsNullOrEmpty(key))
                        {
                            if (line.Contains("<value />"))
                            {
                                value = string.Empty;
                            }
                            else
                            {
                                value += "\r\n" + line;
                            }
                        }
                    }
                    file.Close();
                }

                string dataFile = string.Empty;
                string readFile = string.Empty;
                Encoding utf8WithoutBom = new UTF8Encoding(true);
                // Edit file Item Resource
                if (mode == 0 && listTmpItemResource.Count > 0)
                {
                    foreach (ItemModel itemModel in listTmpItemResource)
                    {
                        string comment = string.IsNullOrEmpty(itemModel.Comment) ? itemModel.Value : itemModel.Comment;
                        listItemResourceInFile.Add(new ItemModel(itemModel.Key.Replace(" ", ""), itemModel.Value, comment.Replace("\r\n", "").Replace("\n", "")));
                    }
                }

                // Check and edit file Item Resource
                if (mode == 0 && (totalItemResource == 0 || listItemResourceInFile.Count != totalItemResource))
                {
                    // Edit data in file .resx
                    dataFile = removeLastLineBlank(createDataFileRESXServer(listItemResourceInFile));
                    if (!string.IsNullOrEmpty(dataFile))
                    {
                        readFile = string.Join("\r\n", File.ReadLines(pathFileServer, utf8WithoutBom).Take(119));
                        dataFile = readFile + "\r\n" + dataFile + "\r\n</root>";
                        File.WriteAllText(pathFileServer, dataFile, utf8WithoutBom);
                    }

                    // Edit data in file Designer.resx
                    dataFile = removeLastLineBlank(createDataFileDesignerServer(listItemResourceInFile));
                    if (!string.IsNullOrEmpty(dataFile))
                    {
                        readFile = string.Join("\r\n", File.ReadLines(pathFileDesignerServer, utf8WithoutBom).Take(61));
                        dataFile = readFile + dataFile + "\r\n    }\r\n}\r\n";
                        File.WriteAllText(pathFileDesignerServer, dataFile, utf8WithoutBom);
                    }

                    // Edit data in file client
                    dataFile = removeLastLineBlank(createDataFileClient(listItemResourceInFile));
                    if (!string.IsNullOrEmpty(dataFile))
                    {
                        readFile = string.Join("\r\n", File.ReadLines(pathFileClient, utf8WithoutBom).Take(19));
                        dataFile = readFile + dataFile + "\r\n\r\n}";
                        File.WriteAllText(pathFileClient, dataFile, utf8WithoutBom);
                    }

                    totalItemResource = listItemResourceInFile.Count;
                    isEditItemResource = true;
                }

                // Edit file Message Resource
                if (mode == 1 && listTmpMessageResources.Count > 0)
                {
                    foreach (ItemModel itemModel in listTmpMessageResources)
                    {
                        string comment = string.IsNullOrEmpty(itemModel.Comment) ? itemModel.Value : itemModel.Comment;
                        listMessageResourcesFile.Add(new ItemModel(itemModel.Key, itemModel.Value, comment.Replace("\r\n", "").Replace("\n", "")));
                    }
                }

                if (mode == 1 && (totalMessageResources == 0 || listMessageResourcesFile.Count != totalMessageResources))
                {
                    // Edit data in file .resx
                    dataFile = removeLastLineBlank(createDataFileRESXServer(listMessageResourcesFile));
                    if (!string.IsNullOrEmpty(dataFile))
                    {
                        readFile = string.Join("\r\n", File.ReadLines(pathFileServer, utf8WithoutBom).Take(119));
                        dataFile = readFile + "\r\n" + dataFile + "\r\n</root>";
                        File.WriteAllText(pathFileServer, dataFile, utf8WithoutBom);
                    }

                    // Edit data in file Designer.resx
                    dataFile = removeLastLineBlank(createDataFileDesignerServer(listMessageResourcesFile));
                    if (!string.IsNullOrEmpty(dataFile))
                    {
                        readFile = string.Join("\r\n", File.ReadLines(pathFileDesignerServer, utf8WithoutBom).Take(62));
                        dataFile = readFile + dataFile + "\r\n    }\r\n}\r\n";
                        File.WriteAllText(pathFileDesignerServer, dataFile, utf8WithoutBom);
                    }

                    // Edit data in file client
                    dataFile = removeLastLineBlank(createDataFileClient(listMessageResourcesFile));
                    if (!string.IsNullOrEmpty(dataFile))
                    {
                        readFile = string.Join("\r\n", File.ReadLines(pathFileClient, utf8WithoutBom).Take(19));
                        dataFile = readFile + dataFile + "\r\n\r\n}";
                        File.WriteAllText(pathFileClient, dataFile, utf8WithoutBom);
                    }

                    totalMessageResources = listMessageResourcesFile.Count;
                    isEditMessageResources = true;
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
                //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
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
                //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
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
        /// Ceate data file .resx in server
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createDataFileRESXServer(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "  <data name=\"{0}\" xml:space=\"preserve\">\r\n    {1}\r\n    <comment>{2}</comment>\r\n  </data>\r\n";
            foreach (ItemModel item in listData)
            {
                string value = string.Empty;
                if (string.IsNullOrEmpty(item.Value))
                {
                    value = "<value />";
                }
                else
                {
                    value = "<value>" + item.Value.Replace("\r\n", "\n").Replace("\n", "\r\n") + "</value>";
                }
                result += string.Format(tmp, item.Key, value, item.Comment.Replace("\r\n", "").Replace("\n", ""));
            }
            return result;
        }

        /// <summary>
        /// Ceate data file .Designer.cs in server
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createDataFileDesignerServer(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "\r\n        /// <summary>\r\n        ///   Looks up a localized string similar to {0}.\r\n        /// </summary>\r\n        public static string {1} {{\r\n            get {{\r\n                return ResourceManager.GetString(\"{1}\", resourceCulture);\r\n            }}\r\n        }}\r\n";
            listData = listData.OrderBy(x => x.Key).ToList();
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Value.Replace("\r\n", "\n").Replace("\n", "\r\n        ///"), item.Key);
            }
            return result;
        }

        /// <summary>
        /// Ceate data file in client
        /// </summary>
        /// <param name="listData"></param>
        /// <returns></returns>
        private string createDataFileClient(List<ItemModel> listData)
        {
            string result = string.Empty;
            string tmp = "\r\n  /** {0} */\r\n  readonly {1} = `{2}`;\r\n";
            foreach (ItemModel item in listData)
            {
                result += string.Format(tmp, item.Comment.Replace("\r\n", "").Replace("\n", ""), item.Key, item.Value.Replace("NULL", "").Replace("\r\n", "\n").Replace("\n", "\r\n"));
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