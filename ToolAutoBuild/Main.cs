using Microsoft.CSharp;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Resources;
using System.Resources.Tools;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using ToolAutoBuild.Model;

namespace ToolAutoBuild
{
    public partial class Main : Form
    {
        // Path folder SVN
        private string pathSVN;
        // Path folder Src
        private string pathSrc;
        // Set time 
        private static Timer timeAutoRun;
        // Branh name
        private readonly string branchName = "Develop";

        /// <summary>
        /// Main
        /// </summary>
        public Main()
        {
            InitializeComponent();

            timeAutoRun = new Timer();
            timeAutoRun.Interval = 3600000; // 1 h = 3600000ms
            timeAutoRun.Tick += (sender, e) =>
            {
                btnRun_Click(sender, e);
            };

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
                btnRun.Enabled = false;

                // get path svn folder in setting
                pathSVN = Properties.Settings.Default.pathSVN;
                // get path src folder in setting
                pathSrc = Properties.Settings.Default.pathSrc;

                txtPathSVN.Text = !string.IsNullOrEmpty(pathSVN) ? pathSVN : string.Empty;
                txtPathSrc.Text = !string.IsNullOrEmpty(pathSrc) ? pathSrc : string.Empty;

                if (!string.IsNullOrEmpty(pathSVN) && !string.IsNullOrEmpty(pathSrc))
                {
                    btnRun.Enabled = true;
                    btnRun.Focus();
                }
                timeAutoRun.Start();
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
                if (!string.IsNullOrEmpty(pathSVN) && !string.IsNullOrEmpty(pathSrc)) btnRun.Enabled = true;
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
                if (!string.IsNullOrEmpty(pathSVN) && !string.IsNullOrEmpty(pathSrc)) btnRun.Enabled = true;
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
        private async void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;

                progressBar.Value = 0;

                // Clear logs
                txtLogs.Text = string.Empty;

                string command = string.Empty;
                if (CheckBranchDevelop())
                {
                    command = $"/C git reset --hard && git pull && svn update \"{pathSVN}/\" && exit";
                }
                else
                {
                    command = $"/C git stash && git checkout {branchName} && git reset --hard && git pull && svn update \"{pathSVN}/\" && exit";
                }

                RunCommandSrc(command, "Update Git and SVN");
                txtLogs.Text += "Update SVN and Git successfully.";
                UpdateProgress(10);

                await readEditFileItemResource();
                UpdateProgress(45);

                await readEditFileMessageResourceAsync();
                UpdateProgress(80);
                txtLogs.Text += "\r\n----------------------------------------------";

                command = $"/C ng build @app-generated/view-models && exit";
                RunCommandSrc(command, "Build view-model");
                txtLogs.Text += "\r\nRun command build ViewModel successfully.";
                UpdateProgress(90);

                command = $"/C git commit -a -m \"Auto edit and commit file ItemResources and MessageResources\" && git push && exit";
                //RunCommandSrc(command, "Commit git");
                txtLogs.Text += "\r\nRun command commit file to Git successfully.";
                UpdateProgress(100);
                txtLogs.Text += "\r\n----------------------------------------------";

                txtLogs.Text += "\r\nThe processing was completed at " + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + ".";

                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                txtLogs.Text += "\r\nDetailed error: " + ex.Message;
                progressBar.Value = 0;

                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
            }
        }

        /// <summary>
        /// Read and edit file item resource
        /// </summary>
        /// <returns></returns>
        private async Task<bool> readEditFileItemResource()
        {
            try
            {
                txtLogs.Text += "\r\n----------------------------------------------";

                string pathFileExcel = pathSVN + "/項目リソース.xlsx";
                if (!File.Exists(pathFileExcel)) throw new Exception("File 項目リソース does not exist");

                // Read data in file excel 
                Dictionary<string, ItemModel> excelList = CUtils.File.ReadExcelToList(pathFileExcel, 0);

                string pathFileResources = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources/ItemResources.resx";
                if (!File.Exists(pathFileExcel)) throw new Exception("File ItemResources does not exist");

                //  Read data in file resx
                List<ItemModel> resxList = CUtils.File.ReadResxData(pathFileResources);

                // Process data and create .resx file from processed dataz
                await GenerateFile(pathFileResources, excelList, resxList);
                txtLogs.Text += "\r\nEdit file ItemResources.resx successfully.";

                string pathFileDesigner = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources/ItemResources.Designer.cs";
                await GenerateDesignerFile(pathFileResources, pathFileDesigner, "ItemResources");
                txtLogs.Text += "\r\nEdit file ItemResources.Designer.cs successfully.";

                string pathFileService = pathSrc + "/src/server/angularapp/projects/app-generated/view-models/src/lib/resources/item-resources.service.ts";
                if (!File.Exists(pathFileService)) throw new Exception("File item-resources.service does not exist");
                txtLogs.Text += "\r\nEdit file item-resources.service.ts successfully.";

                await GenerateServiceFile(pathFileService, resxList);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Read and edit file message resource
        /// </summary>
        /// <returns></returns>
        private async Task<bool> readEditFileMessageResourceAsync()
        {
            try
            {
                txtLogs.Text += "\r\n----------------------------------------------";

                string pathFileExcel = pathSVN + "/資料_メッセージ一覧_MessageResources.xlsx";
                if (!File.Exists(pathFileExcel)) throw new Exception("File 資料_メッセージ一覧_MessageResources does not exist");

                // Read data in file excel 
                Dictionary<string, ItemModel> excelList = CUtils.File.ReadExcelToList(pathFileExcel, 1);

                string pathFileResources = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources/MessageResources.resx";
                if (!File.Exists(pathFileExcel)) throw new Exception("File MessageResources does not exist");

                //  Read data in file resx
                List<ItemModel> resxList = CUtils.File.ReadResxData(pathFileResources);

                // Process data and create .resx file from processed data
                await GenerateFile(pathFileResources, excelList, resxList);
                txtLogs.Text += "\r\nEdit file MessageResources.resx successfully.";

                string pathFileDesigner = pathSrc + "/src/server/ITS.UsoliaShogai.Lib/Resources/MessageResources.designer.cs";
                await GenerateDesignerFile(pathFileResources, pathFileDesigner, "MessageResources");
                txtLogs.Text += "\r\nEdit file MessageResources.designer.cs successfully.";

                string pathFileService = pathSrc + "/src/server/angularapp/projects/app-generated/view-models/src/lib/resources/message-resources.service.ts";
                if (!File.Exists(pathFileService)) throw new Exception("File message-resources.service does not exist");
                txtLogs.Text += "\r\nEdit file message-resources.service.ts successfully.";

                await GenerateServiceFile(pathFileService, resxList);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Process to create ItemResources.resx file with data content from excel file
        /// </summary>
        /// <param name="pathFile"></param>
        /// <param name="excelList"></param>
        /// <param name="resxList"></param>
        /// <returns></returns>
        private Task GenerateFile(string pathFile, Dictionary<string, ItemModel> excelList, List<ItemModel> resxList)
        {
            // Handle data
            foreach (var resxItem in resxList)
            {
                if (excelList.TryGetValue(resxItem.Key, out var excelItem))
                {
                    // Check and update value/comment
                    if (resxItem.Value != excelItem.Value)
                    {
                        resxItem.Value = excelItem.Value;
                    }
                    if (resxItem.Comment != excelItem.Comment)
                    {
                        resxItem.Comment = excelItem.Comment;
                    }
                    excelList.Remove(resxItem.Key);
                }
                else
                {
                    // Find value/comment
                    var excelFindItem = excelList.Values.FirstOrDefault(item => item.Value == resxItem.Value && item.Comment == resxItem.Comment);

                    if (excelFindItem != null)
                    {
                        // Update key and comment
                        resxItem.Key = excelFindItem.Key;
                        excelList.Remove(resxItem.Key);
                    }
                }
            }
            resxList.AddRange(excelList.Values);

            XDocument doc = XDocument.Load(pathFile);
            doc.Root.Elements("data").Remove();
            // Set data
            foreach (var item in resxList)
            {
                doc.Root.Add(new XElement("data",
                    new XAttribute("name", item.Key),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    new XElement("value", item.Value),
                    new XElement("comment", item.Comment)
                ));
            }
            // Save file
            doc.Save(pathFile);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Process to create ItemResources.Designer.cs file with data content from excel file
        /// </summary>
        /// <param name="resxPath"></param>
        /// <param name="designerPath"></param>
        /// <param name="className"></param>
        private Task GenerateDesignerFile(string resxPath, string designerPath, string className)
        {
            string namespaceName = "ITS.UsoliaShogai.Resources";
            if (File.Exists(designerPath)) File.Delete(designerPath);

            using (var resxReader = new ResXResourceReader(resxPath))
            {
                resxReader.BasePath = Path.GetDirectoryName(resxPath);

                var resourceEntries = new System.Collections.Hashtable();
                foreach (System.Collections.DictionaryEntry entry in resxReader)
                {
                    resourceEntries.Add(entry.Key, entry.Value);
                }

                using (var writer = new StreamWriter(designerPath))
                {
                    var provider = new CSharpCodeProvider();
                    var options = new CodeGeneratorOptions { BracingStyle = "C" };

                    // Create CodeCompileUnit
                    var compileUnit = StronglyTypedResourceBuilder.Create(
                        resourceEntries,
                        className,
                        namespaceName,
                        provider,
                        false,
                        out string[] unmatched
                    );

                    provider.GenerateCodeFromCompileUnit(compileUnit, writer, options);
                }
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Process to create item-resources.service.ts file with data content from excel file
        /// </summary>
        /// <param name="servicePath"></param>
        /// <param name="resxList"></param>
        private Task GenerateServiceFile(string servicePath, List<ItemModel> resxList)
        {
            List<string> lines = File.ReadAllLines(servicePath).ToList();

            if (lines.Count >= 19)
            {
                // Remove old content
                lines.RemoveRange(18, lines.Count - 18);

                // Create data
                List<string> newLines = new List<string>();
                foreach (ItemModel item in resxList)
                {
                    string comment = item.Comment.Replace("\r\n", "").Replace("\n", "");
                    string value = item.Value.Replace("NULL", "").Replace("\r\n", "\n").Replace("\n", "\r\n");
                    newLines.Add($"\r\n  /** {(string.IsNullOrEmpty(comment) ? item.Value.Replace("NULL", "").Replace("\r\n", "\n").Replace("\n", "") : comment)} */");
                    newLines.Add($"  readonly {item.Key} = `{value}`;");
                }

                lines.AddRange(newLines);
                lines.Add("}");
            }

            // Save file
            File.WriteAllLines(servicePath, lines);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Check branch is Develop
        /// </summary>
        /// <returns></returns>
        private bool CheckBranchDevelop()
        {
            string branch = string.Empty;
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c git rev-parse --abbrev-ref HEAD",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = pathSrc
                };

                using (Process process = new Process { StartInfo = startInfo })
                {
                    process.Start();
                    branch = process.StandardOutput.ReadToEnd().Trim();
                    process.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return branch.Equals(branchName);
        }

        /// <summary>
        /// Run Command
        /// </summary>
        /// <param name="command"></param>
        /// <param name="action"></param>
        private void RunCommandSrc(string command, string action)
        {
            try
            {
                string path = action.Contains("view-model") ? $"{pathSrc}/src/server/angularapp" : pathSrc;
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = command,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Normal,
                    WorkingDirectory = path
                };

                using (Process process = new Process { StartInfo = startInfo })
                {
                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            if (!e.Data.Contains("Switched to branch") && !e.Data.Contains("Compiling with Angular"))
                            {
                                throw new Exception("Error: " + e.Data);
                            }
                        }
                    };

                    process.Start();
                    process.BeginErrorReadLine();
                    process.BeginOutputReadLine();

                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        throw new Exception($"There was an error during processing. {action} command failed.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Update Progress
        /// </summary>
        /// <param name="step"></param>
        private async void UpdateProgress(int step)
        {
            int target = Math.Min(progressBar.Value + step, progressBar.Maximum);
            while (progressBar.Value < target)
            {
                progressBar.Value++;
                await Task.Delay(10);
            }
        }
    }
}