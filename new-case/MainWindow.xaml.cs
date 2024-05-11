using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Runtime.InteropServices;
using IWshRuntimeLibrary;
using System.Reflection;

namespace new_case
{

    public partial class MainWindow : Window
    {
        string templateFolder = @"D:\CASE_FOLDER_STRUCTURE";
        long totalBytesCopied = 0;
        private static readonly Version? appVersion = new(Assembly.GetExecutingAssembly().GetName().Version.ToString(2));
        public MainWindow()
        {
            InitializeComponent();
            DateTime dateTime = DateTime.Now;
            string thisYear = dateTime.Year.ToString();
            NewCaseTextBox.Text = thisYear + '-';
            SourceTextBox.Text = templateFolder;
            VersionNumber.Content = $"v{appVersion}";
        }
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void ResetProgressBar()
        {
            Dispatcher.Invoke(async () =>
            {
                FileProgressBar.Value = 0;
                TotalProgressBar.Value = 0;
                FileProgressLabel.Content = "";
                TotalProgressLabel.Content = "";
                totalBytesCopied = 0;
                await Task.Delay(1);
            });
        }
        private void ExpandWindow()
        {
            CaseWindow.Height = 345;
            FileProgressBar.Visibility = Visibility.Visible;
            FileProgressLabel.Visibility = Visibility.Visible;
            TotalProgressBar.Visibility = Visibility.Visible;
            TotalProgressLabel.Visibility = Visibility.Visible;
        }
        private void CollapseWindow()
        {
            CaseWindow.Height = 220;
            FileProgressBar.Visibility = Visibility.Hidden;
            FileProgressLabel.Visibility = Visibility.Hidden;
            TotalProgressBar.Visibility = Visibility.Hidden;
            TotalProgressLabel.Visibility = Visibility.Hidden;
        }
        private async void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ResetProgressBar();
            var fileProgress = new Progress<double>(value => FileProgressBar.Value = value);
            var totalProgress = new Progress<double>(value => TotalProgressBar.Value = value);
            string sourcePath;
            if (SourceTextBox.Text.Length == 0)
            {
                sourcePath = templateFolder;
            }
            else
            {
                sourcePath = SourceTextBox.Text;
                templateFolder = sourcePath;
            }
            string destinationPath = DestinationTextBox.Text;
            if (!Directory.Exists(sourcePath))
            {
                MessageBoxResult result = MessageBox.Show(@$"The source directory {sourcePath} does not exist! Click OK to select a new source path, or Cancel to abort.", "Soure directory not found", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                if (result == MessageBoxResult.OK)
                {
                    System.Windows.Forms.DialogResult selectSourceResult = SelectSource();
                    if (selectSourceResult == System.Windows.Forms.DialogResult.Cancel)
                    {
                        return;
                    }
                    else
                    {
                        sourcePath = SourceTextBox.Text;
                        templateFolder = SourceTextBox.Text;
                    }
                }
                else
                {
                    return;
                }
            }
            if ((!Directory.Exists(destinationPath)) || (destinationPath.Length == 0))
            {
                MessageBoxResult result = MessageBox.Show(@$"The destination directory {destinationPath} does not exist! Click OK to select a new destination path, or Cancel to abort.", "Destination directory not found", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                if (result == MessageBoxResult.OK)
                {
                    System.Windows.Forms.DialogResult selectDestinationResult = SelectDestination();
                    if (selectDestinationResult == System.Windows.Forms.DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
            }
            ExpandWindow();
            if (destinationPath.EndsWith($@"\"))
            {
                destinationPath = destinationPath.TrimEnd(new[] { '\\' });
            }
            destinationPath = @$"{destinationPath}\{NewCaseTextBox.Text}";
            try
            {
                await CopyFolder(sourcePath, destinationPath, fileProgress, totalProgress);
                CreateShortcut(destinationPath);
                MessageBox.Show(@$"Completed copying {sourcePath} to {destinationPath}", "Copy Complete!", MessageBoxButton.OK, MessageBoxImage.Information);
                ResetProgressBar();
                CollapseWindow();
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Error attempting to copy files and create shortcut: {ex}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private System.Windows.Forms.DialogResult SelectSource()
        {
            string selectedPath = "";
            System.Windows.Forms.FolderBrowserDialog folderDlg = new()
            {
                Description = "Select the directory containing the case template",
                ShowNewFolderButton = true,
                UseDescriptionForTitle = true,
                RootFolder = Environment.SpecialFolder.Desktop,
                InitialDirectory = Environment.CurrentDirectory
            };
            System.Windows.Forms.DialogResult result = folderDlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                selectedPath = folderDlg.SelectedPath;
                templateFolder = selectedPath;
            }
            if (selectedPath != "")
            {
                SourceTextBox.Text = selectedPath;
            }
            return result;
        }
        private void SourcePicker(object sender, RoutedEventArgs e)
        {
            SelectSource();
        }
        private System.Windows.Forms.DialogResult SelectDestination()
        {
            string selectedPath = "";
            System.Windows.Forms.FolderBrowserDialog folderDlg = new()
            {
                Description = "Select the directory where you would like to create the case folder",
                ShowNewFolderButton = true,
                UseDescriptionForTitle = true,
                RootFolder = Environment.SpecialFolder.Desktop,
                InitialDirectory = Environment.CurrentDirectory
            };
            System.Windows.Forms.DialogResult result = folderDlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                selectedPath = folderDlg.SelectedPath;
            }
            if (selectedPath != "")
            {
                DestinationTextBox.Text = selectedPath;
            }
            return result;
        }
        private void DestinationPicker(object sender, RoutedEventArgs e)
        {
            SelectDestination();
        }

        public async Task CopyFolder(string sourceFolder, string destinationPath, IProgress<double> fileProgress, IProgress<double> totalProgress)
        {
            try
            {
                long fileBytesCopied = 0;
                double fileProgressPercentage;
                double totalProgressPercentage;
                if (!Directory.Exists(destinationPath))
                {
                    var templateInfo = new DirectoryInfo(templateFolder);
                    long templateFolderSize = templateInfo.EnumerateFiles("*", SearchOption.AllDirectories).Sum(x => x.Length);
                    var sourceDir = new DirectoryInfo(sourceFolder);
                    long folderSize = sourceDir.EnumerateFiles("*", SearchOption.AllDirectories).Sum(x => x.Length);
                    DirectoryInfo[] allDirs = sourceDir.GetDirectories();
                    Directory.CreateDirectory(destinationPath);

                    foreach (FileInfo file in sourceDir.GetFiles())
                    {
                        string fileDest = Path.Combine(destinationPath, file.Name);
                        file.CopyTo(fileDest);
                        FileProgressLabel.Content = fileDest;
                        totalBytesCopied += file.Length;
                        fileBytesCopied += file.Length;
                        fileProgressPercentage = (double)fileBytesCopied / folderSize * 100;
                        fileProgress.Report(fileProgressPercentage);
                        await Task.Delay(1);
                    }
                    totalProgressPercentage = (double)totalBytesCopied / templateFolderSize * 100;
                    TotalProgressLabel.Content = (int)totalProgressPercentage + " / 100%";
                    totalProgress.Report(totalProgressPercentage);
                    await Task.Delay(1);
                    foreach (DirectoryInfo subDir in allDirs)
                    {
                        string dirDest = Path.Combine(destinationPath, subDir.Name);
                        await CopyFolder(subDir.FullName, dirDest, fileProgress, totalProgress);
                    }
                }
                else
                {
                    MessageBoxResult result = MessageBox.Show($"Directory {destinationPath} exists! Delete and recreate?", "Directory exists!", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (result == MessageBoxResult.Yes)
                    {
                        Directory.Delete(destinationPath, true);
                        await CopyFolder(templateFolder, destinationPath, fileProgress, totalProgress);
                    }
                }
            }
            catch (Exception ex)
            { 
                MessageBox.Show($@"An error occurred while copying: {ex}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void CreateShortcut(string filePath)
        {
            string currentUser = Environment.UserName;
            string shortcutPath = $@"C:\Users\{currentUser}\Desktop\{NewCaseTextBox.Text}.lnk";
            FileProgressLabel.Content = $@"Creating {shortcutPath}";
            WshShell shell = new();
            try
            {
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = $@"{filePath}\xwf\xwforensics64.exe";
                shortcut.Save();
                Marshal.ReleaseComObject(shortcut);
                Marshal.ReleaseComObject(shell);
                byte[] shortcutBytes = System.IO.File.ReadAllBytes(shortcutPath);
                shortcutBytes[0x15] |= 0x20;
                System.IO.File.WriteAllBytes(shortcutPath, shortcutBytes);
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"An error occurred while trying to create the Desktop shortcut: {ex}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
