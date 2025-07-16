using Microsoft.Win32;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Input;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

namespace Office2PDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindowViewModel ViewModel { get; set; }

        public enum LogLevel
        {
            Trace,
            Info,
            Warning,
            Error
        }

        public MainWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
            ViewModel = new MainWindowViewModel();
            DataContext = ViewModel;
            LogRichTextBox.Document.Blocks.Clear();
            LoadEmbeddedImage();
        }

        private void BrowseFromFolder_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog()
            {
                Title = "选择来源文件夹",
                InitialDirectory = ViewModel.FromRootFolderPath
            };
            if (folderDialog.ShowDialog() == true)
            {
                ViewModel.FromRootFolderPath = folderDialog.FolderName;
            }
        }

        private void BrowseToFolder_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog()
            {
                Title = "选择目标文件夹",
                InitialDirectory = ViewModel.ToRootFolderPath
            };
            if (folderDialog.ShowDialog() == true)
            {
                ViewModel.ToRootFolderPath = folderDialog.FolderName;
            }
        }

        private void OpenFromFolder_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(ViewModel.FromRootFolderPath))
            {
                using (var process = new System.Diagnostics.Process())
                {
                    process.StartInfo.FileName = "explorer.exe";
                    process.StartInfo.Arguments = ViewModel.FromRootFolderPath;
                    process.Start();
                }
            }
            else
            {
                MessageBox.Show("来源文件夹不存在，请重新选择", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OpenToFolder_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(ViewModel.ToRootFolderPath))
            {
                using (var process = new System.Diagnostics.Process())
                {
                    process.StartInfo.FileName = "explorer.exe";
                    process.StartInfo.Arguments = ViewModel.ToRootFolderPath;
                    process.Start();
                }
            }
            else
            {
                MessageBox.Show("目标文件夹不存在，请重新选择", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void StartConvert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ViewModel.CanStartConvert = false;

                var fileTypeHandlers = new (string TypeName, bool IsConvert, string[] Extensions, Action<string, string[]> ConvertAction)[] {
                    ("Word", ViewModel.IsConvertWord, [".doc", ".docx"], (typeName, files) => ConvertToPDF<WordApplication>(typeName, files)),
                    ("Excel", ViewModel.IsConvertExcel,[".xls", ".xlsx"], (typeName, files) => ConvertToPDF<ExcelApplication>(typeName, files)),
                    ("PPT", ViewModel.IsConvertPPT, [".ppt", ".pptx"], (typeName, files) => ConvertToPDF<PowerPointApplication>(typeName, files))
                };

                AppendLog($"==============开始转换==============");

                foreach (var (typeName, isConvert, extensions, convertAction) in fileTypeHandlers)
                {
                    try
                    {
                        if (isConvert)
                        {
                            AppendLog($"{typeName} 转换开始", LogLevel.Info);

                            var searchOption = ViewModel.IsConvertChildrenFolder ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                            var files = Directory.EnumerateFiles(ViewModel.FromRootFolderPath, "*.*", searchOption)
                             .Where(file => extensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase)));

                            if (!files.Any())
                            {
                                AppendLog($"无 {typeName} 文件", LogLevel.Warning);
                            }
                            else
                            {
                                await Task.Run(() => convertAction(typeName, files.ToArray()));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"{typeName} 转换错误：{ex.Message}", LogLevel.Error);
                    }
                    finally
                    {
                        if (isConvert)
                        {
                            AppendLog($"{typeName} 转换结束", LogLevel.Info);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"转换错误：{ex.Message}", LogLevel.Error);
            }
            finally
            {
                AppendLog($"==============转换结束==============\n");

                ViewModel.CanStartConvert = true;
            }
        }

        public void ConvertToPDF<T>(string typeName, string[] fromFilePaths) where T : IOfficeApplication, new()
        {
            using var application = new T();
            try
            {
                AppendLog($"{typeName} 打开进程中...");
                var numberFormat = $"D{fromFilePaths.Length.ToString().Length}";
                foreach ((var index, var fromFilePath) in fromFilePaths.Index())
                {
                    try
                    {
                        if (application is WordApplication wordApp)
                        {
                            wordApp.IsPrintRevisions = ViewModel.IsPrintRevisionsInWord;
                        }
                        else if (application is ExcelApplication excelApp)
                        {
                            excelApp.IsConvertOneSheetOnePDF = ViewModel.IsConvertOneSheetOnePDFInExcel;
                            excelApp.IsPrintToOneColumn = ViewModel.IsPrintToOneColumnInExcel;
                        }
                        application.OpenDocument(fromFilePath);
                        var toFilePath = GetToFilePath(ViewModel.FromRootFolderPath, ViewModel.ToRootFolderPath, fromFilePath, Path.GetFileName(fromFilePath));
                        application.SaveAsPDF(toFilePath);
                        AppendLog($"（{index.ToString(numberFormat)}）{typeName} 转换成功: {Path.GetRelativePath(ViewModel.ToRootFolderPath, toFilePath)}");
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"（{index.ToString(numberFormat)}）{typeName} 转换出错：{fromFilePath} {ex.Message}", LogLevel.Error);
                    }
                    finally
                    {
                        application.CloseDocument();
                    }
                }
            }
            catch (Exception e)
            {
                AppendLog(e.Message, LogLevel.Error);
            }
            finally
            {
                AppendLog($"{typeName} 所有文件已转换完毕，关闭进程中...");
            }
        }

        private string GetToFilePath(string fromRootFolderPath, string toFolderRootPath, string fromFilePath, string toFileName)
        {
            var relativePath = !ViewModel.IsKeepFolderStructure ? "." : Path.GetRelativePath(fromRootFolderPath, Path.GetDirectoryName(fromFilePath) ?? "");
            var toFolderPath = Path.Combine(toFolderRootPath, relativePath);
            if (!Directory.Exists(toFolderPath)) Directory.CreateDirectory(toFolderPath);
            return Path.Combine(toFolderPath, Path.ChangeExtension(toFileName, ".pdf"));
        }

        public void AppendLog(string message, LogLevel level = LogLevel.Trace)
        {
            Dispatcher.Invoke(() => // 在 UI 线程中更新日志
            {
                var paragraph = new System.Windows.Documents.Paragraph();
                paragraph.Inlines.Add(new Run($"[{DateTime.Now:HH:mm:ss}] ") { Foreground = Brushes.Gray });

                Brush color = Brushes.Black;
                switch (level)
                {
                    case LogLevel.Trace: color = Brushes.Gray; break;
                    case LogLevel.Info: color = Brushes.Green; break;
                    case LogLevel.Warning: color = Brushes.Orange; break;
                    case LogLevel.Error: color = Brushes.Red; break;
                }

                paragraph.Inlines.Add(new Run(message) { Foreground = color });
                LogRichTextBox.Document.Blocks.Add(paragraph);
                LogRichTextBox.ScrollToEnd();
            });
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });
            e.Handled = true;
        }

        private void LoadEmbeddedImage()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "Office2PDF.Resources.qrcode.png";

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = stream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        QrcodeForGHImage.Source = bitmap;
                    }
                    else
                    {
                        MessageBox.Show($"找不到嵌入资源: {resourceName}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载图片失败: {ex.Message}");
            }
        }

        private void GongZhongHao_MouseEnter(object sender, MouseEventArgs e)
        {
            imagePopup.IsOpen = true;
        }

        private void GongZhongHao_MouseLeave(object sender, MouseEventArgs e)
        {
            imagePopup.IsOpen = false;
        }

    }

    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private string _fromFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Office2PDF", "source");
        public string FromRootFolderPath
        {
            get => _fromFolderPath;
            set
            {
                if (_fromFolderPath != value)
                {
                    _fromFolderPath = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _toFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Office2PDF", "output");
        public string ToRootFolderPath
        {
            get => _toFolderPath;
            set
            {
                if (_toFolderPath != value)
                {
                    _toFolderPath = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isConvertWord = true;
        public bool IsConvertWord
        {
            get => _isConvertWord;
            set
            {
                if (_isConvertWord != value)
                {
                    _isConvertWord = value;
                    OnPropertyChanged();
                    UpdateIsConvertAll();
                }
            }
        }

        private bool _isConvertPPT = true;
        public bool IsConvertPPT
        {
            get => _isConvertPPT;
            set
            {
                if (_isConvertPPT != value)
                {
                    _isConvertPPT = value;
                    OnPropertyChanged();
                    UpdateIsConvertAll();
                }
            }
        }

        private bool _isConvertExcel = true;
        public bool IsConvertExcel
        {
            get => _isConvertExcel;
            set
            {
                if (_isConvertExcel != value)
                {
                    _isConvertExcel = value;
                    OnPropertyChanged();
                    UpdateIsConvertAll();
                }
            }
        }

        private bool _isConvertAll = true;
        public bool IsConvertAll
        {
            get => _isConvertAll;
            set
            {
                if (_isConvertAll != value)
                {
                    _isConvertAll = value;
                    _isConvertWord = value;
                    _isConvertPPT = value;
                    _isConvertExcel = value;

                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsConvertWord));
                    OnPropertyChanged(nameof(IsConvertPPT));
                    OnPropertyChanged(nameof(IsConvertExcel));
                }
            }
        }

        private bool _isConvertChildrenFolder = true;
        public bool IsConvertChildrenFolder
        {
            get => _isConvertChildrenFolder;
            set
            {
                if (_isConvertChildrenFolder != value)
                {
                    _isConvertChildrenFolder = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isKeepFolderStructure = true;
        public bool IsKeepFolderStructure
        {
            get => _isKeepFolderStructure;
            set
            {
                if (_isKeepFolderStructure != value)
                {
                    _isKeepFolderStructure = value;
                    OnPropertyChanged();
                }
            }
        }


        private bool _isPrintRevisionsInWord = true;
        public bool IsPrintRevisionsInWord
        {
            get => _isPrintRevisionsInWord;
            set
            {
                if (_isPrintRevisionsInWord != value)
                {
                    _isPrintRevisionsInWord = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isConvertOneSheetOnePDFInExcel = true;
        public bool IsConvertOneSheetOnePDFInExcel
        {
            get => _isConvertOneSheetOnePDFInExcel;
            set
            {
                if (_isConvertOneSheetOnePDFInExcel != value)
                {
                    _isConvertOneSheetOnePDFInExcel = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isPrintToOneColumnInExcel = false;
        public bool IsPrintToOneColumnInExcel
        {
            get => _isPrintToOneColumnInExcel;
            set
            {
                if (_isPrintToOneColumnInExcel != value)
                {
                    _isPrintToOneColumnInExcel = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _canStartConvert = true;
        public bool CanStartConvert
        {
            get => _canStartConvert;
            set
            {
                if (_canStartConvert != value)
                {
                    _canStartConvert = value;
                    OnPropertyChanged();
                }
            }
        }

        private void UpdateIsConvertAll()
        {
            _isConvertAll = IsConvertWord && IsConvertPPT && IsConvertExcel;
            OnPropertyChanged(nameof(IsConvertAll));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
