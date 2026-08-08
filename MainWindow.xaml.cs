using SignFill.Models;
using SignFill.Services;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace Sign_Fill;

public partial class MainWindow : Window
{
    private readonly FileResolver _fileResolver;
    private readonly ExcelReader _excelReader;
    private readonly TemplateParser _templateParser;
    private readonly ValidationService _validationService;
    private readonly BatchGenerator _batchGenerator;
    private DateTime _lastExcelWriteTime = DateTime.MinValue;
    private DateTime _lastTemplateWriteTime = DateTime.MinValue;

    // Check file existence every 500ms to update the status in real-time.
    private readonly DispatcherTimer _fileCheckTimer;
    private bool _lastExcelFound;
    private bool _lastTemplateFound;

    public MainWindow()
    {
        InitializeComponent();

        FolderManager.EnsureDirectories();

        _fileResolver = new FileResolver();
        _excelReader = new ExcelReader();
        _templateParser = new TemplateParser();
        _validationService = new ValidationService();
        _batchGenerator = new BatchGenerator();

        CheckInputFiles();

        try
        {
            _fileResolver.GetExcelFile();
            _lastExcelFound = true;
        }
        catch
        {
            _lastExcelFound = false;
        }

        try
        {
            _fileResolver.GetTemplateFile();
            _lastTemplateFound = true;
        }
        catch
        {
            _lastTemplateFound = false;
        }

        LoadAvailableFields();

        _fileCheckTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(500)
        };

        _fileCheckTimer.Tick += FileCheckTimer_Tick;
        _fileCheckTimer.Start();
    }
    private (bool excelFound, bool templateFound) CheckInputFiles()
    {
        bool excelFound = false;
        bool templateFound = false;

        try
        {
            _fileResolver.GetExcelFile();

            InputStatusIcon.Text = "✓";
            InputStatusIcon.Foreground =
                new SolidColorBrush(Colors.Green);

            InputStatusText.Text =
                "Excel detected";

            excelFound = true;
        }
        catch
        {
            InputStatusIcon.Text = "✗";
            InputStatusIcon.Foreground =
                new SolidColorBrush(Colors.Red);

            InputStatusText.Text =
                "Excel not found";
        }

        try
        {
            _fileResolver.GetTemplateFile();

            TemplateStatusIcon.Text = "✓";
            TemplateStatusIcon.Foreground =
                new SolidColorBrush(Colors.Green);

            TemplateStatusText.Text =
                "Word template detected";

            templateFound = true;
        }
        catch
        {
            TemplateStatusIcon.Text = "✗";
            TemplateStatusIcon.Foreground =
                new SolidColorBrush(Colors.Red);

            TemplateStatusText.Text =
                "Word template not found";
        }

        if (!excelFound)
        {
            FieldPanel.Children.Clear();
        }
        return (excelFound, templateFound);
    }

    private void GenerateButton_Click(
        object sender,
        RoutedEventArgs e)
    {
        try
        {
            StatusText.Text =
                "Status: Processing...";

            GenerateButton.IsEnabled = false;

            string excelPath =
                _fileResolver.GetExcelFile();

            string templatePath =
                _fileResolver.GetTemplateFile();

            string fileNamePattern =
                FileNamePatternTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(
                    fileNamePattern))
            {
                throw new InvalidDataException(
                    "Output filename pattern không được để trống.");
            }

            List<DataRow> rows =
                _excelReader.Read(excelPath);

            List<string> templateFields =
                _templateParser.GetFields(
                    templatePath);

            _validationService.Validate(
                rows,
                templateFields);

            int generatedCount =
                _batchGenerator.Generate(
                    templatePath,
                    FolderManager.OutputDirectory,
                    rows,
                    fileNamePattern);

            StatusText.Text =
                $"Status: Success - " +
                $"{generatedCount} file(s) generated.";

            MessageBox.Show(
                $"Generate thành công!\n\n" +
                $"Đã tạo {generatedCount} file.",
                "SignFill");
        }
        catch (Exception ex)
        {
            StatusText.Text =
                "Status: Failed";

            MessageBox.Show(
                ex.Message,
                "SignFill - Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            GenerateButton.IsEnabled = true;
        }
    }
    private void LoadAvailableFields()
    {
        FieldPanel.Children.Clear();

        try
        {
            string excelPath =
                _fileResolver.GetExcelFile();

            List<DataRow> rows =
                _excelReader.Read(excelPath);

            if (rows.Count == 0)
            {
                return;
            }

            foreach (string fieldName in rows[0].Fields.Keys)
            {
                Button fieldButton = new()
                {
                    Content = $"{{{fieldName}}}",
                    Margin = new Thickness(0, 0, 8, 8),
                    Padding = new Thickness(10, 5, 10, 5),
                    Tag = fieldName
                };

                fieldButton.Click += FieldButton_Click;

                FieldPanel.Children.Add(fieldButton);
            }
        }
        catch
        {
            // Excel chưa sẵn sàng.
            // Trạng thái lỗi sẽ được xử lý ở CheckInputFiles().
        }
    }
    private void FieldButton_Click(
    object sender,
    RoutedEventArgs e)
    {
        if (sender is not Button button)
        {
            return;
        }

        if (button.Tag is not string fieldName)
        {
            return;
        }

        string field =
            $"{{{fieldName}}}";

        int selectionStart =
            FileNamePatternTextBox.SelectionStart;

        string currentText =
            FileNamePatternTextBox.Text;

        FileNamePatternTextBox.Text =
            currentText.Insert(
                selectionStart,
                field);

        FileNamePatternTextBox.SelectionStart =
            selectionStart + field.Length;

        FileNamePatternTextBox.Focus();
    }
    private void RefreshFileStatus()
    {
        // =========================
        // CHECK EXCEL
        // =========================

        try
        {
            string excelPath =
                _fileResolver.GetExcelFile();

            DateTime writeTime =
                File.GetLastWriteTime(excelPath);

            // Excel mới xuất hiện hoặc bị thay đổi
            if (!_lastExcelFound ||
                writeTime != _lastExcelWriteTime)
            {
                _lastExcelFound = true;
                _lastExcelWriteTime = writeTime;

                CheckInputFiles();
                LoadAvailableFields();
            }
        }
        catch
        {
            // Excel đã bị xóa
            if (_lastExcelFound)
            {
                _lastExcelFound = false;
                _lastExcelWriteTime = DateTime.MinValue;

                CheckInputFiles();
                FieldPanel.Children.Clear();
            }
        }


        // =========================
        // CHECK TEMPLATE
        // =========================

        try
        {
            string templatePath =
                _fileResolver.GetTemplateFile();

            DateTime writeTime =
                File.GetLastWriteTime(templatePath);

            // Template mới xuất hiện hoặc bị thay đổi
            if (!_lastTemplateFound ||
                writeTime != _lastTemplateWriteTime)
            {
                _lastTemplateFound = true;
                _lastTemplateWriteTime = writeTime;

                CheckInputFiles();
            }
        }
        catch
        {
            // Template đã bị xóa
            if (_lastTemplateFound)
            {
                _lastTemplateFound = false;
                _lastTemplateWriteTime = DateTime.MinValue;

                CheckInputFiles();
            }
        }
    }
    private void FileCheckTimer_Tick(
    object? sender,
    EventArgs e)
    {
        RefreshFileStatus();
    }
}