using System.Windows;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Sign_Fill;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        TestGenerate();
    }

    private void TestGenerate()
    {
        string templatePath = "template_table.docx"; // file Word template có chứa placeholder {{SIGNER_1}}, {{SIGNER_2}}, ...
        string namesPath = "names.txt"; // file chứa danh sách tên, mỗi tên trên một dòng
        string outputPath = "output_table.docx"; // file kết quả sau khi fill

        List<string> names = File
            .ReadAllLines(namesPath)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        Sign_Fill.WordTemplateEngine engine = new();

        engine.Generate(
            templatePath,
            outputPath,
            names
        );

        MessageBox.Show(
            "Đã tạo file:\n" + outputPath,
            "SignFill"
        );
    }
}