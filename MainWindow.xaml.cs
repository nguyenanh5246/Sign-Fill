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
        string templatePath = "template.docx";
        string namesPath = "names.txt";
        string outputPath = "output.docx";

        List<string> names = File
            .ReadAllLines(namesPath)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        SignFill.WordTemplateEngine engine = new();

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