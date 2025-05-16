using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using GongSolutions.Wpf.DragDrop;
using GongSolutions.Wpf.DragDrop.Utilities;
using System.Globalization;

namespace WpfDocCompiler
{
    // Converters moved inside the namespace to match XAML namespace reference
    public class FileNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string path)
            {
                return Path.GetFileName(path);
            }
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class FileIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string path)
            {
                string extension = Path.GetExtension(path).ToLower();

                string iconPath;
                if (extension == ".docx")
                {
                    iconPath = "/WpfDocCompiler;component/Resources/word_icon.png";
                }
                else if (extension == ".txt")
                {
                    iconPath = "/WpfDocCompiler;component/Resources/text_icon.png";
                }
                else
                {
                    iconPath = "/WpfDocCompiler;component/Resources/file_icon.png";
                }

                try
                {
                    return new BitmapImage(new Uri(iconPath, UriKind.Relative));
                }
                catch
                {
                    // If icon loading fails, return null
                    return null;
                }
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public partial class MainWindow : Window, IDropTarget
    {
        public ObservableCollection<string> SelectedFiles { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            SelectedFiles = new ObservableCollection<string>();
            filesListBox.ItemsSource = SelectedFiles;
            DataContext = this;

            // No need to add converters to resources here since they're defined in XAML
            // Remove these lines:
            // this.Resources.Add("FileNameConverter", new FileNameConverter());
            // this.Resources.Add("FileIconConverter", new FileIconConverter());

            // Set up drag and drop handler
            GongSolutions.Wpf.DragDrop.DragDrop.SetDropHandler(filesListBox, this);
        }

        #region Drag and Drop Implementation

        public void DragOver(IDropInfo dropInfo)
        {
            // Check if this is a valid drag operation (reordering within our listbox)
            if (dropInfo.Data is string && dropInfo.TargetCollection is ObservableCollection<string>)
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                dropInfo.Effects = DragDropEffects.Move;
            }
            // Check if files are being dragged from outside
            else if (dropInfo.Data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.Effects = DragDropEffects.Copy;
            }
        }

        public void Drop(IDropInfo dropInfo)
        {
            // Handle internal reordering
            if (dropInfo.Data is string sourceItem && dropInfo.TargetCollection is ObservableCollection<string> targetCollection)
            {
                int sourceIndex = SelectedFiles.IndexOf(sourceItem);
                int targetIndex = dropInfo.InsertIndex;

                if (sourceIndex != targetIndex)
                {
                    // Adjust target index if moving item down
                    if (sourceIndex < targetIndex)
                    {
                        targetIndex--;
                    }

                    SelectedFiles.RemoveAt(sourceIndex);
                    SelectedFiles.Insert(targetIndex, sourceItem);

                    filesListBox.SelectedIndex = targetIndex;
                    UpdateStatus("Item reordered");
                }
            }
            // Handle files dropped from Windows Explorer
            else if (dropInfo.Data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])dataObject.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    if (File.Exists(file))
                    {
                        SelectedFiles.Add(file);
                    }
                }
                UpdateStatus($"{files.Length} file(s) added");
            }
        }

        #endregion

        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Files (*.docx)|*.docx|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    SelectedFiles.Add(fileName);
                }
                UpdateStatus($"{openFileDialog.FileNames.Length} file(s) added");
            }
        }

        private void RemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (filesListBox.SelectedIndex != -1)
            {
                SelectedFiles.RemoveAt(filesListBox.SelectedIndex);
                UpdateStatus("File removed");
            }
            else
            {
                UpdateStatus("No file selected");
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = filesListBox.SelectedIndex;
            if (selectedIndex > 0)
            {
                string item = SelectedFiles[selectedIndex];
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(selectedIndex - 1, item);
                filesListBox.SelectedIndex = selectedIndex - 1;
                UpdateStatus("File moved up");
            }
            else
            {
                UpdateStatus("Cannot move further up");
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = filesListBox.SelectedIndex;
            if (selectedIndex != -1 && selectedIndex < SelectedFiles.Count - 1)
            {
                string item = SelectedFiles[selectedIndex];
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(selectedIndex + 1, item);
                filesListBox.SelectedIndex = selectedIndex + 1;
                UpdateStatus("File moved down");
            }
            else
            {
                UpdateStatus("Cannot move further down");
            }
        }

        private void Compile_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
            {
                UpdateStatus("No files to compile");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            // Determine if we have Word files or just text files
            bool hasWordFiles = SelectedFiles.Any(f => Path.GetExtension(f).Equals(".docx", StringComparison.OrdinalIgnoreCase));

            if (hasWordFiles)
            {
                saveFileDialog.Filter = "Word File (*.docx)|*.docx|JSON File (*.json)|*.json";
                saveFileDialog.FilterIndex = 1;
            }
            else
            {
                saveFileDialog.Filter = "Text File (*.txt)|*.txt|Word File (*.docx)|*.docx";
                saveFileDialog.FilterIndex = 1;
            }

            saveFileDialog.Title = "Save Compiled File";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    string extension = Path.GetExtension(saveFileDialog.FileName).ToLower();

                    if (extension == ".json")
                    {
                        // Extract and save as JSON
                        var allData = ExtractDataFromDocs(SelectedFiles.ToList());
                        File.WriteAllText(saveFileDialog.FileName, JsonConvert.SerializeObject(allData, Formatting.Indented));
                    }
                    else if (extension == ".docx")
                    {
                        // Check if all files are .docx or mix
                        if (SelectedFiles.All(f => Path.GetExtension(f).Equals(".docx", StringComparison.OrdinalIgnoreCase)))
                        {
                            // Merge Word documents
                            MergeWordDocuments(SelectedFiles.ToList(), saveFileDialog.FileName);
                        }
                        else
                        {
                            // Create a new Word document with text content
                            CreateWordFromMixedFiles(SelectedFiles.ToList(), saveFileDialog.FileName);
                        }
                    }
                    else // .txt
                    {
                        // Compile as plain text
                        CompileAsText(SelectedFiles.ToList(), saveFileDialog.FileName);
                    }

                    UpdateStatus("Compilation completed successfully");
                    MessageBox.Show($"File saved at: {saveFileDialog.FileName}", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    UpdateStatus("Error in compilation: " + ex.Message);
                    MessageBox.Show("Error compiling files: " + ex.Message, "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateStatus(string message)
        {
            statusTextBlock.Text = message;
        }

        private void CompileAsText(List<string> files, string outputPath)
        {
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (string filePath in files)
                {
                    if (File.Exists(filePath))
                    {
                        if (Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract text from Word files
                            string content = ExtractTextFromWord(filePath);
                            writer.Write(content);
                        }
                        else
                        {
                            // Normal text files
                            string content = File.ReadAllText(filePath);
                            writer.Write(content);
                        }
                        writer.WriteLine(); // Add a line break between files
                    }
                }
            }
        }

        private string ExtractTextFromWord(string filePath)
        {
            // Simple text extraction from Word document
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                if (wordDoc.MainDocumentPart != null)
                {
                    return wordDoc.MainDocumentPart.Document.Body.InnerText;
                }
            }
            return string.Empty;
        }

        private void CreateWordFromMixedFiles(List<string> files, string outputPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                foreach (string filePath in files)
                {
                    if (File.Exists(filePath))
                    {
                        // Add file name header
                        Paragraph fileNamePara = new Paragraph(
                            new Run(
                                new Text($"File: {Path.GetFileName(filePath)}")
                            )
                        );
                        body.AppendChild(fileNamePara);

                        string content;
                        if (Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract text from Word files
                            content = ExtractTextFromWord(filePath);
                        }
                        else
                        {
                            // Normal text files
                            content = File.ReadAllText(filePath);
                        }

                        // Add content
                        Paragraph contentPara = new Paragraph(
                            new Run(
                                new Text(content)
                            )
                        );
                        body.AppendChild(contentPara);

                        // Add page break if not the last file
                        if (filePath != files.Last())
                        {
                            Paragraph pageBreakPara = new Paragraph(
                                new Run(
                                    new Break() { Type = BreakValues.Page }
                                )
                            );
                            body.AppendChild(pageBreakPara);
                        }
                    }
                }
            }
        }

        private void MergeWordDocuments(List<string> sources, string destination)
        {
            if (File.Exists(destination)) File.Delete(destination);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(destination, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                int chunkId = 1;
                foreach (string source in sources)
                {
                    // Only process .docx files
                    if (!Path.GetExtension(source).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Add AlternativeFormatImportPart
                    AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML, "chunk" + chunkId);

                    using (FileStream fileStream = File.OpenRead(source))
                    {
                        chunk.FeedData(fileStream);
                    }

                    // Add AltChunk to the document
                    AltChunk altChunk = new AltChunk { Id = "chunk" + chunkId };
                    body.AppendChild(altChunk);

                    // Add page break between articles if not the last one
                    if (source != sources.Last())
                    {
                        Paragraph pageBreak = new Paragraph(
                            new Run(
                                new Break { Type = BreakValues.Page }
                            )
                        );
                        body.AppendChild(pageBreak);
                    }

                    chunkId++;
                }

                // Save the document
                mainPart.Document.Save();
            }

            // Reset page numbering
            ResetPageNumbering(destination);
        }

        private void ResetPageNumbering(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                // Process each section
                foreach (SectionProperties sectPr in doc.MainDocumentPart.Document.Descendants<SectionProperties>())
                {
                    // Remove existing page number types
                    sectPr.RemoveAllChildren<PageNumberType>();

                    // Add new page number type starting at 1
                    sectPr.AppendChild(new PageNumberType { Start = 1 });
                }

                // Get or create document settings part
                DocumentSettingsPart settingsPart = doc.MainDocumentPart.DocumentSettingsPart;
                if (settingsPart == null)
                {
                    settingsPart = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                }
                else if (settingsPart.Settings == null)
                {
                    settingsPart.Settings = new Settings();
                }

                // Configure footnotes to restart in each section
                FootnoteDocumentWideProperties footProps = settingsPart.Settings.Elements<FootnoteDocumentWideProperties>().FirstOrDefault();
                if (footProps == null)
                {
                    footProps = new FootnoteDocumentWideProperties();
                    settingsPart.Settings.AppendChild(footProps);
                }
                footProps.NumberingRestart = new NumberingRestart { Val = RestartNumberValues.EachSection };

                // Configure endnotes to restart in each section
                EndnoteDocumentWideProperties endProps = settingsPart.Settings.Elements<EndnoteDocumentWideProperties>().FirstOrDefault();
                if (endProps == null)
                {
                    endProps = new EndnoteDocumentWideProperties();
                    settingsPart.Settings.AppendChild(endProps);
                }
                endProps.NumberingRestart = new NumberingRestart { Val = RestartNumberValues.EachSection };

                // Save the settings
                doc.MainDocumentPart.Document.Save();
            }
        }

        private List<DocumentData> ExtractDataFromDocs(List<string> files)
        {
            var documentDataList = new List<DocumentData>();

            foreach (string file in files)
            {
                // Process only .docx files
                if (!Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    continue;

                var documentData = new DocumentData
                {
                    FileName = Path.GetFileName(file),
                    Paragraphs = new List<string>(),
                    Authors = new List<Author>()
                };

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file, false))
                {
                    if (wordDoc.MainDocumentPart != null && wordDoc.MainDocumentPart.Document.Body != null)
                    {
                        // Get all paragraphs
                        var paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

                        // Find the index of "Abstract" or "Resumo" paragraph
                        int abstractIndex = -1;
                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            string text = paragraphs[i].InnerText.Trim();
                            if (text.StartsWith("Resumo", StringComparison.OrdinalIgnoreCase) ||
                                text.StartsWith("Abstract", StringComparison.OrdinalIgnoreCase))
                            {
                                abstractIndex = i;
                                break;
                            }
                        }

                        if (abstractIndex < 0) abstractIndex = paragraphs.Count;

                        // Extract all paragraphs before the abstract
                        for (int i = 1; i < abstractIndex; i++)
                        {
                            string text = paragraphs[i].InnerText.Trim();
                            if (!string.IsNullOrEmpty(text))
                            {
                                documentData.Paragraphs.Add(text);

                                // Try to parse author
                                Author author = ParseAuthor(text);
                                if (author != null)
                                {
                                    documentData.Authors.Add(author);
                                }
                            }
                        }
                    }
                }

                documentDataList.Add(documentData);
            }

            return documentDataList;
        }

        private Author ParseAuthor(string text)
        {
            // Look for email
            var emailMatch = Regex.Match(text, @"\b[\w\.-]+@[\w\.-]+\.\w+\b");
            string email = emailMatch.Success ? emailMatch.Value : string.Empty;

            // Look for ID number
            var idMatch = Regex.Match(text, @"\b\d+\b");
            string id = idMatch.Success ? idMatch.Value : string.Empty;

            // Remove email and ID from the text
            string remaining = text;
            if (emailMatch.Success) remaining = remaining.Replace(email, "");
            if (idMatch.Success) remaining = remaining.Replace(id, "");

            // Remove common labels
            remaining = Regex.Replace(remaining, "Email|E-mail|ID|Id", "", RegexOptions.IgnoreCase).Trim();

            // Split the remaining text to get name and school
            var parts = remaining.Split(new[] { '-', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(p => p.Trim()).ToArray();

            string name = parts.Length > 0 ? parts[0] : string.Empty;
            string school = parts.Length > 1 ? parts[1] : string.Empty;

            // Only return an author if we have at least a name or email
            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(email)) return null;

            return new Author
            {
                Nome = name,
                Email = email,
                Escola = school,
                Id = id
            };
        }
    }

    public class Author
    {
        public string Nome { get; set; }
        public string Email { get; set; }
        public string Escola { get; set; }
        public string Id { get; set; }
    }

    public class DocumentData
    {
        public string FileName { get; set; }
        public List<string> Paragraphs { get; set; }
        public List<Author> Authors { get; set; }
    }
}