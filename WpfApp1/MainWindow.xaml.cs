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
using System.Globalization;
using System.Collections.Generic;
using System.Xml.Linq;

namespace WpfDocCompiler
{
    public partial class MainWindow : Window, IDropTarget
    {
        public ObservableCollection<string> SelectedFiles { get; set; }
        private string editorialFilePath;
        private Dictionary<string, List<Author>> articleAuthors;

        public MainWindow()
        {
            InitializeComponent();
            SelectedFiles = new ObservableCollection<string>();
            filesListBox.ItemsSource = SelectedFiles;
            articleAuthors = new Dictionary<string, List<Author>>();
            DataContext = this;

            UpdateStatus("Add articles and editorial to compile document");
        }

        #region Drag and Drop Implementation
        public void DragOver(IDropInfo dropInfo)
        {
            if (dropInfo.Data is string && dropInfo.TargetCollection is ObservableCollection<string>)
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                dropInfo.Effects = DragDropEffects.Move;
            }
            else if (dropInfo.Data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.Effects = DragDropEffects.Copy;
            }
        }

        public void Drop(IDropInfo dropInfo)
        {
            if (dropInfo.Data is string sourceItem && dropInfo.TargetCollection is ObservableCollection<string> targetCollection)
            {
                int sourceIndex = SelectedFiles.IndexOf(sourceItem);
                int targetIndex = dropInfo.InsertIndex;

                if (sourceIndex != targetIndex)
                {
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
            else if (dropInfo.Data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])dataObject.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    if (File.Exists(file) && Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
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
                Filter = "Word Files (*.docx)|*.docx",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    SelectedFiles.Add(fileName);
                }
                UpdateStatus($"{openFileDialog.FileNames.Length} article(s) added");
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
        }

        private void AddEditorial_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Files (*.docx)|*.docx",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                editorialFilePath = openFileDialog.FileName;
                UpdateStatus($"Editorial selected: {Path.GetFileName(editorialFilePath)}");

                // Update editorial status display
                var editorialStatusTextBlock = this.FindName("editorialStatus") as TextBlock;
                if (editorialStatusTextBlock != null)
                {
                    editorialStatusTextBlock.Text = $"Editorial: {Path.GetFileName(editorialFilePath)}";
                }

                btnCompile.IsEnabled = true;
            }
        }

        private void Compile_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("Please add at least one article.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word File (*.docx)|*.docx",
                Title = "Save Compiled Document"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    CreateDocumentWithTOC(saveFileDialog.FileName);

                    UpdateStatus($"Document compiled successfully");

                    MessageBox.Show($"Document saved successfully!\nLocation: {saveFileDialog.FileName}",
                        "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    UpdateStatus("Error: " + ex.Message);
                    MessageBox.Show("Error compiling document: " + ex.Message,
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateStatus(string message)
        {
            statusTextBlock.Text = message;
        }

        private void CreateDocumentWithTOC(string outputPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Define styles
                StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = new Styles();
                stylesPart.Styles = styles;
                AddCustomStyles(styles);

                // Extract all authors and article info
                articleAuthors.Clear();
                var allAuthors = new List<Author>();
                var articleInfoList = new List<ArticleInfo>();

                foreach (var article in SelectedFiles)
                {
                    var articleInfo = ExtractArticleInfo(article);
                    if (articleInfo != null)
                    {
                        articleInfoList.Add(articleInfo);
                        articleAuthors[article] = articleInfo.Authors;
                        allAuthors.AddRange(articleInfo.Authors);
                    }
                }

                // Remove duplicates from author list
                allAuthors = allAuthors.GroupBy(a => a.Email ?? a.Nome)
                    .Select(g => g.First())
                    .OrderBy(a => a.Nome)
                    .ToList();

                // 1. Add Author List
                AddAuthorList(body, allAuthors);

                // 2. Add Table of Contents placeholder
                AddTableOfContents(body);

                // 3. Add Editorial if exists
                if (!string.IsNullOrEmpty(editorialFilePath) && File.Exists(editorialFilePath))
                {
                    AddEditorial(body, editorialFilePath);
                }

                // 4. Add Articles
                foreach (var articleInfo in articleInfoList)
                {
                    AddArticle(body, articleInfo);
                }

                // Update fields (for TOC)
                AddSettingsToDocument(mainPart);

                mainPart.Document.Save();
            }
        }

        private void AddCustomStyles(Styles styles)
        {
            // Title style
            Style titleStyle = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "ArticleTitle",
                StyleName = new StyleName() { Val = "Article Title" }
            };

            StyleParagraphProperties titlePPr = new StyleParagraphProperties();
            titlePPr.Append(new OutlineLevel() { Val = 0 });
            titlePPr.Append(new SpacingBetweenLines() { Before = "240", After = "120" });
            titleStyle.Append(titlePPr);

            StyleRunProperties titleRPr = new StyleRunProperties();
            titleRPr.Append(new Bold());
            titleRPr.Append(new FontSize() { Val = "32" });
            titleStyle.Append(titleRPr);

            styles.Append(titleStyle);

            // Author style
            Style authorStyle = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "ArticleAuthor",
                StyleName = new StyleName() { Val = "Article Author" }
            };

            StyleParagraphProperties authorPPr = new StyleParagraphProperties();
            authorPPr.Append(new SpacingBetweenLines() { Before = "0", After = "240" });
            authorStyle.Append(authorPPr);

            StyleRunProperties authorRPr = new StyleRunProperties();
            authorRPr.Append(new Italic());
            authorRPr.Append(new FontSize() { Val = "24" });
            authorStyle.Append(authorRPr);

            styles.Append(authorStyle);
        }

        private ArticleInfo ExtractArticleInfo(string filePath)
        {
            var articleInfo = new ArticleInfo
            {
                FilePath = filePath,
                Title = Path.GetFileNameWithoutExtension(filePath),
                Authors = new List<Author>()
            };

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    if (wordDoc.MainDocumentPart != null && wordDoc.MainDocumentPart.Document.Body != null)
                    {
                        var paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

                        // First paragraph is usually the title
                        if (paragraphs.Count > 0)
                        {
                            articleInfo.Title = paragraphs[0].InnerText.Trim();
                        }

                        // Find authors (before Abstract/Resumo)
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

                        // Extract authors from paragraphs 1 to abstractIndex-1
                        for (int i = 1; i < abstractIndex && i < paragraphs.Count; i++)
                        {
                            string text = paragraphs[i].InnerText.Trim();
                            if (!string.IsNullOrEmpty(text))
                            {
                                Author author = ParseAuthor(text);
                                if (author != null)
                                {
                                    articleInfo.Authors.Add(author);
                                }
                            }
                        }

                        // Get full content
                        articleInfo.Content = wordDoc.MainDocumentPart.Document.Body.CloneNode(true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading article {Path.GetFileName(filePath)}: {ex.Message}",
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            return articleInfo;
        }

        private Author ParseAuthor(string text)
        {
            var emailMatch = Regex.Match(text, @"\b[\w\.-]+@[\w\.-]+\.\w+\b");
            string email = emailMatch.Success ? emailMatch.Value : string.Empty;

            var idMatch = Regex.Match(text, @"\b\d{5,}\b");
            string id = idMatch.Success ? idMatch.Value : string.Empty;

            string remaining = text;
            if (emailMatch.Success) remaining = remaining.Replace(email, "");
            if (idMatch.Success) remaining = remaining.Replace(id, "");

            remaining = Regex.Replace(remaining, @"Email|E-mail|ID|Id|^\d+\s*[-–]\s*", "", RegexOptions.IgnoreCase).Trim();
            remaining = remaining.Trim(' ', '-', ',', '.', '–');

            var parts = remaining.Split(new[] { '-', '–', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(p => p.Trim()).ToArray();

            string name = parts.Length > 0 ? parts[0] : string.Empty;
            string school = parts.Length > 1 ? string.Join(" - ", parts.Skip(1)) : string.Empty;

            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(email)) return null;

            return new Author
            {
                Nome = name,
                Email = email,
                Escola = school,
                Id = id
            };
        }

        private void AddAuthorList(Body body, List<Author> authors)
        {
            // Title
            Paragraph titleParagraph = new Paragraph();
            Run titleRun = new Run(new Text("Lista de Autores"));
            titleRun.RunProperties = new RunProperties(new Bold(), new FontSize() { Val = "32" });
            titleParagraph.Append(titleRun);
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new SpacingBetweenLines() { After = "240" }
            );
            body.AppendChild(titleParagraph);

            // Authors
            foreach (var author in authors)
            {
                string authorText = author.Nome;
                if (!string.IsNullOrEmpty(author.Email))
                    authorText += $" - {author.Email}";
                if (!string.IsNullOrEmpty(author.Escola))
                    authorText += $" - {author.Escola}";

                Paragraph authorParagraph = new Paragraph(new Run(new Text(authorText)));
                body.AppendChild(authorParagraph);
            }

            // Page break
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddTableOfContents(Body body)
        {
            // Title
            Paragraph titleParagraph = new Paragraph();
            Run titleRun = new Run(new Text("Índice"));
            titleRun.RunProperties = new RunProperties(new Bold(), new FontSize() { Val = "32" });
            titleParagraph.Append(titleRun);
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new SpacingBetweenLines() { After = "240" }
            );
            body.AppendChild(titleParagraph);

            // TOC Field
            Paragraph tocParagraph = new Paragraph();
            Run tocRun = new Run();
            tocRun.Append(new FieldChar() { FieldCharType = FieldCharValues.Begin });
            tocRun.Append(new FieldCode() { Text = " TOC \\o \"1-1\" \\h \\z \\u " });
            tocRun.Append(new FieldChar() { FieldCharType = FieldCharValues.Separate });
            tocRun.Append(new Text("Right-click to update field"));
            tocRun.Append(new FieldChar() { FieldCharType = FieldCharValues.End });
            tocParagraph.Append(tocRun);
            body.AppendChild(tocParagraph);

            // Page break
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddEditorial(Body body, string editorialPath)
        {
            // Title with Heading1 style for TOC
            Paragraph titleParagraph = new Paragraph();
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading1" }
            );
            Run titleRun = new Run(new Text("Editorial"));
            titleParagraph.Append(titleRun);
            body.AppendChild(titleParagraph);

            // Copy editorial content
            try
            {
                using (WordprocessingDocument editorialDoc = WordprocessingDocument.Open(editorialPath, false))
                {
                    if (editorialDoc.MainDocumentPart != null && editorialDoc.MainDocumentPart.Document.Body != null)
                    {
                        foreach (var element in editorialDoc.MainDocumentPart.Document.Body.Elements())
                        {
                            body.AppendChild(element.CloneNode(true));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                body.AppendChild(new Paragraph(new Run(new Text($"Error loading editorial: {ex.Message}"))));
            }

            // Page break
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddArticle(Body body, ArticleInfo articleInfo)
        {
            // Article title with Heading1 style for TOC
            Paragraph titleParagraph = new Paragraph();
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading1" }
            );
            Run titleRun = new Run(new Text(articleInfo.Title));
            titleParagraph.Append(titleRun);
            body.AppendChild(titleParagraph);

            // Authors in italic (nome 1 / nome 2 / nome 3)
            if (articleInfo.Authors.Count > 0)
            {
                Paragraph authorParagraph = new Paragraph();
                Run authorRun = new Run();
                authorRun.RunProperties = new RunProperties(new Italic());
                string authorNames = string.Join(" / ", articleInfo.Authors.Select(a => a.Nome));
                authorRun.Append(new Text(authorNames));
                authorParagraph.Append(authorRun);
                authorParagraph.ParagraphProperties = new ParagraphProperties(
                    new SpacingBetweenLines() { After = "240" }
                );
                body.AppendChild(authorParagraph);
            }

            // Article content
            if (articleInfo.Content != null)
            {
                // Skip the original title and author paragraphs when copying content
                var elements = ((Body)articleInfo.Content).Elements().ToList();
                int startIndex = Math.Min(articleInfo.Authors.Count + 1, elements.Count);

                for (int i = startIndex; i < elements.Count; i++)
                {
                    body.AppendChild(elements[i].CloneNode(true));
                }
            }

            // Page break after article
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddSettingsToDocument(MainDocumentPart mainPart)
        {
            DocumentSettingsPart settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            Settings settings = new Settings();
            settings.Append(new UpdateFieldsOnOpen() { Val = true });
            settingsPart.Settings = settings;
        }
    }

    // Support classes
    public class Author
    {
        public string Nome { get; set; }
        public string Email { get; set; }
        public string Escola { get; set; }
        public string Id { get; set; }
    }

    public class ArticleInfo
    {
        public string FilePath { get; set; }
        public string Title { get; set; }
        public List<Author> Authors { get; set; }
        public OpenXmlNode Content { get; set; }
    }

    // Converters remain the same
    public class FileNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string filePath)
            {
                return Path.GetFileName(filePath);
            }
            return value;
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
            if (value is string filePath)
            {
                return new BitmapImage(new Uri("pack://application:,,,/Images/word_icon.png"));
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}