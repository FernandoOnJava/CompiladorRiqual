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

        // This function shows the editorial form and gets the content
        private string GetEditorialContent()
        {
            // Create a new editorial form
            EditorialForm editorialForm = new EditorialForm();
            editorialForm.Owner = this;

            // Show the form as a dialog and wait for user input
            bool? result = editorialForm.ShowDialog();

            // Return the editorial content if the user clicked OK, otherwise null
            return result == true ? editorialForm.EditorialContent : null;
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
            // Get the editorial content first
            string editorialContent = GetEditorialContent();
            if (editorialContent == null)
            {
                UpdateStatus("Compilação cancelada - sem Editorial");
                return;
            }

            // Gather articles from the user
            var articles = GetArticlesFromUser();
            if (articles == null || articles.Count == 0)
            {
                UpdateStatus("Nenhum artigo selecionado.");
                return;
            }

            // Create the author list
            var authors = CreateAuthorList(articles);

            // Proceed to create the document
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word File (*.docx)|*.docx",
                Title = "Save Compiled File"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    CreateDocumentWithEditorialAndArticles(saveFileDialog.FileName, editorialContent, authors, articles);
                    MessageBox.Show($"Ficheiro guardado em: {saveFileDialog.FileName}", "Sucesso",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    UpdateStatus("Erro na compilação: " + ex.Message);
                    MessageBox.Show("Erro ao compilar ficheiros: " + ex.Message, "Erro",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private List<string> GetArticlesFromUser()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Files (*.docx)|*.docx|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileNames.ToList();
            }
            return null;
        }
        private List<Author> CreateAuthorList(List<string> articles)
        {
            var authors = new List<Author>();
            foreach (var article in articles)
            {
                if (Path.GetExtension(article).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    var docData = ExtractAuthorData(article);
                    authors.AddRange(docData.Authors);
                }
            }
            return authors.Distinct().ToList(); // Remove duplicates if necessary
        }

        private void CreateDocumentWithEditorialAndArticles(string outputPath, string editorialContent, List<Author> authors, List<string> articles)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Add Editorial
                AddEditorialPage(body, editorialContent);

                // Add Author List
                AddAuthorsPage(body, authors);

                // Add Articles
                foreach (var article in articles)
                {
                    if (File.Exists(article))
                    {
                        string content = ExtractTextFromWord(article);
                        Paragraph contentParagraph = new Paragraph(new Run(new Text(content)));
                        body.AppendChild(contentParagraph);
                        body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page }))); // Page break after each article
                    }
                }

                mainPart.Document.Save();
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

        private void CreateWordFromMixedFiles(List<string> files, string outputPath, string editorialContent)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Extract all authors from documents
                var allAuthors = new List<Author>();
                foreach (string file in files)
                {
                    if (Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        var docData = ExtractAuthorData(file);
                        if (docData.Authors.Count > 0)
                        {
                            allAuthors.AddRange(docData.Authors);
                        }
                    }
                }

                // Add authors list page
                AddAuthorsPage(body, allAuthors);

                // Add page break after authors page
                Paragraph pageBreakAfterAuthors = new Paragraph(
                    new Run(
                        new Break { Type = BreakValues.Page }
                    )
                );
                body.AppendChild(pageBreakAfterAuthors);

                // Add Editorial page if content was provided
                if (!string.IsNullOrWhiteSpace(editorialContent))
                {
                    AddEditorialPage(body, editorialContent);

                    // Add page break after editorial page
                    Paragraph pageBreakAfterEditorial = new Paragraph(
                        new Run(
                            new Break { Type = BreakValues.Page }
                        )
                    );
                    body.AppendChild(pageBreakAfterEditorial);
                }

                // Now proceed with adding each file as before
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

            // Update status with authors count
            UpdateStatus($"Compilação concluída com {files.Count} ficheiros");
        }
        private void AddEditorialPage(Body body, string editorialContent)
        {
            // "Editorial" header paragraph - Bold, Aptos Display
            Paragraph titleParagraph = new Paragraph();
            Run titleRun = new Run(new Text("Editorial"));
            RunProperties titleRunProps = new RunProperties(
                new Bold(),
                new RunFonts() { Ascii = "Aptos Display", HighAnsi = "Aptos Display" },
                new FontSize() { Val = "28" }  // 14pt = 28 half-points (slightly larger than author list title)
            );
            titleRun.PrependChild(titleRunProps);
            titleParagraph.AppendChild(titleRun);
            body.AppendChild(titleParagraph);

            // Add a paragraph with the editorial content
            Paragraph contentParagraph = new Paragraph();
            ParagraphProperties pPr = new ParagraphProperties(
                new SpacingBetweenLines() { After = "0", Before = "240" } // 12pt before (240 twentiethPoints)
            );
            contentParagraph.AppendChild(pPr);

            // Editorial content in Times New Roman
            Run contentRun = new Run(new Text(editorialContent));
            RunProperties contentRunProps = new RunProperties(
                new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                new FontSize() { Val = "24" }  // 12pt = 24 half-points
            );
            contentRun.PrependChild(contentRunProps);
            contentParagraph.AppendChild(contentRun);
            body.AppendChild(contentParagraph);
        }

        private void MergeWordDocuments(List<string> sources, string destination, string editorialContent)
        {
            if (File.Exists(destination)) File.Delete(destination);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(destination, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Extract all authors from documents
                var allAuthors = new List<Author>();
                foreach (string source in sources)
                {
                    if (Path.GetExtension(source).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        var docData = ExtractAuthorData(source);
                        if (docData.Authors.Count > 0)
                        {
                            allAuthors.AddRange(docData.Authors);
                        }
                    }
                }

                // Add authors list page
                AddAuthorsPage(body, allAuthors);

                // Add page break after authors page
                Paragraph pageBreakAfterAuthors = new Paragraph(
                    new Run(
                        new Break { Type = BreakValues.Page }
                    )
                );
                body.AppendChild(pageBreakAfterAuthors);

                // Add Editorial page if content was provided
                if (!string.IsNullOrWhiteSpace(editorialContent))
                {
                    AddEditorialPage(body, editorialContent);

                    // Add page break after editorial page
                    Paragraph pageBreakAfterEditorial = new Paragraph(
                        new Run(
                            new Break { Type = BreakValues.Page }
                        )
                    );
                    body.AppendChild(pageBreakAfterEditorial);
                }

                // Now add all documents
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

            // Update status with authors count
            UpdateStatus($"Compilação concluída com {sources.Count} artigos");
        }

        private void AddAuthorsPage(Body body, List<Author> authors)
        {
            // Remove duplicate authors by email (if available) or by name
            var uniqueAuthors = new List<Author>();
            foreach (var author in authors)
            {
                if (!string.IsNullOrEmpty(author.Email))
                {
                    if (!uniqueAuthors.Any(a => a.Email == author.Email))
                    {
                        uniqueAuthors.Add(author);
                    }
                }
                else if (!string.IsNullOrEmpty(author.Nome))
                {
                    if (!uniqueAuthors.Any(a => a.Nome == author.Nome))
                    {
                        uniqueAuthors.Add(author);
                    }
                }
            }

            // Sort authors by name
            uniqueAuthors = uniqueAuthors.OrderBy(a => a.Nome).ToList();

            // Document page margins - 2.5cm on all sides
            SectionProperties sectionProps = new SectionProperties();
            PageMargin pageMargin = new PageMargin()
            {
                Top = 1440, // 2.5cm in twips (1440 twips = 2.54cm)
                Right = 1440,
                Bottom = 1440,
                Left = 1440
            };
            sectionProps.AppendChild(pageMargin);
            body.AppendChild(sectionProps);

            // "Lista de Autores:" header paragraph - Bold, Aptos Display
            Paragraph titleParagraph = new Paragraph();
            Run titleRun = new Run(new Text("Lista de Autores:"));
            RunProperties titleRunProps = new RunProperties(
                new Bold(),
                new RunFonts() { Ascii = "Aptos Display", HighAnsi = "Aptos Display" },
                new FontSize() { Val = "22" }  // 11pt = 22 half-points
            );
            titleRun.PrependChild(titleRunProps);
            titleParagraph.AppendChild(titleRun);
            body.AppendChild(titleParagraph);

            // No spacer paragraph here anymore - removed the blank line

            // Add each author
            foreach (var author in uniqueAuthors)
            {
                Paragraph authorParagraph = new Paragraph();
                ParagraphProperties pPr = new ParagraphProperties(
                    new SpacingBetweenLines() { After = "0" }
                );
                authorParagraph.AppendChild(pPr);

                // Author name in bold
                Run nameRun = new Run(new Text(author.Nome));
                RunProperties nameRunProps = new RunProperties(
                    new Bold(),
                    new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                    new FontSize() { Val = "22" }  // 11pt = 22 half-points
                );
                nameRun.PrependChild(nameRunProps);
                authorParagraph.AppendChild(nameRun);

                // Comma with space after it
                Run commaRun = new Run(new Text(", "));
                commaRun.PrependChild(new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                    new FontSize() { Val = "22" }
                ));
                authorParagraph.AppendChild(commaRun);

                // Institution without bold
                Run schoolRun = new Run(new Text(author.Escola));
                schoolRun.PrependChild(new RunProperties(
                    new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                    new FontSize() { Val = "22" }
                ));
                authorParagraph.AppendChild(schoolRun);

                body.AppendChild(authorParagraph);
            }
        }


        private DocumentData ExtractAuthorData(string file)
        {
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
                        string text = paragraphs[i].InnerText.Trim().ToLower();
                        if (text.StartsWith("resumo") || text.StartsWith("abstract"))
                        {
                            abstractIndex = i;
                            break;
                        }
                    }

                    if (abstractIndex < 0) abstractIndex = paragraphs.Count;

                    // Find the title (first paragraph)
                    if (paragraphs.Count > 0)
                    {
                        documentData.Paragraphs.Add(paragraphs[0].InnerText.Trim());
                    }

                    // Extract author information from paragraphs between the title and abstract
                    if (paragraphs.Count > 1)
                    {
                        int startIndex = 1; // Start from second line

                        // Process paragraphs and look for author info pattern
                        List<Author> authors = ExtractAuthorsFromParagraphs(paragraphs, startIndex, abstractIndex);
                        documentData.Authors.AddRange(authors);
                    }
                }
            }

            return documentData;
        }

        private List<Author> ExtractAuthorsFromParagraphs(List<Paragraph> paragraphs, int startIndex, int endIndex)
        {
            List<Author> authors = new List<Author>();
            int currentLine = startIndex;

            while (currentLine < endIndex)
            {
                // Need at least 3 lines for an author (name, email, school)
                if (currentLine + 2 >= endIndex) break;

                string name = paragraphs[currentLine].InnerText.Trim();
                string email = paragraphs[currentLine + 1].InnerText.Trim();
                string school = paragraphs[currentLine + 2].InnerText.Trim();

                // Check if we have a valid author pattern
                if (!string.IsNullOrEmpty(name) &&
                    !string.IsNullOrEmpty(email) &&
                    email.Contains("@") &&
                    !string.IsNullOrEmpty(school))
                {
                    Author author = new Author
                    {
                        Nome = name,
                        Email = email,
                        Escola = school,
                        Id = ""
                    };

                    // Check if there's an optional ID line
                    if (currentLine + 3 < endIndex)
                    {
                        string potentialId = paragraphs[currentLine + 3].InnerText.Trim();
                        // Check if the potential ID line matches ID pattern (e.g., 0009-0003-4042-1897)
                        if (Regex.IsMatch(potentialId, @"^\d{4}-\d{4}-\d{4}-\d{4}$") ||
                            Regex.IsMatch(potentialId, @"^\d+-\d+-\d+-\d+$"))
                        {
                            author.Id = potentialId;
                            currentLine += 4; // Move past all 4 lines (name, email, school, id)
                        }
                        else
                        {
                            currentLine += 3; // Move past 3 lines (name, email, school)
                        }
                    }
                    else
                    {
                        currentLine += 3; // Move past 3 lines (no ID)
                    }

                    authors.Add(author);
                }
                else
                {
                    // If the pattern doesn't match, move to the next line
                    currentLine++;
                }
            }

            return authors;
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