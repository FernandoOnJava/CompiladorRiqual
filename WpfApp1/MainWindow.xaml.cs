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
using System.Windows.Input;
using System.Collections.Generic;

namespace WpfDocCompiler
{
    public partial class MainWindow : Window, IDropTarget
    {
        public ObservableCollection<string> SelectedFiles { get; set; }
        private string editorialFilePath;
        private string editorialContent;

        public MainWindow()
        {
            InitializeComponent();
            SelectedFiles = new ObservableCollection<string>();
            filesListBox.ItemsSource = SelectedFiles;
            DataContext = this;

            // Update initial status
            UpdateStatus("Select and order your articles, then proceed to Editorial");
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

        private void EditorialStep_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("Please add at least one article file before proceeding to the Editorial step.", 
                    "No Articles Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Show the Editorial Form
            EditorialForm editForm = new EditorialForm();
            
            // If we already have editorial content, pass it to the form
            if (!string.IsNullOrWhiteSpace(editorialContent))
            {
                editForm.SetContent(editorialContent);
            }
            
            bool? result = editForm.ShowDialog();
            
            if (result == true)
            {
                // Store the editorial content and file path
                editorialContent = editForm.EditorialContent;
                editorialFilePath = editForm.EditorialFilePath;
                
                UpdateStatus("Editorial added. Ready to compile document.");
                
                // Enable the compile button
                btnCompile.IsEnabled = true;
            }
        }

        private void Compile_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("No articles to compile.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(editorialContent))
            {
                MessageBoxResult mbr = MessageBox.Show(
                    "You haven't uploaded an editorial. Do you want to proceed anyway?",
                    "No Editorial",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (mbr == MessageBoxResult.No)
                {
                    UpdateStatus("Compilation cancelled - no Editorial");
                    return;
                }
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
                    // Create the document with all components in the right order
                    CreateDocumentWithAllComponents(saveFileDialog.FileName);
                    
                    UpdateStatus($"Compilation completed with {SelectedFiles.Count} articles");
                    
                    MessageBox.Show($"Document saved successfully to: {saveFileDialog.FileName}", 
                        "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    UpdateStatus("Error during compilation: " + ex.Message);
                    MessageBox.Show("Error compiling document: " + ex.Message, 
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateStatus(string message)
        {
            statusTextBlock.Text = message;
        }

        private void CreateDocumentWithAllComponents(string outputPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // 1. Add Author List from all articles
                var authors = ExtractAuthorsFromDocs(SelectedFiles.ToList());
                AddAuthorList(body, authors);

                // 2. Add Index with Editorial and Articles
                AddIndex(body, editorialFilePath, SelectedFiles);

                // 3. Add Editorial after the index
                if (!string.IsNullOrWhiteSpace(editorialContent))
                {
                    AddEditorialPage(body, editorialContent);
                }

                // 4. Add Articles in the selected order
                foreach (var article in SelectedFiles)
                {
                    if (File.Exists(article))
                    {
                        // Add page break before each article
                        body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                        
                        // Add file name header
                        Paragraph fileNamePara = new Paragraph(
                            new Run(new Text($"Article: {Path.GetFileNameWithoutExtension(article)}")))
                        {
                            ParagraphProperties = new ParagraphProperties
                            {
                                ParagraphStyleId = new ParagraphStyleId { Val = "Heading1" }
                            }
                        };
                        body.AppendChild(fileNamePara);

                        // Add article content
                        string content = ExtractTextFromWord(article);
                        Paragraph contentPara = new Paragraph(new Run(new Text(content)));
                        body.AppendChild(contentPara);
                    }
                }

                mainPart.Document.Save();
            }
        }

        private List<Author> ExtractAuthorsFromDocs(List<string> files)
        {
            var authorList = new List<Author>();

            foreach (string file in files)
            {
                // Process only .docx files
                if (!Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
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
                                    // Try to parse author
                                    Author author = ParseAuthor(text);
                                    if (author != null)
                                    {
                                        authorList.Add(author);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // If there's an error with one file, continue with others
                    continue;
                }
            }

            return authorList;
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

        private void AddAuthorList(Body body, List<Author> authors)
        {
            // Add title for author list
            Paragraph titleParagraph = new Paragraph(
                new Run(new Text("Author List")))
            {
                ParagraphProperties = new ParagraphProperties
                {
                    ParagraphStyleId = new ParagraphStyleId { Val = "Heading1" }
                }
            };
            body.AppendChild(titleParagraph);

            // Add each author
            foreach (var author in authors)
            {
                string authorText = $"{author.Nome}";
                if (!string.IsNullOrEmpty(author.Email))
                    authorText += $" - {author.Email}";
                if (!string.IsNullOrEmpty(author.Escola))
                    authorText += $" - {author.Escola}";

                Paragraph authorParagraph = new Paragraph(new Run(new Text(authorText)));
                body.AppendChild(authorParagraph);
            }

            // Add page break after author list
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddIndex(Body body, string editorialPath, IEnumerable<string> articles)
        {
            // Add title for index
            Paragraph titleParagraph = new Paragraph(
                new Run(new Text("Index")))
            {
                ParagraphProperties = new ParagraphProperties
                {
                    ParagraphStyleId = new ParagraphStyleId { Val = "Heading1" }
                }
            };
            body.AppendChild(titleParagraph);

            // Add editorial to index
            if (!string.IsNullOrEmpty(editorialPath))
            {
                Paragraph editorialParagraph = new Paragraph(
                    new Run(new Text("Editorial")))
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = "Heading2" }
                    }
                };
                body.AppendChild(editorialParagraph);
            }

            // Add each article to index
            Paragraph articlesTitleParagraph = new Paragraph(
                new Run(new Text("Articles")))
            {
                ParagraphProperties = new ParagraphProperties
                {
                    ParagraphStyleId = new ParagraphStyleId { Val = "Heading2" }
                }
            };
            body.AppendChild(articlesTitleParagraph);

            int index = 1;
            foreach (var article in articles)
            {
                string fileName = Path.GetFileNameWithoutExtension(article);
                Paragraph indexParagraph = new Paragraph(new Run(new Text($"{index}. {fileName}")));
                body.AppendChild(indexParagraph);
                index++;
            }

            // Add page break after index
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddEditorialPage(Body body, string editorialContent)
        {
            // Add title for editorial
            Paragraph titleParagraph = new Paragraph(
                new Run(new Text("Editorial")))
            {
                ParagraphProperties = new ParagraphProperties
                {
                    ParagraphStyleId = new ParagraphStyleId { Val = "Heading1" }
                }
            };
            body.AppendChild(titleParagraph);

            // Add editorial content
            Paragraph contentParagraph = new Paragraph(new Run(new Text(editorialContent)));
            body.AppendChild(contentParagraph);

            // Add page break after editorial
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private string ExtractTextFromWord(string filePath)
        {
            // Check if it's a Word document
            if (!Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                // If not a Word document, return content as text
                return File.ReadAllText(filePath);
            }

            // Extract text from a Word document
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                if (wordDoc.MainDocumentPart != null)
                {
                    return wordDoc.MainDocumentPart.Document.Body.InnerText;
                }
            }
            return string.Empty;
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

    public class Article
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public string FilePath { get; set; }
    }


    // Converters for the UI display
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
                string extension = Path.GetExtension(filePath).ToLower();
                if (extension == ".docx")
                {
                    return new BitmapImage(new Uri("pack://application:,,,/Images/word_icon.png"));
                }
                else if (extension == ".txt")
                {
                    return new BitmapImage(new Uri("pack://application:,,,/Images/text_icon.png"));
                }
                else
                {
                    return new BitmapImage(new Uri("pack://application:,,,/Images/file_icon.png"));
                }
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}