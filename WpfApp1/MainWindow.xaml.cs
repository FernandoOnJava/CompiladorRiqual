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
    // Converters continuam os mesmos...

    public partial class MainWindow : Window, IDropTarget
    {
        public ObservableCollection<string> SelectedFiles { get; set; }
        private string editorialContent; // Armazenar o conteúdo editorial recebido

        public MainWindow(string editorialContent)
        {
            InitializeComponent();
            SelectedFiles = new ObservableCollection<string>();
            filesListBox.ItemsSource = SelectedFiles;
            DataContext = this;

            // Armazenar o conteúdo editorial
            this.editorialContent = editorialContent;

            // Atualizar a barra de status para refletir o fato de que o editorial foi concluído
            UpdateStatus("Editorial concluído. Selecione os arquivos para compilar.");

            // Configurar o manipulador de arrastar e soltar
            GongSolutions.Wpf.DragDrop.DragDrop.SetDropHandler(filesListBox, this);

            // Manipular o fechamento da janela principal
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Quando o MainWindow for fechado, o aplicativo deve ser encerrado
            Application.Current.Shutdown();
        }

        #region Drag and Drop Implementation
        // Código de implementação de drag-and-drop continua o mesmo...
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

        // Método modificado para não mostrar o EditorialForm, já que temos o conteúdo
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
                UpdateStatus("Nenhum ficheiro para compilar");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word File (*.docx)|*.docx|JSON File (*.json)|*.json",
                Title = "Save Compiled File"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    string extension = Path.GetExtension(saveFileDialog.FileName).ToLower();

                    // Já temos o conteúdo editorial, não precisamos pedir novamente
                    if (string.IsNullOrWhiteSpace(editorialContent))
                    {
                        MessageBoxResult mbr = MessageBox.Show(
                            "Continuar sem Editorial?",
                            "Editorial não foi preenchido",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);

                        if (mbr == MessageBoxResult.No)
                        {
                            UpdateStatus("Compilação cancelada - sem Editorial");
                            return;
                        }
                    }

                    if (extension == ".json")
                    {
                        // Extract and save as JSON
                        var allData = ExtractDataFromDocs(SelectedFiles.ToList());
                        File.WriteAllText(saveFileDialog.FileName, JsonConvert.SerializeObject(allData, Formatting.Indented));
                        UpdateStatus($"Compilação JSON concluída com {SelectedFiles.Count} artigos");
                    }
                    else if (extension == ".docx")
                    {
                        // Create a new Word document with editorial and articles
                        CreateDocumentWithEditorialAndArticles(saveFileDialog.FileName, editorialContent, SelectedFiles);
                        UpdateStatus($"Compilação DOCX concluída com {SelectedFiles.Count} artigos");
                    }
                    else // .txt
                    {
                        // Compile as plain text
                        CompileAsText(SelectedFiles.ToList(), saveFileDialog.FileName);
                        UpdateStatus($"Compilação TXT concluída com {SelectedFiles.Count} ficheiros");
                    }

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

        // Modificado para editar o editorial
        public void EditEditorial_Click(object sender, RoutedEventArgs e)
        {
            // Criar um novo formulário editorial para editar o conteúdo existente
            EditorialForm editForm = new EditorialForm();
            editForm.Owner = this;

            // Preencher com o conteúdo atual
            editForm.SetContent(editorialContent);

            // Mostrar como diálogo
            bool? result = editForm.ShowDialog();

            // Atualizar o conteúdo se o usuário confirmou
            if (result == true)
            {
                editorialContent = editForm.EditorialContent;
                UpdateStatus("Editorial atualizado");
            }
        }

        private void UpdateStatus(string message)
        {
            statusTextBlock.Text = message;
        }

        // Resto da implementação continua a mesma...
        // Métodos como CreateDocumentWithEditorialAndArticles, CompileAsText, etc.
        // Apenas usam o campo editorialContent em vez de chamar GetEditorialContent()

        // Método para criar um documento com o editorial e artigos
        private void CreateDocumentWithEditorialAndArticles(string outputPath, string editorialContent, IEnumerable<string> articles)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Add Editorial
                if (!string.IsNullOrWhiteSpace(editorialContent))
                {
                    AddEditorialPage(body, editorialContent);
                    body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page }))); // Page break after editorial
                }

                // Add Articles
                foreach (var article in articles)
                {
                    if (File.Exists(article))
                    {
                        // Add file name header
                        Paragraph fileNamePara = new Paragraph(new Run(new Text($"File: {Path.GetFileName(article)}")));
                        body.AppendChild(fileNamePara);

                        string content = ExtractTextFromWord(article);
                        Paragraph contentPara = new Paragraph(new Run(new Text(content)));
                        body.AppendChild(contentPara);

                        // Add page break after each article
                        body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                    }
                }

                mainPart.Document.Save();
            }
        }

        // Método necessário para extrair texto de documentos Word
        private string ExtractTextFromWord(string filePath)
        {
            // Verificar se é um arquivo Word
            if (!Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                // Se não for um arquivo Word, retornar o conteúdo como texto
                return File.ReadAllText(filePath);
            }

            // Extrair texto de um documento Word
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                if (wordDoc.MainDocumentPart != null)
                {
                    return wordDoc.MainDocumentPart.Document.Body.InnerText;
                }
            }
            return string.Empty;
        }

        // Método para compilar como texto
        private void CompileAsText(List<string> files, string outputPath)
        {
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Adicionar o editorial primeiro, se existir
                if (!string.IsNullOrWhiteSpace(editorialContent))
                {
                    writer.WriteLine("EDITORIAL:");
                    writer.WriteLine(editorialContent);
                    writer.WriteLine();
                    writer.WriteLine("-------------------------------");
                    writer.WriteLine();
                }

                // Adicionar cada arquivo
                foreach (string filePath in files)
                {
                    if (File.Exists(filePath))
                    {
                        writer.WriteLine($"ARQUIVO: {Path.GetFileName(filePath)}");
                        writer.WriteLine();

                        if (Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extrair texto de arquivos Word
                            string content = ExtractTextFromWord(filePath);
                            writer.Write(content);
                        }
                        else
                        {
                            // Arquivos de texto normais
                            string content = File.ReadAllText(filePath);
                            writer.Write(content);
                        }
                        writer.WriteLine();
                        writer.WriteLine("-------------------------------");
                        writer.WriteLine();
                    }
                }
            }
        }

        // Método para adicionar uma página editorial
        private void AddEditorialPage(Body body, string editorialContent)
        {
            // "Editorial" header paragraph
            Paragraph titleParagraph = new Paragraph(new Run(new Text("Editorial")));
            body.AppendChild(titleParagraph);

            // Add the editorial content
            Paragraph contentParagraph = new Paragraph(new Run(new Text(editorialContent)));
            body.AppendChild(contentParagraph);
        }

        // Método para extrair dados de documentos
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

        // Método para analisar dados de autor
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

    // Classes Author e DocumentData permanecem as mesmas
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