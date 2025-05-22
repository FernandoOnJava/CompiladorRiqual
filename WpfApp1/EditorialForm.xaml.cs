using System;
using System.Windows;

namespace WpfDocCompiler
{
    public partial class EditorialForm : Window
    {
        public string EditorialContent { get; private set; }

        public EditorialForm()
        {
            InitializeComponent();
        }

        // Set content if editing an existing editorial
        public void SetContent(string content)
        {
            if (!string.IsNullOrEmpty(content))
            {
                editorialTextBox.Text = content;
            }
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            // Save the editorial content
            EditorialContent = editorialTextBox.Text;

            // Open the MainWindow with the editorial content
            MainWindow mainWindow = new MainWindow(EditorialContent);
            mainWindow.Show();

            // Set DialogResult to true to indicate successful completion
            DialogResult = true;

            // Close this window
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // Set DialogResult to false to indicate cancellation
            DialogResult = false;

            // Close this window
            Close();
        }
    }
}