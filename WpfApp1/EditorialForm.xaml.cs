using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WpfDocCompiler
{
    public partial class EditorialForm : Window
    {
        private TextBox editorialTextBox;
        public string EditorialContent { get; private set; }

        public EditorialForm()
        {
            // Window properties
            Width = 700;
            Height = 500;
            Title = "Editorial";
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanResize;

            // Create a grid layout
            Grid mainGrid = new Grid();
            mainGrid.Margin = new Thickness(10);

            // Define rows
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // For title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // For text box
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });  // For buttons

            // Create a title label
            TextBlock titleLabel = new TextBlock
            {
                Text = "Digite o conteúdo do Editorial:",
                FontWeight = FontWeights.Bold,
                FontSize = 14,
                Margin = new Thickness(0, 0, 0, 5)
            };
            Grid.SetRow(titleLabel, 0);
            mainGrid.Children.Add(titleLabel);

            // Create a text box with scrolling
            editorialTextBox = new TextBox
            {
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(0, 0, 0, 10),
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(5),
                FontFamily = new FontFamily("Times New Roman"),
                FontSize = 12
            };
            Grid.SetRow(editorialTextBox, 1);
            mainGrid.Children.Add(editorialTextBox);

            // Create button panel
            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetRow(buttonPanel, 2);

            // Create OK button
            Button okButton = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 30,
                Margin = new Thickness(0, 0, 10, 0),
                IsDefault = true
            };
            okButton.Click += (sender, e) =>
            {
                EditorialContent = editorialTextBox.Text;
                DialogResult = true;
                Close();
            };
            buttonPanel.Children.Add(okButton);

            // Create Cancel button
            Button cancelButton = new Button
            {
                Content = "Cancelar",
                Width = 80,
                Height = 30,
                IsCancel = true
            };
            cancelButton.Click += (sender, e) =>
            {
                DialogResult = false;
                Close();
            };
            buttonPanel.Children.Add(cancelButton);

            mainGrid.Children.Add(buttonPanel);

            // Set the content of the window
            Content = mainGrid;
        }

        // Method to preload editorial content if needed
        public void SetContent(string content)
        {
            if (editorialTextBox != null)
            {
                editorialTextBox.Text = content;
            }
        }
    }
}