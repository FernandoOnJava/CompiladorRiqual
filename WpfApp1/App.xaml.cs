using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Data;
using WpfDocCompiler;

namespace WpfApp1
{
    /// <summary>
    /// Entry point class for the application
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Start with EditorialForm instead of MainWindow
            EditorialForm editorialForm = new EditorialForm();
            editorialForm.Show();
        }
    }
}