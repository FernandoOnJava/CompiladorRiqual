using System;
using System.Windows;

namespace WpfDocCompiler
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Start directly with MainWindow
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }
    }
}