using System;
using System.Windows;

namespace WpfDocCompiler
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Start with the EditorialForm
            EditorialForm editorialForm = new EditorialForm();
            editorialForm.Show();
        }
    }
}