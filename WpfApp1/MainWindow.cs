public partial class MainWindow : Window, IDropTarget
{
    // Remove duplications of the `IComponentConnector` implementation
    // Ensure `_contentLoaded` and `InitializeComponent` are defined only once
    private bool _contentLoaded;

    public void InitializeComponent()
    {
        if (_contentLoaded)
        {
            return;
        }
        _contentLoaded = true;
        System.Uri resourceLocater = new System.Uri("/WpfDocCompiler;component/mainwindow.xaml", System.UriKind.Relative);
        System.Windows.Application.LoadComponent(this, resourceLocater);
    }

    // Remove any duplicate `IComponentConnector.Connect` implementations
    void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target)
    {
        switch (connectionId)
        {
            case 1:
                this.filesListBox = (System.Windows.Controls.ListBox)(target);
                return;
            case 2:
                this.btnAddFiles = (System.Windows.Controls.Button)(target);
                return;
            case 3:
                this.btnRemoveFile = (System.Windows.Controls.Button)(target);
                return;
            case 4:
                this.btnMoveUp = (System.Windows.Controls.Button)(target);
                return;
            case 5:
                this.btnMoveDown = (System.Windows.Controls.Button)(target);
                return;
            case 6:
                this.btnCompile = (System.Windows.Controls.Button)(target);
                return;
            case 7:
                this.statusTextBlock = (System.Windows.Controls.TextBlock)(target);
                return;
        }
        this._contentLoaded = true;
    }
}
