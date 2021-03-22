using System.Windows;

namespace SpasticityClient
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainWindowViewModel mainWindowViewModel;
        public MainWindow()
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MjQyNDEwQDMxMzgyZTMxMmUzMGNKSzduS2FuZktxRmJERmkySTJsZjVCWlFPeVRGM3pMa1NPYWlVeUttSzA9");
            InitializeComponent();
            mainWindowViewModel = (MainWindowViewModel)this.DataContext;
        }
    }
}
