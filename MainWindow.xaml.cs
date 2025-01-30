using System.Windows;
using System.Windows.Controls;

namespace APIConnectionLB
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Classes.APIInteraction apiinteraction;
        public MainWindow()
        {
            InitializeComponent();
            apiinteraction = new Classes.APIInteraction();
        }

        private void ButtonGetFullName_Click(object sender, RoutedEventArgs e)
        {
            TextBlockFullName.Text = apiinteraction.GetFullName();
        }

        private void ButtonSendResult_Click(object sender, RoutedEventArgs e)
        {
            TextBlockResult.Text = apiinteraction.FillDocument();
        }
    }
}
