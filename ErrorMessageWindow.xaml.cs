using System.Windows;
using System.Windows.Input;

namespace Scale_Program
{
    /// <summary>
    ///     Interaction logic for MessageBoxWindow.xaml
    /// </summary>
    public partial class ErrorMessageWindow : Window
    {
        public ErrorMessageWindow()
        {
            InitializeComponent();
            Loaded += OnMessageBoxWindowLoaded;
            MouseDown += OnMMouseDown;
        }

        public string TitleText { get; set; }

        public string Message { get; set; }

        private void OnMMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void OnMessageBoxWindowLoaded(object sender, RoutedEventArgs e)
        {
            lblTitle.Content = TitleText;
            txtMessage.Text = Message;
        }

        private void OnAcceptButtonClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Focus();
        }
    }
}