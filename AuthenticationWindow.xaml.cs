using System;
using System.Windows;

namespace Scale_Program
{
    public partial class AuthenticationWindow : Window
    {
        private const string pass = "12345";
        
        public AuthenticationWindow()
        {
            InitializeComponent();

            this.Loaded += OnAuthenticationWindowLoaded;
        }

        private void OnAuthenticationWindowLoaded(object sender, RoutedEventArgs e)
        {
            txtPassword.Focus();
        }

        private void OnLoginButtonClick(object sender, RoutedEventArgs e)
        {
            if( AuthenticateUser() )
            {
                this.DialogResult = true;
            }
        }

        private bool AuthenticateUser()
        {
            bool succedded = false;

            try
            {
                if (txtPassword.Password == pass)
                    succedded = true;

                if(!succedded)
                    ShowErroMessage("Contraseña incorrecta.");
            }
            catch(Exception ex)
            {
                ShowErroMessage("Error.");
            }

            return succedded;
        }

        private void ShowErroMessage(string message )
        {
            ErrorMessageWindow err = new ErrorMessageWindow();

            err.TitleText = "Falla en Autenticación";
            err.Message = message;

            err.WindowStartupLocation = WindowStartupLocation.Manual;

            double w = ((err.Width - this.Width) / 2);
            double h = ((err.Height - this.Height) / 2);

            err.Left = this.Left - w;
            err.Top = this.Top - h;

            err.ShowDialog();
        }

        private void OnCancelButtonClick(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}
