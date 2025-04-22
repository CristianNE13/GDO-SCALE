using Scale_Program.Functions;
using System;
using System.Linq;
using System.Windows;

namespace Scale_Program
{
    public partial class AuthenticationWindow : Window
    {
        private dc_missingpartsEntities db;
        private Configuracion defaultSettings;

        public AuthenticationWindow()
        {
            InitializeComponent();
            db = new dc_missingpartsEntities();
            this.Loaded += OnAuthenticationWindowLoaded;
        }

        private void OnAuthenticationWindowLoaded(object sender, RoutedEventArgs e)
        {
            txtPassword.Focus();
            defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);
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
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var usuario = db.users.FirstOrDefault(u => u.username == defaultSettings.User);

                    if (usuario == null)
                    {
                        ShowErroMessage("Usuario no encontrado.");
                        return false;
                    }

                    if (ConfiguracionWindow.PasswordHash.VerifyPassword(txtPassword.Password, usuario.password))
                    {
                        return true;
                    }
                    else
                    {
                        ShowErroMessage("Contraseña incorrecta.");
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                ShowErroMessage("Error al autenticar usuario.");
                return false;
            }
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
