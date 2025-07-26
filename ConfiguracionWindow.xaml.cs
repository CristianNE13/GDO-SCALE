using System;
using System.Drawing.Printing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using Scale_Program.Functions;

namespace Scale_Program
{
    public partial class ConfiguracionWindow : Window
    {
        private readonly int MinLenght = 5;
        private readonly PuertosFunc ports = new PuertosFunc();
        private IBasculaFunc bascula1;

        public ConfiguracionWindow()
        {
            InitializeComponent();

            InicializarComboBox(15, 15);
        }

        public static string RutaArchivo => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "configuration.xml");

        private void ConfiguracionWindow1_Loaded(object sender, RoutedEventArgs e)
        {
            CargarConfiguracion();

            InicializarPuertos();
        }

        private async void OnSaveButtonClick(object sender, RoutedEventArgs e)
        {
            GuardarConfiguracion();

            await Task.Delay(100);

            CargarConfiguracion();
        }

        private void OnCancelButtonClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btnZero_1_Click(object sender, RoutedEventArgs e)
        {
            bascula1.EnviarZero();
        }

        private void InicializarComboBox(int maxInputs, int maxOutputs)
        {
            if (FindName("cboxInputPick2L0") is ComboBox pick1)
                LlenarComboBox(pick1, maxInputs);

            if (FindName("cboxOutputPick2L0") is ComboBox pick2)
                LlenarComboBox(pick2, maxOutputs);

            if (FindName("cboxOutputPiston") is ComboBox piston)
                LlenarComboBox(piston, maxOutputs);

            if (FindName("cboxOutputShutOff") is ComboBox shutoff)
                LlenarComboBox(shutoff, maxOutputs);

            if (FindName("cboxInputBoton") is ComboBox boton)
                LlenarComboBox(boton, maxInputs);

            if (FindName("cboxInputSensor0") is ComboBox sensor0)
                LlenarComboBox(sensor0, maxInputs);

            if (FindName("cboxInputSensor1") is ComboBox sensor1)
                LlenarComboBox(sensor1, maxInputs);
        }

        private void LlenarComboBox(ComboBox comboBox, int maxItems)
        {
            comboBox.Items.Clear();
            for (var i = 0; i <= maxItems; i++)
                comboBox.Items.Add(new ComboBoxItem
                {
                    Content = i.ToString()
                });
        }

        private void InicializarPuertos()
        {
            var puertos = ports.GetPorts();
            if (puertos == null || puertos.Count == 0)
            {
                MostrarVentanaDeError("Advertencia", "No se encontraron puertos disponibles.");
                return;
            }

            AgregarPuertosAComboBox(cboxScalePort_1);
            AgregarPuertosAComboBox(cboxPortSeaLevel);

            if (puertos.Count > 0) cboxScalePort_1.SelectedIndex = 0;

            if (puertos.Count > 1)
                cboxPortSeaLevel.SelectedIndex = 1;
        }

        private void MostrarVentanaDeError(string titulo, string mensaje)
        {
            ErrorMessageWindow ventanaDeError = new ErrorMessageWindow();
            ventanaDeError.TitleText = titulo;
            ventanaDeError.Message = mensaje;
            ventanaDeError.ShowDialog();
            ventanaDeError.Focus();
        }

        private void AgregarPuertosAComboBox(ComboBox comboBox)
        {
            var puertos = ports.GetPorts();
            foreach (var puerto in puertos)
                comboBox.Items.Add(new ComboBoxItem
                {
                    Content = puerto
                });
        }

        private void btPrintTest_Click(object sender, RoutedEventArgs e)
        {
            var printer = new PrintDocument();

            printer.PrinterSettings.PrinterName = txbZebra.Text;

            try
            {
                var (zpl, integer, fractional) = ZebraPrinter.GenerateZplBody("10");

                RawPrinterHelper.SendStringToPrinter(printer.PrinterSettings.PrinterName, zpl);

                MessageBox.Show("Etiqueta enviada correctamente.", "Impresión exitosa", MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MostrarError($"Error al imprimir: {ex.Message}");
            }
        }

        private void GuardarConfiguracion()
        {
            try
            {
                var filePath = RutaArchivo;

                var configuracion = new Configuracion
                {
                    PuertoBascula1 = cboxScalePort_1.Text,
                    BaudRateBascula12 = 9600,
                    ParityBascula12 = Parity.None.ToString(),
                    StopBitsBascula12 = StopBits.One.ToString(),
                    DataBitsBascula12 = 8,
                    User = txbUser.Text,
                    InputPick2L0 = cboxInputPick2L0.SelectedIndex,
                    OutputPick2L0 = cboxOutputPick2L0.SelectedIndex,
                    InputBoton = cboxInputBoton.SelectedIndex,
                    CheckShutOff = chkBox_Shutoff.IsChecked ?? false,
                    ShutOff = cboxOutputShutOff.SelectedIndex,
                    Piston = cboxOutputPiston.SelectedIndex,
                    PuertoSealevel = cboxPortSeaLevel.Text,
                    BaudRateSea = int.TryParse(cboxBaudSea.Text, out var baudRate) ? baudRate : 9600,
                    IpCamara = txbIPCamara.Text,
                    PuertoCamara = int.TryParse(txbPuerto.Text, out var puertoCamara) ? puertoCamara : 0,
                    BasculaMarca = cboxMarca.Text,
                    CheckSealevelEthernet = chkBox_SeaEthernet.IsChecked ?? false,
                    SealevelIP = txbIPSealevel.Text,
                    InputSensor0 = cboxInputSensor0.SelectedIndex,
                    InputSensor1 = cboxInputSensor1.SelectedIndex
                };

                var serializer = new XmlSerializer(typeof(Configuracion));
                using (var writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, configuracion);
                }

                MessageBox.Show("Configuración guardada correctamente.", "Éxito", MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MostrarError($"Error al guardar configuración: {ex.Message}");
            }
        }

        private void CargarConfiguracion()
        {
            try
            {
                var filePath = RutaArchivo;

                if (!File.Exists(filePath))
                    return;

                var serializer = new XmlSerializer(typeof(Configuracion));
                using (var reader = new StreamReader(filePath))
                {
                    var configuracion = (Configuracion)serializer.Deserialize(reader);

                    if (configuracion != null)
                    {
                        cboxScalePort_1.Text = configuracion.PuertoBascula1;
                        cboxPortSeaLevel.Text = configuracion.PuertoSealevel;

                        cboxInputPick2L0.Text = configuracion.InputPick2L0.ToString();
                        cboxInputBoton.Text = configuracion.InputBoton.ToString();
                        cboxInputSensor0.Text = configuracion.InputSensor0.ToString();
                        cboxInputSensor1.Text = configuracion.InputSensor1.ToString();
                        cboxOutputShutOff.Text = configuracion.ShutOff.ToString();
                        chkBox_Shutoff.IsChecked = configuracion.CheckShutOff;
                        cboxOutputPiston.Text = configuracion.Piston.ToString();

                        cboxOutputPick2L0.Text = configuracion.OutputPick2L0.ToString();

                        txbUser.Text = configuracion.User;
                        txbIPCamara.Text = configuracion.IpCamara;
                        txbPuerto.Text = configuracion.PuertoCamara.ToString();

                        chkBox_SeaEthernet.IsChecked = configuracion.CheckSealevelEthernet;
                        txbIPSealevel.Text = configuracion.SealevelIP;

                        if (configuracion.BasculaMarca == "Pennsylvania")
                            cboxMarca.SelectedIndex = 1;
                        else
                            cboxMarca.SelectedIndex = 0;

                        foreach (ComboBoxItem item in cboxBaudSea.Items)
                            if (item.Content.ToString() == configuracion.BaudRateSea.ToString())
                            {
                                cboxBaudSea.SelectedItem = item;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                MostrarError($"Error al cargar configuración: {ex.Message}");
            }
        }

        private void MostrarError(string error)
        {
            var _error = new ErrorMessageWindow();
            _error.TitleText = "ERROR";
            _error.Message = error;
            _error.Show();
        }

        private void btnConectar_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (cboxScalePort_1.SelectedIndex >= 0)
                {
                    if (bascula1 != null && bascula1.GetPuerto() != null && bascula1.GetPuerto().IsOpen)
                        bascula1.ClosePort();

                    if (cboxMarca.Text == "Pennsylvania")
                        bascula1 = new BasculaFuncPennsylvania();
                    else
                        bascula1 = new BasculaFuncGFC();

                    bascula1.AsignarPuertoBascula(new SerialPort(cboxScalePort_1.Text, 9600, Parity.None, 8));
                    bascula1.OpenPort();
                    bascula1.AsignarControles(Dispatcher);
                    bascula1.OnDataReady += Bascula1_OnDataReady;
                }

                MessageBox.Show("Conexión exitosa a las básculas.", "Éxito", MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MostrarError($"Error al conectar con las básculas: {ex.Message}");
            }
        }

        private void ConfiguracionWindow1_Closed(object sender, EventArgs e)
        {
            try
            {
                if (bascula1 != null && bascula1.GetPuerto() != null && bascula1.GetPuerto().IsOpen)
                    bascula1.ClosePort();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
                throw;
            }
        }

        private void Bascula1_OnDataReady(object sender, BasculaEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                var weight = e.Value;
                var isStable = e.IsStable;
                listLogEntries_1.Items.Add($"Peso: {weight} kg - Estable: {isStable}");
                if (listLogEntries_1.Items.Count > 16)
                    listLogEntries_1.Items.Clear();
            });
        }

        private void cboxUnidades_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
                if (e.AddedItems[0] is ComboBoxItem nuevoItem)
                {
                    var nuevaSeleccion = nuevoItem.Content.ToString();

                    if (bascula1 != null && bascula1.GetPuerto() != null && bascula1.GetPuerto().IsOpen)
                    {
                        if (nuevaSeleccion == "lb")
                            bascula1.EnviarComandoABascula("UNP");
                        else if (nuevaSeleccion == "kg")
                            bascula1.EnviarComandoABascula("UNS");
                    }
                }
        }


        private void btnChangePassword_Click(object sender, RoutedEventArgs e)
        {
            ChangePassword();
        }

        private void ChangePassword()
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var usuario = db.users.FirstOrDefault(s => s.username == txbUser.Text);

                    if (PasswordHash.VerifyPassword(txtPassword.Password, usuario.password))
                    {
                        if (txtNewPassword.Password.Length < MinLenght)
                        {
                            MessageBox.Show($"La contraseña debe tener al menos {MinLenght} caracteres.", "Error");
                            txtNewPassword.Focus();
                        }
                        else if (txtPassword.Password == txtNewPassword.Password)
                        {
                            MessageBox.Show("La nueva contraseña no debe ser igual a la anterior.", "Error");
                            txtPassword.Focus();
                        }
                        else if (txtNewPassword.Password != txtPasswordConfirmation.Password)
                        {
                            MessageBox.Show("La nueva contraseña no coincide con la confirmación.", "Error");
                            txtPasswordConfirmation.Focus();
                        }
                        else
                        {
                            usuario.password = PasswordHash.CreateHash(txtNewPassword.Password);
                            db.SaveChanges();

                            lblSuccess.Visibility = Visibility.Visible;

                            txtPassword.Clear();
                            txtNewPassword.Clear();
                            txtPasswordConfirmation.Clear();
                            txtPassword.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Contraseña actual incorrecta.", "Error");
                        txtPassword.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Error");
            }
        }

        private async void btnProbarCamara_Click(object sender, RoutedEventArgs e)
        {
            KeyenceTcpClient keyence = null;

            try
            {
                lblResultadoVision.Content = "";
                keyence = new KeyenceTcpClient(txbIPCamara.Text, int.Parse(txbPuerto.Text));

                var connected = await keyence.ConnectAsync();
                if (!connected)
                {
                    MessageBox.Show("No se pudo conectar a la cámara.", "Error");
                    return;
                }

                var response = await keyence.SendTrigger();
                var resultado = keyence.Formato(response);
                foreach (var text in resultado)
                    lblResultadoVision.Content += text + "\n";
            }
            catch (Exception error)
            {
                MessageBox.Show($"Error: {error.Message}", "Error");
            }
            finally
            {
                if (keyence != null)
                    keyence.Dispose();
            }
        }


        private async void btnCambiarPrograma_Click(object sender, RoutedEventArgs e)
        {
            KeyenceTcpClient keyence = null;

            try
            {
                lblResultadoVision.Content = "";
                keyence = new KeyenceTcpClient(txbIPCamara.Text, int.Parse(txbPuerto.Text));

                var connected = await keyence.ConnectAsync();
                if (!connected)
                {
                    MessageBox.Show("No se pudo conectar a la cámara.", "Error");
                    return;
                }

                var programa = int.Parse(txbPrograma.Text);
                var response = await keyence.ChangeProgram(programa);
                var resultado = keyence.Formato(response);
                foreach (var text in resultado)
                    lblResultadoVision.Content += text + "\n";
            }
            catch (Exception error)
            {
                MessageBox.Show($"Error: {error.Message}", "Error");
            }
            finally
            {
                if (keyence != null)
                    keyence.Dispose();
            }
        }

        private void chkBox_Shutoff_Checked(object sender, RoutedEventArgs e)
        {
            lblShutOff.Visibility = Visibility.Visible;
            cboxOutputShutOff.Visibility = Visibility.Visible;
        }

        private void chkBox_Shutoff_Unchecked(object sender, RoutedEventArgs e)
        {
            lblShutOff.Visibility = Visibility.Hidden;
            cboxOutputShutOff.Visibility = Visibility.Hidden;
        }

        private void chkBox_SeaEthernet_Checked(object sender, RoutedEventArgs e)
        {
            lblPuertoES.Content = "IP Sealevel:";
            txbIPSealevel.Visibility = Visibility.Visible;
            lblBaudSea.Visibility = Visibility.Hidden;
            cboxBaudSea.Visibility = Visibility.Hidden;
            cboxPortSeaLevel.Visibility = Visibility.Hidden;
        }

        private void chkBox_SeaEthernet_Unchecked(object sender, RoutedEventArgs e)
        {
            lblPuertoES.Content = "Puerto Sealevel:";
            txbIPSealevel.Visibility = Visibility.Hidden;
            lblBaudSea.Visibility = Visibility.Visible;
            cboxBaudSea.Visibility = Visibility.Visible;
            cboxPortSeaLevel.Visibility = Visibility.Visible;
        }


        public static class PasswordHash
        {
            private const int SaltSize = 16;
            private const int KeySize = 32;
            private const int Iterations = 10000;

            public static string CreateHash(string password)
            {
                using (var rng = new RNGCryptoServiceProvider())
                {
                    var salt = new byte[SaltSize];
                    rng.GetBytes(salt);

                    using (var pbkdf2 = new Rfc2898DeriveBytes(password, salt, Iterations))
                    {
                        var key = pbkdf2.GetBytes(KeySize);
                        var hashBytes = new byte[SaltSize + KeySize];
                        Array.Copy(salt, 0, hashBytes, 0, SaltSize);
                        Array.Copy(key, 0, hashBytes, SaltSize, KeySize);
                        return Convert.ToBase64String(hashBytes);
                    }
                }
            }

            public static bool VerifyPassword(string password, string hashedPassword)
            {
                var hashBytes = Convert.FromBase64String(hashedPassword);

                var salt = new byte[SaltSize];
                Array.Copy(hashBytes, 0, salt, 0, SaltSize);

                using (var pbkdf2 = new Rfc2898DeriveBytes(password, salt, Iterations))
                {
                    var key = pbkdf2.GetBytes(KeySize);

                    for (var i = 0; i < KeySize; i++)
                        if (hashBytes[i + SaltSize] != key[i])
                            return false;

                    return true;
                }
            }
        }
    }
}