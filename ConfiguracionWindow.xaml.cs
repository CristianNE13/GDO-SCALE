using System;
using System.Drawing.Printing;
using System.IO;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using Scale_Program.Functions;

namespace Scale_Program
{
    public partial class ConfiguracionWindow : Window
    {
        private readonly BasculaFunc bascula1 = new BasculaFunc();
        private readonly BasculaFunc bascula2 = new BasculaFunc();
        private readonly PuertosFunc ports = new PuertosFunc();


        public ConfiguracionWindow()
        {
            InitializeComponent();

            InicializarComboBox(15, 15);

            InicializarPuertos();
        }

        public static string RutaArchivo => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "configuration.xml");

        private void ConfiguracionWindow1_Loaded(object sender, RoutedEventArgs e)
        {
            CargarConfiguracion();
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
            bascula1.EnviarComandoABascula("ZRO");
        }

        private void btnZero_2_Click(object sender, RoutedEventArgs e)
        {
            bascula2.EnviarComandoABascula("ZRO");
        }

        private void InicializarComboBox(int maxInputs, int maxOutputs)
        {
            LlenarComboBox(cboxOutputUnitaria, maxOutputs);
            LlenarComboBox(cboxOutputProp65, maxOutputs);
            LlenarComboBox(cboxInputUnitaria, maxInputs);
            LlenarComboBox(cboxInputProp65, maxInputs);
            LlenarComboBox(cboxInputMaster, maxInputs);
            LlenarComboBox(cboxOutputMaster, maxOutputs);
            LlenarComboBox(cboxOutputSelladora, maxInputs);
            LlenarComboBox(cboxInputSelladora, maxInputs);

            cboxOutputUnitaria.SelectedIndex = 0;
            cboxOutputProp65.SelectedIndex = 1;
            cboxOutputProp65.SelectedIndex = 2;
            cboxInputUnitaria.SelectedIndex = 0;
            cboxInputProp65.SelectedIndex = 1;
            cboxInputMaster.SelectedIndex = 2;
            cboxOutputSelladora.SelectedIndex = 3;
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
            AgregarPuertosAComboBox(cboxScalePort_2);
            AgregarPuertosAComboBox(cboxPortSeaLevel);

            if (puertos.Count > 0) cboxScalePort_1.SelectedIndex = 0;

            if (puertos.Count > 1)
            {
                cboxScalePort_2.SelectedIndex = 1;
                cboxPortSeaLevel.SelectedIndex = 1;
            }
        }

        private void MostrarVentanaDeError(string titulo, string mensaje)
        {
            var ventanaDeError = new ErrorMessageWindow
            {
                Owner = this
            };

            ventanaDeError.lblTitle.Content = titulo;
            ventanaDeError.txtMessage.Text = mensaje;

            ventanaDeError.ShowDialog();
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
                    PuertoBascula2 = cboxScalePort_2.Text,
                    BaudRateBascula12 = 9600,
                    ParityBascula12 = Parity.None.ToString(),
                    StopBitsBascula12 = StopBits.One.ToString(),
                    DataBitsBascula12 = 8,
                    EntradaSensorUnitaria = cboxInputUnitaria.SelectedIndex,
                    EntradaSensorProp65 = cboxInputProp65.SelectedIndex,
                    EntradaSensorMaster = cboxInputMaster.SelectedIndex,
                    EntradaSensorSelladora = cboxInputSelladora.SelectedIndex,
                    SalidaDispensadoraMaster = cboxOutputMaster.SelectedIndex,
                    SalidaDispensadoraUnitaria = cboxOutputUnitaria.SelectedIndex,
                    SalidaDispensadoraProp65 = cboxOutputProp65.SelectedIndex,
                    SalidaSelladora = cboxOutputSelladora.SelectedIndex,
                    PuertoSealevel = cboxPortSeaLevel.Text,
                    BaudRateSea = int.TryParse(cboxBaudSea.Text, out var baudRate) ? baudRate : 9600,
                    ZebraName = txbZebra.Text
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
                        cboxScalePort_2.Text = configuracion.PuertoBascula2;

                        cboxInputUnitaria.SelectedIndex = configuracion.EntradaSensorUnitaria;
                        cboxInputProp65.SelectedIndex = configuracion.EntradaSensorProp65;
                        cboxInputMaster.SelectedIndex = configuracion.EntradaSensorMaster;
                        cboxInputSelladora.SelectedIndex = configuracion.EntradaSensorSelladora;

                        cboxOutputProp65.SelectedIndex = configuracion.SalidaDispensadoraProp65;
                        cboxOutputUnitaria.SelectedIndex = configuracion.SalidaDispensadoraUnitaria;
                        cboxOutputMaster.SelectedIndex = configuracion.SalidaDispensadoraMaster;
                        cboxOutputSelladora.SelectedIndex = configuracion.SalidaSelladora;

                        cboxPortSeaLevel.Text = configuracion.PuertoSealevel;
                        txbZebra.Text = configuracion.ZebraName;

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
                    bascula1.AsignarPuertoBascula(new SerialPort(cboxScalePort_1.Text, 9600, Parity.None, 8,
                        StopBits.One));
                    bascula1.OpenPort();
                    bascula1.AsignarControles(Dispatcher);
                    bascula1.OnDataReady += Bascula1_OnDataReady;
                }

                if (cboxScalePort_2.SelectedIndex >= 0)
                {
                    bascula2.AsignarPuertoBascula(new SerialPort(cboxScalePort_2.Text, 9600, Parity.None, 8,
                        StopBits.One));
                    bascula2.OpenPort();
                    bascula2.AsignarControles(Dispatcher);
                    bascula2.OnDataReady += Bascula2_OnDataReady;
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
            bascula1.ClosePort();
            bascula2.ClosePort();
        }

        private void Bascula1_OnDataReady(object sender, BasculaEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                var weight = e.Value;
                var isStable = e.IsStable;
                listLogEntries_1.Items.Add($"Peso: {weight} kg - Estable: {isStable}");
            });
        }

        private void Bascula2_OnDataReady(object sender, BasculaEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                var weight = e.Value;
                var isStable = e.IsStable;
                listLogEntries_2.Items.Add($"Peso: {weight} kg - Estable: {isStable}");
            });
        }

        private void cboxUnidades_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
                if (e.AddedItems[0] is ComboBoxItem nuevoItem)
                {
                    var nuevaSeleccion = nuevoItem.Content.ToString();

                    if (bascula1.GetPuerto() != null && bascula1.GetPuerto().IsOpen)
                    {
                        if (nuevaSeleccion == "lb")
                            bascula1.EnviarComandoABascula("UNP");
                        else if (nuevaSeleccion == "kg")
                            bascula1.EnviarComandoABascula("UNS");
                    }
                }
        }

        private void cboxUnidades2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)

                if (e.AddedItems[0] is ComboBoxItem nuevoItem)
                {
                    var nuevaSeleccion = nuevoItem.Content.ToString();

                    if (bascula2.GetPuerto() != null && bascula2.GetPuerto().IsOpen)
                    {
                        if (nuevaSeleccion == "lb")
                            bascula2.EnviarComandoABascula("UNP");
                        else if (nuevaSeleccion == "kg")
                            bascula2.EnviarComandoABascula("UNS");
                    }
                }
        }
    }
}