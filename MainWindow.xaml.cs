using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Scale_Program.Functions;
using Brushes = System.Windows.Media.Brushes;
using Rectangle = System.Windows.Shapes.Rectangle;

namespace Scale_Program
{
    public partial class MainWindow : Window
    {
        private readonly BasculaFunc bascula;
        private readonly List<SequenceStep> valoresBolsas = new List<SequenceStep>();
        private Modelo ModeloData;
        private double _accumulatedWeight;
        private AlertWindow _alertWindow;
        private Catalogos _catalogos = new Catalogos();
        private int _consecutiveCount;
        private int _currentStepIndex;
        private ErrorMessageWindow _errorWindow;
        private bool _etapa1;
        private bool _etapa2;
        private bool _isInitializing;
        private double _lastWeight;
        private bool _stopBascula1 = true;
        private bool _validacion = true;
        private string codigo;
        private int contador;
        private Configuracion defaultSettings;
        private string fraction;
        private string integrer;
        public IOInterface ioInterface;
        public IOScanner ioScanner;
        private bool ioScannerActivado;
        private List<SequenceStep> pasosFiltrados;
        private double pieceWeight;
        private int stepIndex = 0;
        private double runningWeight = 0.0;
        private bool sensorMasterActivo;
        private bool sensorProp65Activo;
        private string zpl;
        private bool sensorSelladoraActivo;
        private bool sensorUnitariaActivo;
        private List<SequenceStep> sequence;
        private DispatcherTimer _estadoBasculasTimer;

        private readonly string rutaImagenes = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes");

        public MainWindow()
        {
            InitializeComponent();

            defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);

            bascula = new BasculaFunc();
            bascula.AsignarControles(Dispatcher);
            bascula.OnDataReady += Bascula1_OnDataReady;
        }

        private void Cbox_Modelo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);
                //IniciarSealevel();

                ModeloData = Cbox_Modelo.SelectedValue as Modelo ?? throw new Exception("Modelo no reconocido.");

                ProcesarModeloValido(ModeloData.NoModelo);

                lblProgreso.Visibility = ModeloData.UsaConteoCajas ? Visibility.Visible : Visibility.Hidden;
                lblConteoCajas.Visibility = ModeloData.UsaConteoCajas ? Visibility.Visible : Visibility.Hidden;
            }
            catch (Exception exception)
            {
                ShowAlertError($"Error al seleccionar el modelo: {exception.Message}");
            }
        }

        private void ProcesarModeloValido(string modeloSeleccionado)
        {
            CargarSecuenciasPorModelo(modeloSeleccionado);
            CargarProcesos();

            HideAll();
            Cbox_Proceso.IsEnabled = true;
            Cbox_Proceso.Visibility = Visibility.Visible;
            CerrarAlertas();
            Cbox_Proceso.Text = "";
        }

        private void Cbox_Proceso_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitializing || Cbox_Proceso.SelectedValue == null || Cbox_Modelo.SelectedValue == null)
                return;

            try
            {
                if (!int.TryParse(Cbox_Proceso.SelectedValue.ToString(), out int procesoSeleccionado))
                    throw new Exception("El valor del proceso seleccionado no es válido.");

                pasosFiltrados = ObtenerValoresProceso(ModeloData.NoModelo, procesoSeleccionado);

                if (pasosFiltrados == null || pasosFiltrados.Count == 0)
                    throw new Exception("No hay pasos definidos para este proceso.");
                ProcesarModeloYProceso();

                SecuenciaASeguir(ModeloData);

                var currentStep = pasosFiltrados[_currentStepIndex];
                ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight, pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al cargar la secuencia: {ex.Message}");
            }
        }

        private List<SequenceStep> ObtenerValoresProceso(string modeloSeleccionado, int procesoSeleccionado)
        {
            try
            {
                int modProceso = ObtenerModeloModProceso(modeloSeleccionado);
                if (modProceso == -1) return null;

                return sequence
                    .Where(s => s.PartProceso == procesoSeleccionado && s.ModProceso == modProceso)
                    .OrderBy(s => int.Parse(s.PartOrden))
                    .ToList();

            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al obtener la secuencia filtrada para '{modeloSeleccionado}': {ex.Message}");
                return null;
            }
        }

        private void ProcesarModeloYProceso()
        {
            ResetSequence(pasosFiltrados);
            HideAll();
            SetValuesEtapas(pasosFiltrados, 1);
            HideFerreteria(false, pasosFiltrados);
        }

        private void CargarSecuenciasPorModelo(string modeloSeleccionado)
        {
            try
            {
                var modeloModProceso = ObtenerModeloModProceso(modeloSeleccionado);

                List<SequenceStep> pasosFiltrados = new List<SequenceStep>();

                using (var db = new dc_missingpartsEntities())
                {
                    var articulos = db.Articulos
                        .Where(a => a.ModProceso == modeloModProceso)
                        .OrderBy(a => a.Paso)
                        .ToList();

                    pasosFiltrados.AddRange(articulos.Select(CrearSequenceStepDesdeFila).Where(step => step != null));
                }

                sequence = pasosFiltrados;
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al cargar secuencias para el modelo {modeloSeleccionado}: {ex.Message}");
            }
        }

        private int ObtenerModeloModProceso(string modeloSeleccionado)
        {
            try
            {
                modeloSeleccionado = modeloSeleccionado.ToUpper();

                using (var db = new dc_missingpartsEntities())
                {
                    var modelo = db.Modelos.FirstOrDefault(m => m.NoModelo.ToUpper() == modeloSeleccionado);
                    if (modelo != null)
                    {
                        return modelo.ModProceso;
                    }
                }

                throw new ArgumentException($"Modelo '{modeloSeleccionado}' no encontrado en la hoja 'Modelos'.");
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al obtener ModProceso para '{modeloSeleccionado}': {ex.Message}");
                return -1;
            }
        }

        private SequenceStep CrearSequenceStepDesdeFila(Articulo articulo)
        {
            return new SequenceStep
            {
                MinWeight = articulo.PesoMin,
                MaxWeight = articulo.PesoMax,
                IsCompleted = false,
                DetectedWeight = "",
                Tag = "",
                PartOrden = articulo.Paso.ToString(),
                GrdPart = $"Part{articulo.Paso - 1}",
                PartNoParte = articulo.NoParte,
                PartImagen = $"Part_Imagen{articulo.Paso - 1}",
                PartIndicator = $"Part_Indicator{articulo.Paso - 1}",
                PartPeso = $"Part_Peso{articulo.Paso - 1}",
                PartSecuencia = $"Part_Secuencia{articulo.Paso - 1}",
                PartProceso = articulo.Proceso,
                PartCantidad = articulo.Cantidad.ToString(),
                ModProceso = articulo.ModProceso,
                Descripcion = articulo.Descripcion ?? ""
            };
        }


        private void Window_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                //OutputsOff();

                if (bascula.GetPuerto() != null && bascula.GetPuerto().IsOpen)
                    bascula.ClosePort();

                CerrarAlertas();

                Environment.Exit(0);
            }
            catch (Exception exception)
            {
                ShowAlertError(exception.ToString());
                throw;
            }

        }

        private void ResetSequence(List<SequenceStep> steps)
        {
            if (ModeloData == null)
                throw new ArgumentException("No hay un modelo seleccionado.");

            var modeloSeleccionado = ModeloData.NoModelo;

            if (ModeloExiste(modeloSeleccionado))
            {
                ReiniciarPasos(steps, modeloSeleccionado);
            }
            else
            {
                ShowAlertError($"Modelo '{modeloSeleccionado}' no reconocido.");
            }
        }

        private void ReiniciarPasos(List<SequenceStep> steps, string modeloSeleccionado)
        {
            foreach (var step in steps)
            {
                step.IsCompleted = false;


                if (FindName(step.PartSecuencia) is TextBlock secuencia)
                    secuencia.Text = $"Rango:{step.MinWeight:F5}/{step.MaxWeight:F5}";

                if (FindName(step.PartIndicator) is Rectangle indicator)
                    indicator.Fill = Brushes.Red;

                if (FindName(step.PartPeso) is TextBlock pesoTextBlock)
                    pesoTextBlock.Text = "0.0Kgs";
            }

            SetValues(pasosFiltrados);
            _accumulatedWeight = 0.0;
            _currentStepIndex = 0;
        }

        private void ActualizarMainWindow()
        {
            CargarModelos();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ShowAlertError("RECUERDA CONFIGURAR LAS ENTRADAS Y EL CATALOGO ANTES DE EMPEZAR");
            IniciarMonitoreoEstadoBasculas();
            CargarModelos();
        }

        private void CargarModelos()
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var modelos = db.Modelos
                        .Where(m => m.Activo == true)
                        .OrderBy(m => m.NoModelo)
                        .ToList();

                    Cbox_Modelo.ItemsSource = modelos;
                    Cbox_Modelo.DisplayMemberPath = "NoModelo";
                    Cbox_Modelo.SelectedValuePath = ".";   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar modelos desde la base de datos: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ModeloExiste(string modelo)
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var modeloDb = db.Modelos.FirstOrDefault(m => m.NoModelo.ToUpper() == modelo.ToUpper());
                    if (modeloDb == null)
                    {
                        ShowAlertError($"Modelo '{modelo}' no encontrado en la base de datos.");
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al verificar el modelo en Excel: {ex.Message}");
                return false;
            }
        }

        private void ProcesarResetModelo()
        {
            try
            {
                if (_currentStepIndex < pasosFiltrados.Count)
                    RegistrarPasoRechazado(pasosFiltrados[_currentStepIndex]);

                CerrarAlertas();

                _stopBascula1 = false;

                _consecutiveCount = 0;
                _accumulatedWeight = 0;
                _currentStepIndex = 0;
                lblProgreso.Content = 0;
                contador = 0;
                stepIndex = 0;
                runningWeight = 0.0;
                valoresBolsas.Clear();
                _etapa1 = true;
                _etapa2 = false;

                txbPesoActual.Text = "0.0Kg";
                PesoGeneral.Text = "0.0Kg";
                codigo = "";
                SetImagesBox();

                if (Cbox_Modelo.SelectedValue == null)
                {
                    ShowAlertError("No hay un modelo seleccionado.");
                    return;
                }

                var modeloSeleccionado = ModeloData.NoModelo;

                if (!ModeloExiste(modeloSeleccionado))
                {
                    ShowAlertError($"Modelo '{modeloSeleccionado}' no reconocido.");
                    return;
                }

                CargarSecuenciasPorModelo(modeloSeleccionado);
                CargarProcesos();
                HideAll();

                if (Cbox_Proceso.SelectedValue == null)
                {
                    ShowAlertError("No hay un proceso seleccionado.");
                    return;
                }

                Cbox_Proceso.SelectedIndex = 0;
                var procesoSeleccionado = int.Parse(Cbox_Proceso.SelectedValue.ToString());
                sequence = ObtenerValoresProceso(modeloSeleccionado, procesoSeleccionado);
                ProcesarModeloValido(ModeloData.NoModelo);
                ProcesarModeloYProceso();
                SetValuesEtapas(sequence, 1);
                SetImagesBox();
                SecuenciaASeguir(ModeloData);
                Cbox_Proceso.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al procesar el reset del modelo: {ex.Message}");
            }
        }

        private void RegistrarPasoRechazado(SequenceStep step)
        {
            if (step != null && !step.IsCompleted && !string.IsNullOrEmpty(step.DetectedWeight))
                LogRejectedStepFerre(step);
        }

        private void CargarProcesos()
        {
            try
            {
                if (Cbox_Modelo.SelectedValue == null)
                    return;

                var procesosDisponibles = sequence
                    .Where(s => s.ModProceso == ModeloData.ModProceso)
                    .Select(s => s.PartProceso)
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();

                Cbox_Proceso.ItemsSource = procesosDisponibles;
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al cargar procesos: {ex.Message}");
            }
            finally
            {
                _isInitializing = false;
            }
        }

        #region VISUALES

        private void HideParts(bool isHidden, List<SequenceStep> steps)
        {
            foreach (var partGrid in steps.Select(step => FindName(step.GrdPart)).OfType<UIElement>())
                partGrid.Visibility = isHidden ? Visibility.Hidden : Visibility.Visible;
        }

        private void HideAll()
        {
            for (var i = 0; i < 8; i++)
                if (FindName($"Part{i}") is UIElement partGrid)
                    partGrid.Visibility = Visibility.Hidden;
        }

        private void HideFerreteria(bool isHidden, List<SequenceStep> lista)
        {
            HideParts(isHidden, lista);
        }

        private void SetValues(List<SequenceStep> steps)
        {
            foreach (var step in steps)
            {
                if (FindName(step.PartImagen) is Image imageControl)
                {
                    var imagePath = Path.Combine(rutaImagenes, $"{step.PartNoParte}.PNG");

                    if (FindName($"Part_Cantidad{int.Parse(step.PartOrden) - 1}") is TextBlock partCantidad)
                        partCantidad.Text = step.PartCantidad;

                    imageControl.Source = File.Exists(imagePath) ? new BitmapImage(new Uri(imagePath, UriKind.Absolute)) : null;
                }


                if (FindName(step.PartIndicator) is UIElement ledLabel) ledLabel.Visibility = Visibility.Visible;
                if (FindName(step.PartPeso) is UIElement ledPeso) ledPeso.Visibility = Visibility.Visible;

                var partPesoIndex = step.PartPeso.Replace("Part_Peso", "");
                var partNoTextBlockName = $"Part_NoParte{partPesoIndex}";
                var partSecuencia = $"Part_Secuencia{partPesoIndex}";


                if (FindName(partSecuencia) is TextBlock partSecuenciaTextBlock)
                    partSecuenciaTextBlock.Text = $"Rango:{step.MinWeight:F5}/{step.MaxWeight:F5}";

                if (FindName(partNoTextBlockName) is TextBlock partNoTextBlock)
                    partNoTextBlock.Text = step.PartNoParte;
            }
        }

        private void SetValuesEtapas(List<SequenceStep> steps, int etapa)
        {
            if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 || ModeloData.UsaBascula1 && !ModeloData.UsaBascula2)
                switch (etapa)
                {
                    case 1:
                        var etapa1 = steps.FirstOrDefault(step => step.PartNoParte.Contains(ModeloData.Etapa1));

                        _etapa1 = true;
                        if (etapa1 != null)
                            pasosFiltrados = steps
                                .Where(s => int.Parse(s.PartOrden) <= int.Parse(etapa1.PartOrden))
                                .ToList();
                        break;

                    case 2:

                        var startStep = steps.FirstOrDefault(step => step.PartNoParte.Contains(ModeloData.Etapa2));

                        _etapa2 = true;
                        if (startStep != null)
                            pasosFiltrados = steps
                                .Where(s => int.Parse(s.PartOrden) >= int.Parse(startStep.PartOrden))
                                .ToList();
                        break;

                    default:
                        throw new ArgumentException($"Etapa '{etapa}' no es válida.");
                }

            else if (!ModeloData.UsaBascula1 && ModeloData.UsaBascula2)
                switch (etapa)
                {
                    case 1:
                        var etapa1 = steps.FirstOrDefault(step => step.PartNoParte.Contains(ModeloData.Etapa2));

                        _etapa1 = true;
                        if (etapa1 != null)
                            pasosFiltrados = steps
                                .Where(s => int.Parse(s.PartOrden) <= int.Parse(etapa1.PartOrden))
                                .ToList();
                        break;
                    default:
                        throw new ArgumentException($"Etapa '{etapa}' no es válida.");
                }

            foreach (var step in pasosFiltrados)
            {
                if (FindName(step.PartImagen) is Image imageControl)
                {
                    var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes",
                        $"{step.PartNoParte}.PNG");


                    if (FindName($"Part_Cantidad{int.Parse(step.PartOrden) - 1}") is TextBlock partCantidad)
                        partCantidad.Text = step.PartCantidad;

                    imageControl.Source = File.Exists(imagePath) ? new BitmapImage(new Uri(imagePath, UriKind.Absolute)) : null;
                }


                if (FindName(step.PartIndicator) is UIElement ledLabel) ledLabel.Visibility = Visibility.Visible;
                if (FindName(step.PartPeso) is UIElement ledPeso) ledPeso.Visibility = Visibility.Visible;

                var partPesoIndex = step.PartPeso.Replace("Part_Peso", "");
                var partNoTextBlockName = $"Part_NoParte{partPesoIndex}";

                if (FindName(partNoTextBlockName) is TextBlock partNoTextBlock)
                    partNoTextBlock.Text = step.PartNoParte;
            }
        }

        private void ReindexarPasos(List<SequenceStep> pasos)
        {
            for (var i = 0; i < pasos.Count; i++)
            {
                var step = pasos[i];

                step.PartImagen = $"Part_Imagen{i}";
                step.PartIndicator = $"Part_Indicator{i}";
                step.PartPeso = $"Part_Peso{i}";
                step.PartSecuencia = $"Part_Secuencia{i}";
                step.GrdPart = $"Part{i}";
                step.PartOrden = $"{i+1}";


                if (FindName(step.PartImagen) is Image imageControl)
                {
                    imageControl.Visibility = Visibility.Visible;
                    imageControl.Source = null;
                }

                if (FindName(step.PartPeso) is TextBlock pesoControl)
                {
                    pesoControl.Visibility = Visibility.Visible;
                    pesoControl.Text = "0.0 kg";
                }

                if (FindName(step.PartIndicator) is Rectangle indicatorControl)
                {
                    indicatorControl.Visibility = Visibility.Visible;
                    indicatorControl.Fill = Brushes.Red;
                }

                if (FindName(step.PartSecuencia) is TextBlock secuencia)
                    secuencia.Text = $"Rango:{step.MinWeight:F5}/{step.MaxWeight:F5}";

                if (FindName(step.GrdPart) is Grid gridControl) gridControl.Visibility = Visibility.Visible;
            }
        }

        private void SetImagesBox()
        {
            var steps = pasosFiltrados;

            HideAll();

            for (var counter = 0; counter < steps.Count; counter++)
            {
                var step = steps[counter];

                var imageControlName = $"Part_Imagen{counter}";

                if (FindName($"Part_Indicator{counter}") is Rectangle indicador && FindName($"Part_Peso{counter}") is TextBlock peso)
                {
                    indicador.Fill = Brushes.Red;
                    peso.Text = "0.0 kg";
                }

                if (FindName($"Part_Secuencia{counter}") is TextBlock secuencia)
                    secuencia.Text = $"Rango:{step.MinWeight:F5}/{step.MaxWeight:F5}";

                if (FindName($"Part_NoParte{counter}") is TextBlock textNoParte)
                    textNoParte.Text = step.PartNoParte;

                if (FindName(imageControlName) is Image imageControl)
                {
                    var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes", $"{step.PartNoParte}.PNG");

                    if (FindName($"Part_Cantidad{counter}") is TextBlock partCantidad)
                        partCantidad.Text = step.PartCantidad;

                    if (FindName($"Part{counter}") is Grid part)
                        part.Visibility = Visibility.Visible;

                    imageControl.Source = File.Exists(imagePath) ? new BitmapImage(new Uri(imagePath, UriKind.Absolute)) : null;
                }
            }
        }

        private void CerrarAlertas()
        {
            if (_alertWindow != null && _alertWindow.IsVisible) _alertWindow.Close();
            if (_errorWindow != null && _errorWindow.IsVisible) _errorWindow.Close();
        }

        private void ShowPruebaCorrecta()
        {
            _alertWindow = new AlertWindow();
            _alertWindow.ShowPruebaCorrecta();
            _alertWindow.Show();
        }

        private void ShowPickToLight(string name)
        {
                lblPesoArt.Text = "TOMAR EL ARTICULO";

                txbArticulo.Text = name.ToUpper();

                txbPesoMax.Visibility = Visibility.Visible;
                txbPesoMin.Visibility = Visibility.Visible;
                txbPesoActual.Visibility = Visibility.Visible;
                lblPesoMin.Visibility = Visibility.Hidden;
                lblPesoMax.Visibility = Visibility.Hidden;
                lblPesoActual.Visibility = Visibility.Hidden;
                btnReset.Visibility = Visibility.Visible;
                btnRechazo.Visibility = Visibility.Visible;

                txbPesoMax.Text = $"";
                txbPesoMin.Text = $"";
                txbPesoActual.Text = $"";

                grdValidacion.Visibility = Visibility.Visible;
                recValidacion.Visibility = Visibility.Visible;
        }

        private void ShowBolsasRestantes(string name, double pesoMin, double pesoMax, double pesoActual,
            bool validacion, int bolsasRestantes)
        {
            if (!validacion)
            {
                lblPesoArt.Text = "FAVOR DE PESAR EL ARTICULO";

                txbPesoMax.Visibility = Visibility.Visible;
                txbPesoMin.Visibility = Visibility.Visible;
                txbPesoActual.Visibility = Visibility.Visible;
                lblPesoMin.Visibility = Visibility.Visible;
                lblPesoMax.Visibility = Visibility.Visible;
                lblPesoActual.Visibility = Visibility.Visible;
                txbScanner.Visibility = Visibility.Hidden;
                btnReset.Visibility = Visibility.Visible;
                btnRechazo.Visibility = Visibility.Visible;


                txbArticulo.Text = name.ToUpper() + $" Y RESTAN: {bolsasRestantes}";
                txbPesoMax.Text = $"{pesoMax:F5} kg";
                txbPesoMin.Text = $"{pesoMin:F5} kg";
                txbPesoActual.Text = $"{pesoActual:F5} kg";

                grdValidacion.Visibility = Visibility.Visible;
                recValidacion.Visibility = Visibility.Visible;
            }
        }

        private void ShowAlertError(string mensaje)
        {
            if (_errorWindow != null && _errorWindow.IsVisible)
                _errorWindow.Close();

            if (_alertWindow != null && _alertWindow.IsVisible)
                _alertWindow.Close();

            _errorWindow = new ErrorMessageWindow
            {
                TitleText = "ADVERTENCIA",
                Message = mensaje
            };

            _errorWindow.Show();
        }

        public static async Task ShowMensaje(string message, Brush color, int time)
        {
            var messageBox = new Window
            {
                Title = "Alerta",
                Height = 500,
                Width = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = color,
                WindowStyle = WindowStyle.ToolWindow
            };

            var stackPanel = new StackPanel
                { VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };

            var textBlock = new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 30,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Black,
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(10)
            };

            stackPanel.Children.Add(textBlock);
            messageBox.Content = stackPanel;

            messageBox.Show();

            await Task.Delay(time);
            messageBox.Close();
        }

        private void IniciarMonitoreoEstadoBasculas()
        {
            _estadoBasculasTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(1000)
            };
            _estadoBasculasTimer.Tick += VerificarEstadoBasculas;
            _estadoBasculasTimer.Start();
        }

        private void VerificarEstadoBasculas(object sender, EventArgs e)
        {
            borderBascula1.Background = _stopBascula1 ? Brushes.Red : Brushes.Green;
        }

        #endregion

        #region SEALEVEL

        private void ActivarSalida(string tipo)
        {
            /*var salida = ObtenerSalidaPorTipo(tipo);
            ioInterface.WriteSingleOutput(salida, true);*/
        }

        private void DesactivarSalida(string tipo)
        {
           /* var salida = ObtenerSalidaPorTipo(tipo);
            ioInterface.WriteSingleOutput(salida, false);*/
        }

        private int ObtenerSalidaPorTipo(string tipo)
        {
            switch (tipo.ToLower())
            {
                //case "unitaria":
                //    return defaultSettings.SalidaDispensadoraUnitaria;
                //case "master":
                //    return defaultSettings.SalidaDispensadoraMaster;
                //case "prop65":
                //    return defaultSettings.SalidaDispensadoraProp65;
                //case "selladora":
                //    return defaultSettings.SalidaSelladora;
                default:
                    throw new ArgumentException($"Tipo de salida '{tipo}' no válido.");
            }
        }

        private void IniciarSealevel()
        {
            if (ioScannerActivado)
                return;

            ioInterface = new IOInterface(defaultSettings.PuertoSealevel);
            ioScanner = new IOScanner(ioInterface);

            ioScanner.Tick += OnScannerTick;
            ioScanner.Start();
            ioScannerActivado = true;
        }

        private void OutputsOff()
        {
            try
            {
                if (ioScanner != null && ioScanner.IsRunning())
                {
                    ioInterface?.WriteMultipleOutputs(0, 0, 16);
                }
                else
                {
                    IniciarSealevel();
                    ioInterface?.WriteMultipleOutputs(0, 0, 16);
                    ioScanner?.Stop();
                    ioInterface?.Dispose();
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"OutputsOff: {ex.Message}");
            }
        }

        private void OnScannerTick(object sender, uint inputsState)
        {
            Application.Current.Dispatcher.Invoke(async () =>
            {
                //var sensorUnitaria = (inputsState & (1 << defaultSettings.EntradaSensorUnitaria)) != 0;
                //var sensorProp65 = (inputsState & (1 << defaultSettings.EntradaSensorProp65)) != 0;
                //var sensorMaster = (inputsState & (1 << defaultSettings.EntradaSensorMaster)) != 0;
                //var sensorSelladora = (inputsState & (1 << defaultSettings.EntradaSensorSelladora)) != 0;

                //if (sensorUnitaria && !sensorUnitariaActivo)
                //    DesactivarSalida("unitaria");

                //if (sensorMaster && !sensorMasterActivo)
                //    DesactivarSalida("master");


                //if (sensorProp65 && !sensorProp65Activo)
                //    DesactivarSalida("prop65");

                //if (sensorSelladora && sensorSelladoraActivo && _activarSelladora)
                //{
                //    ActivarSalida("selladora");

                //    await Task.Delay(700);

                //    DesactivarSalida("selladora");
                //    _activarSelladora = false;
                //}

                //sensorUnitariaActivo = sensorUnitaria;
                //sensorMasterActivo = sensorMaster;
                //sensorProp65Activo = sensorProp65;
                //sensorSelladoraActivo = sensorSelladora;
            });
        }

        #endregion

        #region LOGS

        private void LogFerreteriaStep(SequenceStep step, string resultado, string matchedTag)
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var registro = new Completado
                    {
                        Fecha = DateTime.Now,
                        NoParte = step.PartNoParte,
                        ModProceso = step.ModProceso,
                        Proceso = step.PartProceso,
                        PesoDetectado = double.TryParse(step.DetectedWeight, out var peso) ? peso : 0,
                        Estado = resultado,
                        Tag = matchedTag ?? ""
                    };

                    db.Completados.Add(registro);
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al registrar ferretería: {ex.Message}");
            }
        }


        private void LogRejectedStepFerre(SequenceStep step)
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var registro = new Completado
                    {
                        Fecha = DateTime.Now,
                        NoParte = step.PartNoParte,
                        ModProceso = step.ModProceso,
                        Proceso = step.PartProceso,
                        PesoDetectado = double.TryParse(step.DetectedWeight, out var peso) ? peso : 0,
                        Estado = "Rechazado",
                        Tag = step.Tag ?? ""
                    };

                    db.Completados.Add(registro);
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al registrar ferretería: {ex.Message}");
            }
        }

        #endregion

        #region Basculas

        private async void ReadInputBascula1()
        {
            await InicializarPuertoDeBasculaAsync(defaultSettings.PuertoBascula1, bascula);
        }

        private async Task InicializarPuertoDeBasculaAsync(string puerto, BasculaFunc bascula)
        {
            await Task.Run(() =>
            {
                bool error = false;

                try
                {
                    if (string.IsNullOrWhiteSpace(puerto))
                    {
                        error = true;
                        Application.Current.Dispatcher.Invoke(() =>
                            ShowAlertError("El puerto de la báscula no está configurado."));
                        return;
                    }

                    var puertoBascula = new SerialPort(puerto, defaultSettings.BaudRateBascula12,
                        Parity.None, defaultSettings.DataBitsBascula12, StopBits.One);

                    bascula.AsignarPuertoBascula(puertoBascula);
                    bascula.OpenPort();
                }
                catch (IOException ioEx)
                {
                    error = true;
                    bascula.ClosePort();
                    Application.Current.Dispatcher.Invoke(() =>
                        ShowAlertError($"No se pudo abrir el puerto ({puerto}) de comunicación con la báscula.\n\n" +
                                       "Verifica que sea el puerto correcto o intenta cerrando y abriendo la aplicación nuevamente."));
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    error = true;
                    bascula.ClosePort();
                    Application.Current.Dispatcher.Invoke(() =>
                        ShowAlertError($"El puerto ({puerto}) ya está siendo usado por otra aplicación."));
                }
                catch (Exception ex)
                {
                    error = true;
                    bascula.ClosePort();
                    Application.Current.Dispatcher.Invoke(() =>
                        ShowAlertError($"Error inesperado al abrir el puerto ({puerto}): {ex.Message}"));
                }
                finally
                {
                    if (error)
                    {
                        _stopBascula1 = true;
                    }
                }
            });
        }

        private void Bascula1_OnDataReady(object sender, BasculaEventArgs e)
        {
            if (_stopBascula1) return;

            Dispatcher.Invoke(() =>
            {
                var weight = e.Value;
                var isStable = e.IsStable;
                PesoGeneral.Text = $"Peso: {weight:F5} kg";
                if (isStable)
                {
                    if (Math.Abs(weight - _lastWeight) < 0.0001)
                        _consecutiveCount++;
                    else
                        _consecutiveCount = 1;

                    _lastWeight = weight;

                    if (_consecutiveCount == 3)
                        ProcessStableWeight(weight);
                }

            });
        }

        private void SecuenciaASeguir(Modelo modeloseleccionado)
        {
            try
            {
                if (modeloseleccionado.UsaBascula1)
                {

                    if (bascula?.GetPuerto() == null || !bascula.GetPuerto().IsOpen)
                        ReadInputBascula1();

                    _stopBascula1 = false;
                }
                else
                    ShowAlertError($"Error al cargar la secuencia, secuencia inexistente.");
            }
            catch (Exception e)
            {
                ShowAlertError(e.ToString());
                throw;
            }

        }

        private void ProcessStableWeight(double currentWeight)
        {
            _consecutiveCount = 0;
            var pasosPorPagina = 8;
            var totalPasos = pasosFiltrados.Count;

            var paginaActual = _currentStepIndex / pasosPorPagina;

            var pasosAMostrar = pasosFiltrados
                .Skip(paginaActual * pasosPorPagina)
                .Take(pasosPorPagina)
                .ToList();

            var indexEnPagina = _currentStepIndex % pasosPorPagina;
            if (_currentStepIndex > 0 && indexEnPagina == 0)
            {
                _currentStepIndex = 0;
                HideAll();
                pasosFiltrados = pasosAMostrar;
                ReindexarPasos(pasosFiltrados);
                SetImagesBox();
            }

            if (_currentStepIndex >= totalPasos) return; 

            var currentStep = pasosAMostrar[indexEnPagina];
            pieceWeight = currentWeight - _accumulatedWeight;


            if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock cantidadTextBlock)) return;

            ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight,
                _validacion, int.Parse(cantidadTextBlock.Text));

            if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) || !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

            if (pieceWeight >= currentStep.MinWeight && pieceWeight <= currentStep.MaxWeight)
            {
                CompleteCurrentStep(currentStep, indicator, pesoTextBlock, cantidadTextBlock, currentWeight);

                _currentStepIndex++;

                if (_currentStepIndex < pasosFiltrados.Count)
                {
                    pasosPorPagina = 8;

                    if (_currentStepIndex % pasosPorPagina == 0)
                    {
                        paginaActual = _currentStepIndex / pasosPorPagina;

                        HideAll();
                        pasosFiltrados = pasosFiltrados
                            .Skip(paginaActual * pasosPorPagina)
                            .Take(pasosPorPagina)
                            .ToList();

                        ReindexarPasos(pasosFiltrados);
                        SetImagesBox();

                        _currentStepIndex = 0;
                    }

                    currentStep = pasosFiltrados[_currentStepIndex];
                    pieceWeight = currentWeight - _accumulatedWeight;

                    ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                        pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));
                    return;
                }
                ShowPickToLight(currentStep.PartNoParte);
                var lastBag = valoresBolsas.LastOrDefault();

                if (lastBag != null)
                {
                    lastBag.PartNoParte = currentStep.PartNoParte;
                    lastBag.MinWeight = currentStep.MinWeight;
                    lastBag.MaxWeight = currentStep.MaxWeight;
                    lastBag.DetectedWeight = currentWeight.ToString("F5");
                    lastBag.PartIndicator = currentStep.PartIndicator;
                    lastBag.PartPeso = currentStep.PartPeso;
                    lastBag.PartOrden = currentStep.PartOrden;
                    lastBag.PartCantidad = currentStep.PartCantidad;
                    lastBag.PartProceso = currentStep.PartProceso;
                }

                _currentStepIndex = 0;
                _accumulatedWeight = 0;
                pieceWeight = 0;
                SetImagesBox();
                contador++;
                lblCompletados.Content = contador;
                codigo = "";

                Dispatcher.Invoke(ShowPruebaCorrecta);
            }

            else
            {
                indicator.Fill = Brushes.Red;
                pesoTextBlock.Text = $"{pieceWeight:F5} kg";
                currentStep.DetectedWeight = currentWeight.ToString("F5");
            }
        }

        private void CompleteCurrentStep(SequenceStep currentStep, Rectangle indicator, TextBlock pesoTextBlock,
            TextBlock cantidadTextBlock, double currentWeight)
        {
            currentStep.DetectedWeight = pieceWeight.ToString();
            indicator.Fill = Brushes.Green;
            pesoTextBlock.Text = $"{pieceWeight:F5} kg";

            if (int.TryParse(cantidadTextBlock.Text, out var cantidadRestante) && cantidadRestante > 0)
            {
                cantidadRestante--;
                cantidadTextBlock.Text = cantidadRestante.ToString();
                _accumulatedWeight = currentWeight;

                if (cantidadRestante == 0)
                {
                    if (!currentStep.IsCompleted)
                        LogFerreteriaStep(currentStep, "OK", null);

                    if (_currentStepIndex == pasosFiltrados.Count - 1)
                    {
                        (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(ModeloData.NoModelo);
                        codigo = $"{integrer}.{fraction}";

                        valoresBolsas.Add(new SequenceStep
                        {
                            PartNoParte = currentStep.PartNoParte,
                            MinWeight = currentStep.MinWeight,
                            MaxWeight = currentStep.MaxWeight,
                            DetectedWeight = currentWeight.ToString(),
                            Tag = codigo,
                            IsCompleted = false,
                            PartIndicator = currentStep.PartIndicator,
                            PartPeso = currentStep.PartPeso,
                            PartOrden = currentStep.PartOrden,
                            PartCantidad = currentStep.PartCantidad
                        });
                    }
                }
            }
        }
        #endregion

        #region Botones

        private void BtnCerrarPeso_Click(object sender, RoutedEventArgs e)
        {
            _validacion = true;
            recValidacion.Visibility = Visibility.Hidden;
            grdValidacion.Visibility = Visibility.Hidden;
        }

        private void BtnValidaciones_Click(object sender, RoutedEventArgs e)
        {
            switch (_validacion)
            {
                case true:
                    _validacion = false;
                    recValidacion.Visibility = Visibility.Visible;
                    grdValidacion.Visibility = Visibility.Visible;
                    break;
                case false:

                    _validacion = true;
                    recValidacion.Visibility = Visibility.Hidden;
                    grdValidacion.Visibility = Visibility.Hidden;
                    break;
            }
        }

        private void Configuracion_btn_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationWindow auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                var newConfig = new ConfiguracionWindow();
                newConfig.Show();
                newConfig.Focus();
            }
        }

        private void BtnCatalogos_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationWindow auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                _catalogos = new Catalogos();
                _catalogos.CambiosGuardados += ActualizarMainWindow;
                _catalogos.Show();
            }
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationWindow auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                if (Cbox_Modelo.SelectedItem == null)
                {
                    ShowAlertError("No hay un modelo seleccionado.");
                    return;
                }

                var modeloSeleccionado = ModeloData.NoModelo.ToUpper();

                if (ModeloExiste(modeloSeleccionado))
                {
                    ProcesarResetModelo();
                }
                else
                {
                    ShowAlertError($"Modelo '{modeloSeleccionado}' no reconocido.");
                }
            }
        }

        private void BtnRechazo_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationWindow auth = new AuthenticationWindow();

            if (auth.ShowDialog() != true) return;

            try
            {

                if (Cbox_Modelo.SelectedValue == null)
                {
                    ShowAlertError("No hay un modelo seleccionado.");
                    return;
                }

                var modeloSeleccionado = ModeloData.NoModelo;

                if (!ModeloExiste(modeloSeleccionado))
                {
                    ShowAlertError($"Modelo '{modeloSeleccionado}' no reconocido.");
                    return;
                }

                pasosFiltrados = sequence;

                if (_currentStepIndex < pasosFiltrados.Count)
                {
                    var currentStep = pasosFiltrados[_currentStepIndex];

                    RegistrarPasoRechazado(currentStep);

                    _ = ShowMensaje($"Se rechazó la pieza: {currentStep.PartNoParte}", Brushes.IndianRed, 5000);
                }
                else
                {
                    ShowAlertError("No hay una pieza válida para rechazar en FERRETERÍA.");
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al registrar el rechazo: {ex.Message}");
            }
        }

        #endregion 
    }
}