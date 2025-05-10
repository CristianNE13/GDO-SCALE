using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        private IBasculaFunc bascula;
        private readonly List<SequenceStep> valoresBolsas = new List<SequenceStep>();
        private KeyenceTcpClient keyence;
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
        private string zpl;
        private List<SequenceStep> sequence;
        private DispatcherTimer _estadoBasculasTimer;
        private bool _inicioZero = true;
        private bool _esperandoPickToLight = false;
        private SequenceStep _stepEsperandoPick;
        private bool _manual;
        private bool _activarBoton = false;
        private bool sensorActivo = false;
        private bool _verificarArticulo = false;
        private bool _verificacionCompletado = false;
        private bool botonActivo;
        private bool _ignorarInput = false;
        private bool _siguienteManualCompletado = false;
        private bool _siguienteManual = false;


        private readonly string rutaImagenes = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes");

        public MainWindow()
        {
            InitializeComponent();

            defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);

            keyence = new KeyenceTcpClient(defaultSettings.IpCamara, defaultSettings.PuertoCamara);

            if (defaultSettings.BasculaMarca == "Pennsylvania")
                bascula = new BasculaFuncPennsylvania();
            else
                bascula = new BasculaFuncGFC();

            bascula.AsignarControles(Dispatcher);
            bascula.OnDataReady += Bascula1_OnDataReady;

            keyence.OnCameraStatusChanged += UpdateCameraStatus;
        }


        private void Cbox_Modelo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                recValidacion.Visibility = Visibility.Hidden;
                grdValidacion.Visibility = Visibility.Hidden;
                lbx_Codes.Visibility = Visibility.Hidden;

                defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);
                IniciarSealevel();
                OutputsOff();

                ModeloData = Cbox_Modelo.SelectedValue as Modelo ?? throw new Exception("Modelo no reconocido.");

                ProcesarModeloValido(ModeloData.NoModelo);
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

        private async void Cbox_Proceso_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lbx_Codes.Visibility = Visibility.Hidden;

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

                _inicioZero = true;
                ShowIniciar();

                if (defaultSettings.CheckShutOff)
                    ActivarSalida(defaultSettings.ShutOff);


                if (ModeloData.UsaCamaraVision)
                {
                    borderCamaraStatus.Visibility = Visibility.Visible;

                    var programasVision = (ModeloData.ProgramaVision ?? "")
                        .Split(',')
                        .Select(p => p.Trim())
                        .Where(p => !string.IsNullOrEmpty(p))
                        .ToList();

                    foreach (var programa in programasVision)
                    {
                        if (int.TryParse(programa, out var numero)) continue;
                        MessageBox.Show($"El programa de visión '{programa}' no es válido.", "Error", MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        return;
                    }

                    keyence.StopMonitoring();

                    var connected = await keyence.ConnectAsync();
                    if (!connected)
                    {
                        MessageBox.Show("No se pudo conectar a la cámara.", "Error");
                        UpdateCameraStatus(0);
                    }
                    keyence.StartMonitoring();

                    if (ModeloData.ProgramaVision != null)
                        await keyence.ChangeProgram(int.Parse(programasVision[procesoSeleccionado - 1]));
                }
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
                Id = articulo.Id,
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
                OutputsOff();

                if (bascula.GetPuerto() != null && bascula.GetPuerto().IsOpen)
                    bascula.ClosePort();

                keyence?.Dispose();

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
            HideAll();
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
                _manual = false;

                _consecutiveCount = 0;
                _accumulatedWeight = 0;
                _currentStepIndex = 0;
                lblProgreso.Content = 0;
                valoresBolsas.Clear();
                _etapa1 = true;
                _etapa2 = false;

                if (ioScannerActivado)
                    ioInterface.WriteSingleOutput(defaultSettings.Piston,false);

                txbPesoActual.Text = "0.0Kg";
                PesoGeneral.Text = "0.0Kg";
                codigo = "";

                _inicioZero = true;
                _activarBoton = true;

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



        private void UpdateCameraStatus(int status)
        {
            Dispatcher.Invoke(() =>
            {
                switch (status)
                {
                    case 0:
                        borderCamaraStatus.Background = Brushes.Red;
                        break;
                    case 1:
                        borderCamaraStatus.Background = Brushes.Yellow;
                        break;
                    case 2:
                        borderCamaraStatus.Background = Brushes.Green;
                        break;
                }
            });
        }

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

                if (FindName($"QuickSetup{partPesoIndex}_Btn") is Button quickSetupButton)
                    quickSetupButton.DataContext = step;
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

        private void ShowPickToLight(SequenceStep current)
        {
            SalidaPick2Orden(current.PartOrden, true);
            _esperandoPickToLight = true;
            _stepEsperandoPick = current;

            lblPesoArt.Text = "TOMAR EL ARTICULO";

            txbArticulo.Text = $"PICK2LIGHT-{int.Parse(current.PartOrden)-1}";

            txbPesoMax.Visibility = Visibility.Visible;
            txbPesoMin.Visibility = Visibility.Visible;
            txbPesoActual.Visibility = Visibility.Visible;
            lblPesoMin.Visibility = Visibility.Hidden;
            lblPesoMax.Visibility = Visibility.Hidden;
            lblPesoActual.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Visible;
            btnRechazo.Visibility = Visibility.Visible;

            txbPesoMax.Text = "";
            txbPesoMin.Text = "";
            txbPesoActual.Text = current.PartNoParte.ToUpper();;

            grdValidacion.Visibility = Visibility.Visible;
            recValidacion.Visibility = Visibility.Visible;

            btnCerrarPeso.Visibility = Visibility.Hidden;
            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnRechazo.Visibility = Visibility.Hidden;
            btnIniciarZero.Visibility = Visibility.Hidden;
        }

        private async void HidePickToLight(SequenceStep current)
        {
            grdValidacion.Visibility = Visibility.Hidden;
            recValidacion.Visibility = Visibility.Hidden;
            await Task.Delay(1500);
            SalidaPick2Orden(current.PartOrden,false);
        }

        private void ShowSensor(string articulo)
        {
            _stopBascula1 = true;
            lblPesoArt.Text = "";

            txbArticulo.Text = $"VERIFICAR {articulo} CON SENSOR";

            lblPesoMin.Visibility = Visibility.Hidden;
            lblPesoMax.Visibility = Visibility.Hidden;
            lblPesoActual.Visibility = Visibility.Hidden;

            txbPesoMax.Visibility = Visibility.Hidden;
            txbPesoMin.Visibility = Visibility.Hidden;
            txbPesoActual.Visibility = Visibility.Hidden;

            txbPesoMax.Text = "";
            txbPesoMin.Text = "";
            txbPesoActual.Text = "";

            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Visible;
            btnRechazo.Visibility = Visibility.Visible;

            grdValidacion.Visibility = Visibility.Visible;
            recValidacion.Visibility = Visibility.Visible;

            btnIniciarZero.Visibility = Visibility.Hidden;
            btnCerrarPeso.Visibility = Visibility.Hidden;
            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnRechazo.Visibility = Visibility.Hidden;
        }

        private void HideSensor()
        {
            _stopBascula1 = false;
            lblPesoArt.Text = "";
            txbArticulo.Text = "";

            lblPesoMin.Visibility = Visibility.Hidden;
            lblPesoMax.Visibility = Visibility.Hidden;
            lblPesoActual.Visibility = Visibility.Hidden;

            txbPesoMax.Visibility = Visibility.Hidden;
            txbPesoMin.Visibility = Visibility.Hidden;
            txbPesoActual.Visibility = Visibility.Hidden;

            txbPesoMax.Text = "";
            txbPesoMin.Text = "";
            txbPesoActual.Text = "";

            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Visible;
            btnRechazo.Visibility = Visibility.Visible;

            grdValidacion.Visibility = Visibility.Hidden;
            recValidacion.Visibility = Visibility.Hidden;

            btnIniciarZero.Visibility = Visibility.Hidden;
            btnCerrarPeso.Visibility = Visibility.Hidden;
            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnRechazo.Visibility = Visibility.Hidden;
        }

        private void ShowIniciar()
        {
            _activarBoton = true;

            lblPesoArt.Text = "PRESIONAR BOTON DE INICIO";
            lblPesoMin.Visibility = Visibility.Hidden;
            lblPesoMax.Visibility = Visibility.Hidden;
            lblPesoActual.Visibility = Visibility.Hidden;

            txbPesoMax.Visibility = Visibility.Hidden;
            txbPesoMin.Visibility = Visibility.Hidden;
            txbPesoActual.Visibility = Visibility.Hidden;
            txbArticulo.Text = "SET BASCULA A ZERO";
            txbPesoMax.Text = "";
            txbPesoMin.Text = "";
            txbPesoActual.Text = "";

            btnIniciarZero.Visibility = Visibility.Hidden;
            btnCerrarPeso.Visibility = Visibility.Hidden;
            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Visible;
            btnRechazo.Visibility = Visibility.Hidden;

            grdValidacion.Visibility = Visibility.Visible;
            recValidacion.Visibility = Visibility.Visible;
        }

        private void ShowAlertCamara()
        {
            if (_manual)
                lblPesoArt.Text = "INSPECCION DE MANUAL O CARTON";
            else
                lblPesoArt.Text = "INSPECCION CON CAMARA";

            lblPesoMin.Visibility = Visibility.Hidden;
            lblPesoMax.Visibility = Visibility.Hidden;
            lblPesoActual.Visibility = Visibility.Hidden;

            txbPesoMax.Visibility = Visibility.Hidden;
            txbPesoMin.Visibility = Visibility.Hidden;
            txbPesoActual.Visibility = Visibility.Hidden;
            txbArticulo.Text = "PRESIONAR BOTON";
            txbPesoMax.Text = "";
            txbPesoMin.Text = "";
            txbPesoActual.Text = "";

            grdValidacion.Visibility = Visibility.Visible;
            recValidacion.Visibility = Visibility.Visible;

            btnIniciarZero.Visibility = Visibility.Hidden;
            btnInspeccionCamara.Visibility = Visibility.Hidden;
            btnCerrarPeso.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Visible;
            btnRechazo.Visibility = Visibility.Hidden;
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
                btnIniciarZero.Visibility = Visibility.Hidden;


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
                Width = 750,
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

        private void SalidaPick2Orden(string partOrden, bool activo)
        {
            int start = defaultSettings.InputPick2L0;

            if (int.TryParse(partOrden, out int offset) && offset > 0)
            {
                int index = start + (offset - 1);

                if (activo)
                    ActivarSalida(index);
                else
                    DesactivarSalida(index);
            }
            else
            {
                MessageBox.Show($"Orden no válida o fuera de rango: {partOrden}");
            }
        }


        private void ActivarSalida(int tipo)
        {
            ioInterface.WriteSingleOutput(tipo, true);
        }

        private void DesactivarSalida(int tipo)
        {
            ioInterface.WriteSingleOutput(tipo, false);
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

                if (_esperandoPickToLight && _stepEsperandoPick != null)
                {
                    if (!int.TryParse(_stepEsperandoPick.PartOrden, out int ordenPaso)) return;

                    var sensorPick = (inputsState & (1 << ordenPaso-1)) != 0;

                    if (sensorPick)
                    {
                        HidePickToLight(_stepEsperandoPick);
                        _esperandoPickToLight = false;
                        _stepEsperandoPick = null;
                    }
                }

                //var sensorHeader = (inputsState & (1 << defaultSettings.Fixture)) != 0;
                var botonInspeccion = (inputsState & (1 << defaultSettings.InputBoton)) != 0;

                //if (sensorHeader && !sensorActivo && _verificarArticulo)
                //{
                //    _verificarArticulo = false;
                //    _verificacionCompletado = true;
                //    HideSensor();
                //}

                if (botonInspeccion && !botonActivo && _activarBoton)
                {
                    if (_ignorarInput)
                        return;

                    _ignorarInput = true;
                    InspeccionarValidacionFunc();

                    await Task.Delay(2000);
                    _ignorarInput = false;
                }
                
                //sensorActivo = sensorHeader;
                botonActivo = botonInspeccion;
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

        private async Task InicializarPuertoDeBasculaAsync(string puerto, IBasculaFunc bascula)
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
                        Parity.None, defaultSettings.DataBitsBascula12);

                    bascula.AsignarPuertoBascula(puertoBascula);
                    bascula.OpenPort();
                }
                catch (IOException)
                {
                    error = true;
                    bascula.ClosePort();
                    Application.Current.Dispatcher.Invoke(() =>
                        ShowAlertError($"No se pudo abrir el puerto ({puerto}) de comunicación con la báscula.\n\n" +
                                       "Verifica que sea el puerto correcto o intenta cerrando y abriendo la aplicación nuevamente."));
                }
                catch (UnauthorizedAccessException)
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

            if (ModeloData.UsaPick2Light && _esperandoPickToLight)
                return;

            if (_inicioZero || _verificarArticulo)
                return;

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

                    //if (_consecutiveCount == 1) //ANTERIOR
                    //    ProcessStableWeight(weight);

                    if (_currentStepIndex == 0 && weight < -0.02 && !_inicioZero && defaultSettings.BasculaMarca != "GFC")
                    {
                        if (_consecutiveCount >= 2) // Se puede ajustar a 2 o 3 lecturas estables consecutivas
                        {
                            bascula.EnviarZero();
                            _inicioZero = true; // Marcamos que se ha enviado el Zero para evitar repetidos
                        }
                    }
                    else if (_consecutiveCount == 1)
                    {
                        ProcessStableWeight(weight);
                    }
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

        private async void ProcessStableWeight(double currentWeight)
        {
            _consecutiveCount = 0;

            //if (_currentStepIndex == 0 && currentWeight < -0.02 && !_inicioZero && defaultSettings.BasculaMarca != "GFC")
            //{
            //    bascula.EnviarZero();
            //    return;
            //} ANTERIOR

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

            if (currentStep.PartNoParte.Contains("VERIFY") && !_verificarArticulo)
            {
                ShowSensor(currentStep.PartNoParte);
                _verificarArticulo = true;
            }

            if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock cantidadTextBlock)) return;

            ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight,
                _validacion, int.Parse(cantidadTextBlock.Text));

            if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) || !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

            if (!currentStep.PartNoParte.Contains("MANUAL") && !currentStep.PartNoParte.Contains("CARTON") && pieceWeight >= currentStep.MinWeight && pieceWeight <= currentStep.MaxWeight)
            {
                CompleteCurrentStep(currentStep, indicator, pesoTextBlock, cantidadTextBlock, currentWeight);

                if (_currentStepIndex + 1 < pasosFiltrados.Count)
                {
                    var siguientepaso = pasosFiltrados[_currentStepIndex + 1];
                    if (siguientepaso.PartNoParte.Contains("MANUAL") || siguientepaso.PartNoParte.Contains("CARTON"))
                    {
                        _siguienteManual = true;

                        ActivarCamaraValidacion();
                        return;
                    }
                }

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

                    if (ModeloData.UsaPick2Light && _currentStepIndex < 6)
                        ShowPickToLight(currentStep); 

                    ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                        pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));
                    return;
                }

                if (ModeloData.UsaCamaraVision)
                {
                    string[] lista = ModeloData.ProgramaVision.Split(',');

                    await keyence.ChangeProgram(int.Parse(lista[1]));

                    ActivarCamaraValidacion();
                    return;
                }

                _manual = false;
                _currentStepIndex = 0;
                _accumulatedWeight = 0;
                pieceWeight = 0;
                SetImagesBox();
                contador++;
                lblCompletados.Content = contador;
                codigo = "";
                _inicioZero = true;
                ShowIniciar();
                Dispatcher.Invoke(ShowPruebaCorrecta);
            }

            if (currentStep.PartNoParte.Contains("MANUAL") || currentStep.PartNoParte.Contains("CARTON"))
            {
                _manual = true;
                ActivarCamaraValidacion();
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

                    if (!currentStep.IsCompleted)
                        LogFerreteriaStep(currentStep, "OK", codigo);
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
                    try
                    {
                        if (Cbox_Modelo.SelectedValue == null || pasosFiltrados == null)
                        {
                            ShowAlertError("No hay un modelo y proceso seleccionado.");
                            return;
                        }
                        var currentStep = pasosFiltrados[_currentStepIndex];
                        ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight, pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));

                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception);
                        throw;
                    }

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

        private async void btnIniciarZero_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] lista = ModeloData.ProgramaVision.Split(',');

                bool connected = await keyence.ConnectAsync();
                if (!connected)
                {
                    MessageBox.Show("No se pudo conectar a la cámara.", "Error");
                    return;
                }
                await keyence.ChangeProgram(int.Parse(lista.First()));

                _inicioZero = false;

                bascula.EnviarZero();

                var valoresNG = "";
                lbx_Codes.Items.Clear();
                lbx_Codes.Visibility = Visibility.Visible;

                string response = await keyence.SendTrigger();
                List<string> resultado = keyence.Formato(response);

                foreach (var item in resultado)
                    lbx_Codes.Items.Add(item);

                var ngValores = lbx_Codes.Items.Cast<string>()
                    .Where(x => x.Contains("NG"))
                    .ToList();

                if (ngValores.Any())
                {
                    valoresNG = string.Join(",", ngValores);
                    _stopBascula1 = true;

                    ShowMensaje($"Errores: {valoresNG}", Brushes.Beige, 3000);
                    return;
                }

                recValidacion.Visibility = Visibility.Hidden;
                grdValidacion.Visibility = Visibility.Hidden;
                lbx_Codes.Visibility = Visibility.Hidden;
                ioInterface.WriteSingleOutput(6,true);

                var currentStep = pasosFiltrados[_currentStepIndex];

                if (ModeloData.UsaPick2Light)
                    ShowPickToLight(currentStep);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al inicializar báscula: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ActivarCamaraValidacion()
        {
            ShowAlertCamara();
            _stopBascula1 = true;
            _activarBoton = true;
        }

        private void btnInspeccionCamara_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InspeccionarValidacionFunc();
            }
            catch (Exception error)
            {
                MessageBox.Show($"Error: {error.Message}", "ERROR");
                throw;
            }
        }

        private async void InspeccionarValidacionFunc()
        {
            try
            {
                string [] lista = ModeloData.ProgramaVision.Split(',');
                bool error;

                if (_inicioZero)
                {
                    await keyence.ChangeProgram(int.Parse(lista.First()));

                    bascula.EnviarZero();

                    error = await VerificacionCamara();

                    if (error)
                    {
                        await ShowMensaje($"Error, verificar imagen de camara", Brushes.Beige, 3000);
                        return;
                    }

                    recValidacion.Visibility = Visibility.Hidden;
                    grdValidacion.Visibility = Visibility.Hidden;
                    lbx_Codes.Visibility = Visibility.Hidden;

                    ActivarSalida(defaultSettings.Piston);

                    var currentStep = pasosFiltrados[_currentStepIndex];

                    _stopBascula1 = false;
                    _activarBoton = false;
                    _inicioZero = false;

                    ShowMensaje("INSPECCION CORRECTA", Brushes.Green, 1500);

                    if (ModeloData.UsaPick2Light)
                        ShowPickToLight(currentStep);
                }

                else
                {
                    if (_siguienteManual && !_siguienteManualCompletado)
                        await keyence.ChangeProgram(int.Parse(lista[1]));

                    else if (_manual)
                        await keyence.ChangeProgram(int.Parse(lista.Last()));

                    error = await VerificacionCamara();

                    if (error)
                    {
                        await ShowMensaje($"Error, verificar imagen de camara", Brushes.Beige, 3000);
                        return;
                    }

                    if (_manual)
                    {
                        var currentStep = pasosFiltrados[_currentStepIndex];
                        if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock cantidadTextBlock)) return;
                        if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) || !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

                        indicator.Fill = Brushes.Green;
                        cantidadTextBlock.Text = "0";
                        pesoTextBlock.Text = "OK";

                        CamaraCompletada();
                    }

                    if (_siguienteManual)
                    {
                        ShowMensaje($"INSPECCION CORRECTA", Brushes.Green, 1500);

                        lbx_Codes.Visibility = Visibility.Hidden;

                        var currentStep = pasosFiltrados[_currentStepIndex];
                        if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock cantidadTextBlock)) return;
                        if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) || !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

                        //indicator.Fill = Brushes.Green;
                        //cantidadTextBlock.Text= "0";
                        //pesoTextBlock.Text= "OK";

                        _currentStepIndex++;
                        _stopBascula1 = false;
                        _siguienteManualCompletado = true;
                        
                    }

                    else
                        CamaraCompletada();
                }
            }
            catch (Exception error)
            {
                MessageBox.Show($"Error: {error.Message}", "Error");
            }
        }

        private async Task<bool> VerificacionCamara()
        {
            var valoresNG = "";
            lbx_Codes.Items.Clear();
            lbx_Codes.Visibility = Visibility.Visible;

            var response = await keyence.SendTrigger();
            var resultado = keyence.Formato(response);

            foreach (var item in resultado)
                lbx_Codes.Items.Add(item);

            var ngValores = lbx_Codes.Items.Cast<string>()
                .Where(x => x.Contains("NG"))
                .ToList();

            if (ngValores.Any())
            {
                valoresNG = string.Join(",", ngValores);
                _stopBascula1 = true;
                return true;
            }
            return false;
        }

        private void CamaraCompletada()
        {
            _siguienteManual = false;
            _manual = false;
            _stopBascula1 = false;
            _activarBoton = true;
            _siguienteManualCompletado = false;

            DesactivarSalida(defaultSettings.Piston);

            lbx_Codes.Visibility = Visibility.Hidden;
            _currentStepIndex = 0;
            _accumulatedWeight = 0;
            pieceWeight = 0;
            SetImagesBox();
            contador++;
            lblCompletados.Content = contador;
            codigo = "";
            _inicioZero = true;
            ShowIniciar();
            Dispatcher.Invoke(ShowPruebaCorrecta);
        }

        private void QuickSetup_Btn_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationWindow auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                SequenceStep modArtQS = (SequenceStep)((Button)sender).DataContext;

                PopupQS.Visibility = Visibility.Visible;

                NoArticuloQS_TBox.Text = modArtQS.PartNoParte;
            
                if (FindName(modArtQS.PartPeso) is TextBlock secuencia)
                    PesoBasculaQS_TBox.Text = secuencia.Text;

                ArticuloIDQS_TBox.Text = modArtQS.Id.ToString();
                PesoMinQS_TBox.Text = modArtQS.MinWeight.ToString();
                PesoMaxQS_TBox.Text = modArtQS.MaxWeight.ToString();
                peso_max_qs_text.Text = modArtQS.MaxWeight.ToString();
                peso_min_qs_text.Text = modArtQS.MinWeight.ToString();
            }
        }

        private void SubirVariacion_Btn_Click(object sender, RoutedEventArgs e)
        {
            decimal peso = Convert.ToDecimal(PesoQS_TBox.Text);
            PesoQS_TBox.Text = (peso + Convert.ToDecimal(0.0025)).ToString();
            ActualizarPesosQS();
        }

        private void BajarVariacion_Btn_Click(object sender, RoutedEventArgs e)
        {
            
            decimal peso = Convert.ToDecimal(PesoQS_TBox.Text);
            PesoQS_TBox.Text = (peso - Convert.ToDecimal(0.0025)).ToString();
            ActualizarPesosQS();
        }

        private void SubirPeso_min_Btn_Click(object sender, RoutedEventArgs e)
        {
            BajarSubirPesoQS(true, false);
        }
        private void BajarPeso_min_Btn_Click(object sender, RoutedEventArgs e)
        {
            BajarSubirPesoQS(false, false);
        }

        private void SubirPeso_max_Btn_Click(object sender, RoutedEventArgs e)
        {
            BajarSubirPesoQS(true, true);
        }

        private void BajarPeso_max_Btn_Click(object sender, RoutedEventArgs e)
        {
            BajarSubirPesoQS(false, true);
        }

        private void BajarSubirPesoQS(bool subir, bool max)
        {
            decimal peso_actual = 0;
            decimal peso_aumentar_disminuir = Convert.ToDecimal(PesoQS_TBox.Text);

            if (max)
            {
                peso_actual = Convert.ToDecimal(PesoMaxQS_TBox.Text);
            }
            else
            {
                peso_actual = Convert.ToDecimal(PesoMinQS_TBox.Text);
            }

            if (subir && max)
            {
                peso_actual += peso_aumentar_disminuir;
                PesoMaxQS_TBox.Text = peso_actual.ToString();
            }
            else if (!subir && max)
            {
                peso_actual -= peso_aumentar_disminuir;
                PesoMaxQS_TBox.Text = peso_actual.ToString();
            }
            else if (subir && !max)
            {
                peso_actual += peso_aumentar_disminuir;
                PesoMinQS_TBox.Text = peso_actual.ToString();
            }
            else
            {
                peso_actual -= peso_aumentar_disminuir;
                PesoMinQS_TBox.Text = peso_actual.ToString();
            }

        }

        private void CancelarQS_Btn_Click(object sender, RoutedEventArgs e)
        {
            PopupQS.Visibility = Visibility.Hidden;
        }

        private void AsignarQS_Btn_Click(object sender, RoutedEventArgs e)
        {
            dc_missingpartsEntities db = new dc_missingpartsEntities();

            AsignarQS_Btn.Content = "GUARDANDO...";

            int articuloID = Convert.ToInt16(ArticuloIDQS_TBox.Text);

            Articulo articulo = db.Articulos.Find(articuloID);
            articulo.PesoMin = Convert.ToDouble(PesoMinQS_TBox.Text);
            articulo.PesoMax = Convert.ToDouble(PesoMaxQS_TBox.Text);

            db.Entry(articulo).State = System.Data.Entity.EntityState.Modified;
            db.SaveChanges();
            PopupQS.Visibility = Visibility.Hidden;
        }

        private void ActualizarPesosQS()
        {
            decimal pesoVariacion = Convert.ToDecimal(PesoQS_TBox.Text);
            decimal pesoBasculaQS = Convert.ToDecimal(PesoBasculaQS_TBox.Text.Substring(0, PesoBasculaQS_TBox.Text.Length - 4));
            decimal pesoMax = pesoBasculaQS + pesoVariacion;
            decimal pesoMin = pesoBasculaQS - pesoVariacion;

            PesoMinQS_TBox.Text = Math.Round(pesoMin, 5).ToString();
            PesoMaxQS_TBox.Text = Math.Round(pesoMax, 5).ToString();
        }
    }
}