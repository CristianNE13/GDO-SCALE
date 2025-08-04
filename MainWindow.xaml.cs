using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
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
        private readonly IBasculaFunc bascula;
        private readonly List<Border> borders = new List<Border>();
        private readonly KeyenceTcpClient keyence;
        private readonly List<Label> labels = new List<Label>();
        private readonly string rutaImagenes = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes");
        private readonly List<SequenceStep> valoresBolsas = new List<SequenceStep>();
        private double _accumulatedWeight;
        private bool _activarBoton;
        private AlertWindow _alertWindow;
        private Catalogos _catalogos = new Catalogos();
        private int _consecutiveCount;
        private bool _contieneCCAM;
        private int _currentStepIndex;
        private ErrorMessageWindow _errorWindow;
        private bool _esperandoPickToLight;
        private DispatcherTimer _estadoBasculasTimer;
        private bool _etapa1;
        private bool _ignorarInput;
        private bool _inicioPicks;
        private bool _inicioZero = true;
        private bool _isInitializing;
        private double _lastWeight;
        private bool _manual;
        private bool _pickCompletado;
        private bool _siguienteManual;
        private bool _siguienteManualCompletado;
        private SequenceStep _stepEsperandoPick;
        private bool _stopBascula1 = true;
        private bool _validacion = true;
        private bool _verificacionCompletado;
        private bool _verificacionIndividual;
        private bool _zeroConfirmed;
        private bool botonActivo;
        private string codigo;
        private Configuracion defaultSettings;
        private string fraction;
        private string integrer;
        public IOInterface ioInterface;
        public IOScanner ioScanner;
        private bool ioScannerActivado;
        private Modelo ModeloData;
        private List<SequenceStep> pasosFiltrados;
        private bool pick0Activo;
        private bool pick0Estado;
        private bool pick1Activo;
        private bool pick1Estado;
        private bool pick2Activo;
        private bool pick2Estado;
        private bool pick3Activo;
        private bool pick3Estado;
        private bool pick4Activo;
        private bool pick4Estado;
        private bool pick5Activo;
        private bool pick5Estado;
        private double pieceWeight;
        private ISeaLevelDevice sealevel;
        private bool sensor0Activo;
        private bool sensor0Completado;
        private bool sensor1Activo;
        private bool sensor1Completado;
        private List<SequenceStep> sequence;
        private string zpl;
        private Bitacora bitacora;
        private RegistroBitacora registroBitacora;
        private string directorioBit;
        private double paso0 = 0;
        private double pesoBascula = 0;
        private bool individual = false;
        private bool nosum = false;

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

            borders = new List<Border>
            {
                borderPick1, borderPick2, borderPick3, borderPick4, borderPick5, borderPick6, borderSensor0,
                borderSensor1
            };

            labels = new List<Label>
                { lblPick_1, lblPick_2, lblPick_3, lblPick_4, lblPick_5, lblPick_6, lblVerificar, lblVerificar1 };

            directorioBit = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
        }


        private void Cbox_Modelo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                Grd_Color.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF005288"));
                recValidacion.Visibility = Visibility.Hidden;
                grdValidacion.Visibility = Visibility.Hidden;
                lbx_Codes.Visibility = Visibility.Hidden;

                defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);
                IniciarSealevel();
                if (ioScannerActivado) 
                    OutputsOff();

                ModeloData = Cbox_Modelo.SelectedValue as Modelo ?? throw new Exception("Modelo no reconocido.");

                ProcesarModeloValido(ModeloData.NoModelo);

                if (!ModeloData.UsaCamaraVision)
                {
                    LabelLedPart1_Copy.Visibility = Visibility.Hidden;
                    borderCamaraStatus.Visibility = Visibility.Hidden;
                }
                else
                {
                    LabelLedPart1_Copy.Visibility = Visibility.Visible;
                    borderCamaraStatus.Visibility = Visibility.Visible;
                }

                bitacora = Bitacora.Cargar(directorioBit);

                if( ModeloData.Id > 0)
                {
                    registroBitacora = bitacora.ObtenerRegistro(ModeloData.Id);

                    if (bitacora.PreviousModeloID != ModeloData.Id)
                    {
                        registroBitacora.Completadas = 0;
                        registroBitacora.Rechazos = 0;
                        registroBitacora.UltimaActualizacion = DateTime.Now;

                        bitacora.PreviousModeloID = ModeloData.Id;

                        bitacora.Guardar(directorioBit);
                    }

                    RefreshStatistics();
                }
                
                Cbox_Proceso.SelectedIndex = 0;
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
            //Cbox_Proceso.IsEnabled = true;
            //Cbox_Proceso.Visibility = Visibility.Visible;
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
                foreach (var border in borders)
                {
                    border.Visibility = Visibility.Hidden;
                    labels[borders.IndexOf(border)].Visibility = Visibility.Hidden;
                }

                if (!int.TryParse(Cbox_Proceso.SelectedValue.ToString(), out var procesoSeleccionado))
                    throw new Exception("El valor del proceso seleccionado no es válido.");

                pasosFiltrados = ObtenerValoresProceso(ModeloData.NoModelo, procesoSeleccionado);

                if (pasosFiltrados == null || pasosFiltrados.Count == 0)
                    throw new Exception("No hay pasos definidos para este proceso.");

                ProcesarModeloYProceso();

                SecuenciaASeguir(ModeloData);

                ResetVariables();
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
                        MessageBox.Show($"El programa de visión '{programa}' no es válido.", "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                        return;
                    }

                    if (borderCamaraStatus.Background == Brushes.Red || borderCamaraStatus.Background == Brushes.Yellow)
                    {
                        keyence.StopMonitoring();
                        keyence.Dispose();
                    }

                    var connected = await keyence.ConnectAsync();
                    if (!connected)
                    {
                        MessageBox.Show("No se pudo conectar a la cámara.", "Error");
                        UpdateCameraStatus(0);
                    }

                    if (borderCamaraStatus.Background == Brushes.Red)
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
                var modProceso = ObtenerModeloModProceso(modeloSeleccionado);
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

                var pasosFiltrados = new List<SequenceStep>();

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
                    if (modelo != null) return modelo.ModProceso;
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
                if (ioScannerActivado)
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
                ReiniciarPasos(steps, modeloSeleccionado);
            else
                ShowAlertError($"Modelo '{modeloSeleccionado}' no reconocido.");
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
                
                Grd_Color.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF005288"));
                txbPesoActual.Text = "0.0Kg";
                PesoGeneral.Text = "0.0Kg";

                CerrarAlertas();
                ResetVariables();

                if (ioScannerActivado)
                {
                    OutputsOff();

                    if (defaultSettings.CheckShutOff)
                        ActivarSalida(defaultSettings.ShutOff);
                }

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

                HideAll();

                if (Cbox_Proceso.SelectedValue == null)
                {
                    ShowAlertError("No hay un proceso seleccionado.");
                    return;
                }

                var procesoSeleccionado = int.Parse(Cbox_Proceso.SelectedValue.ToString());
                sequence = ObtenerValoresProceso(modeloSeleccionado, procesoSeleccionado);
                ProcesarModeloValido(ModeloData.NoModelo);
                ProcesarModeloYProceso();
                SetValuesEtapas(sequence, 1);
                SetImagesBox();
                SecuenciaASeguir(ModeloData);
                Cbox_Proceso.SelectedIndex = procesoSeleccionado - 1;
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al procesar el reset del modelo: {ex.Message}");
            }
        }

        void ResetVariables()
        {
            nosum = false;
            _siguienteManual = false;
            _siguienteManualCompletado =false;
            sensor0Completado = false;
            sensor1Completado = false;
            _verificacionCompletado = false;          
            individual = false;          
            _stopBascula1 = false;          
            _manual = false;
            _zeroConfirmed = false;
            _inicioZero = true;
            _activarBoton = true;
            pick0Estado = true;
            pick1Estado = true;
            pick2Estado = true;
            pick3Estado = true;
            pick4Estado = true;
            pick5Estado = true;
            _pickCompletado = false;
            _esperandoPickToLight = false;    
            _verificacionIndividual = false;
            _etapa1 = true;
            _consecutiveCount = 0;
            _accumulatedWeight = 0;
            _currentStepIndex = 0;
            lblProgreso.Content = 0;
            pieceWeight = 0;
            codigo = "";
            valoresBolsas.Clear();
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

        private async void btnIniciarZero_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var lista = ModeloData.ProgramaVision.Split(',');

                var connected = await keyence.ConnectAsync();
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

                    await ShowMensaje($"Errores: {valoresNG}", Brushes.Beige, 3000);
                    return;
                }

                recValidacion.Visibility = Visibility.Hidden;
                grdValidacion.Visibility = Visibility.Hidden;
                lbx_Codes.Visibility = Visibility.Hidden;
                ioInterface.WriteSingleOutput(6, true);

                var currentStep = pasosFiltrados[_currentStepIndex];
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al inicializar báscula: {ex.Message}", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task ActivarCamaraValidacion()
        {
            ShowAlertCamara();
            _stopBascula1 = true;
            InspeccionarValidacionFunc();
            await Task.Delay(300); // Delay de 300ms para dejar que se realice la inspeccion.
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
                var lista = ModeloData.ProgramaVision.Split(',');
                bool error;

                if (_inicioZero)
                {
                    if (ModeloData.UsaCamaraVision)
                    {
                        await keyence.ChangeProgram(int.Parse(lista.First()));

                        error = await VerificacionCamara();

                        if (error)
                        {
                            _ = ShowMensaje("Error, verificar imagen de camara", Brushes.Beige, 3000);
                            return;
                        }

                        recValidacion.Visibility = Visibility.Hidden;
                        grdValidacion.Visibility = Visibility.Hidden;
                        lbx_Codes.Visibility = Visibility.Hidden;
                    }

                    ActivarSalida(defaultSettings.Piston);

                    _stopBascula1 = false;
                    _activarBoton = false;
                    _inicioZero = false;
                    _zeroConfirmed = false;
                    _pickCompletado = false;
                    _inicioPicks = false;

                    if (ModeloData.UsaCamaraVision)
                    {
                        _ = ShowMensaje("INSPECCION CORRECTA", Brushes.Green, 1500);
                    }

                    else if (!ModeloData.UsaCamaraVision)
                    {
                        recValidacion.Visibility = Visibility.Hidden;
                        grdValidacion.Visibility = Visibility.Hidden;
                        lbx_Codes.Visibility = Visibility.Hidden;
                        _ = ShowMensaje("ESPERA A PESO APROX 0.0KG", Brushes.Green, 1500);
                    }
                }

                else
                {
                    if ((_siguienteManual && !_siguienteManualCompletado) ||
                        (!_siguienteManual && !_siguienteManualCompletado))
                        await keyence.ChangeProgram(int.Parse(lista[1]));

                    else if (_manual)
                        await keyence.ChangeProgram(int.Parse(lista.Last()));

                    error = await VerificacionCamara();

                    if (error)
                    {
                        _ = ShowMensaje("Error, verificar imagen de camara", Brushes.Beige, 3000);
                        return;
                    }

                    if (_manual)
                    {
                        var currentStep = pasosFiltrados[_currentStepIndex];

                        LogCompleteStep(currentStep, "OK", codigo);

                        CamaraCompletada();
                        return;
                    }

                    if (_siguienteManual)
                    {
                        _ = ShowMensaje("INSPECCION CORRECTA", Brushes.Green, 1500);

                        lbx_Codes.Visibility = Visibility.Hidden;

                        var currentStep = pasosFiltrados[_currentStepIndex];

                        if (individual && !nosum)
                        {
                            _currentStepIndex += 1;
                            currentStep = pasosFiltrados[_currentStepIndex];
                            nosum = true;
                        }

                        else if (!individual)
                            currentStep = pasosFiltrados[_currentStepIndex];

                        if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock
                                cantidadTextBlock)) return;
                        if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) ||
                            !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

                        _stopBascula1 = false;
                        _siguienteManualCompletado = true;

                        if (ModeloData.UsaPick2Light)
                        {
                            ActivarSalida(_currentStepIndex + defaultSettings.OutputPick2L0);

                            borders[_currentStepIndex + defaultSettings.OutputPick2L0].Visibility = Visibility.Visible;
                            labels[_currentStepIndex + defaultSettings.OutputPick2L0].Visibility = Visibility.Visible;

                            _esperandoPickToLight = true;

                            switch (_currentStepIndex)
                            {
                                case 0:
                                    pick0Estado = true;
                                    break;
                                case 1:
                                    pick1Estado = true;
                                    break;
                                case 2:
                                    pick2Estado = true;
                                    break;
                                case 3:
                                    pick3Estado = true;
                                    break;
                                case 4:
                                    pick4Estado = true;
                                    break;
                                case 5:
                                    pick5Estado = true;
                                    break;
                            }
                        }

                        _verificacionIndividual = true;
                        _activarBoton = false;

                        recValidacion.Visibility = Visibility.Hidden;
                        grdValidacion.Visibility = Visibility.Hidden;
                        lbx_Codes.Visibility = Visibility.Hidden;
                    }

                    else
                    {
                        CamaraCompletada();
                    }
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

            if (resultado == null ||
                (resultado.Count == 1 && resultado.Contains("No hay conexión activa con la cámara.")))
            {
                lbx_Codes.Items.Add("SIN RESPUESTA");
                _stopBascula1 = true;
                await ShowMensaje("NO SE CONECTO CON LA CAMARA, REINICIAR SISTEMA", Brushes.Red, 5000);
                await Task.Delay(5000);
                return true;
            }

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
            Grd_Color.Background = (SolidColorBrush)(new BrushConverter().ConvertFrom("#FF005288"));
            lbx_Codes.Visibility = Visibility.Hidden;
            lblCompletados.Content = registroBitacora.Completadas;
            bitacora.Guardar(directorioBit);
            ResetVariables();
            DesactivarSalida(defaultSettings.Piston);
            SetImagesBox();
            ShowIniciar();
            Dispatcher.Invoke(ShowPruebaCorrecta);
        }

        private void QuickSetup_Btn_Click(object sender, RoutedEventArgs e)
        {
            var auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                var modArtQS = (SequenceStep)((Button)sender).DataContext;

                PopupQS.Visibility = Visibility.Visible;

                NoArticuloQS_TBox.Text = modArtQS.PartNoParte;

                if (FindName(modArtQS.PartPeso) is TextBlock secuencia)
                { 
                    PesoBasculaQS_TBox.Text = secuencia.Text;
                    if (secuencia.Text == "0.0Kgs")
                    {
                        var lista = PesoGeneral.Text.Split(' ');
                        PesoBasculaQS_TBox.Text = lista[1]+" kg";
                    }
                }


                ArticuloIDQS_TBox.Text = modArtQS.Id.ToString();
                PesoMinQS_TBox.Text = modArtQS.MinWeight.ToString();
                PesoMaxQS_TBox.Text = modArtQS.MaxWeight.ToString();
                peso_max_qs_text.Text = modArtQS.MaxWeight.ToString();
                peso_min_qs_text.Text = modArtQS.MinWeight.ToString();
            }
        }

        private void SubirVariacion_Btn_Click(object sender, RoutedEventArgs e)
        {
            var peso = Convert.ToDecimal(PesoQS_TBox.Text);
            PesoQS_TBox.Text = (peso + Convert.ToDecimal(0.0025)).ToString();
            ActualizarPesosQS();
        }

        private void BajarVariacion_Btn_Click(object sender, RoutedEventArgs e)
        {
            var peso = Convert.ToDecimal(PesoQS_TBox.Text);
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
            var peso_aumentar_disminuir = Convert.ToDecimal(PesoQS_TBox.Text);

            if (max)
                peso_actual = Convert.ToDecimal(PesoMaxQS_TBox.Text);
            else
                peso_actual = Convert.ToDecimal(PesoMinQS_TBox.Text);

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
            var db = new dc_missingpartsEntities();

            AsignarQS_Btn.Content = "GUARDANDO...";

            int articuloID = Convert.ToInt16(ArticuloIDQS_TBox.Text);

            var articulo = db.Articulos.Find(articuloID);
            articulo.PesoMin = Convert.ToDouble(PesoMinQS_TBox.Text);
            articulo.PesoMax = Convert.ToDouble(PesoMaxQS_TBox.Text);

            db.Entry(articulo).State = EntityState.Modified;
            db.SaveChanges();
            PopupQS.Visibility = Visibility.Hidden;
        }

        private void ActualizarPesosQS()
        {
            var pesoVariacion = Convert.ToDecimal(PesoQS_TBox.Text);
            var pesoBasculaQS =
                Convert.ToDecimal(PesoBasculaQS_TBox.Text.Substring(0, PesoBasculaQS_TBox.Text.Length - 4));
            var pesoMax = pesoBasculaQS + pesoVariacion;
            var pesoMin = pesoBasculaQS - pesoVariacion;

            PesoMinQS_TBox.Text = Math.Round(pesoMin, 5).ToString();
            PesoMaxQS_TBox.Text = Math.Round(pesoMax, 5).ToString();
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

                    imageControl.Source = File.Exists(imagePath)
                        ? new BitmapImage(new Uri(imagePath, UriKind.Absolute))
                        : null;
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

                    imageControl.Source = File.Exists(imagePath)
                        ? new BitmapImage(new Uri(imagePath, UriKind.Absolute))
                        : null;
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
                step.PartOrden = $"{i + 1}";


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

                if (FindName($"Part_Indicator{counter}") is Rectangle indicador &&
                    FindName($"Part_Peso{counter}") is TextBlock peso)
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
                    var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes",
                        $"{step.PartNoParte}.PNG");

                    if (FindName($"Part_Cantidad{counter}") is TextBlock partCantidad)
                        partCantidad.Text = step.PartCantidad;

                    if (FindName($"Part{counter}") is Grid part)
                        part.Visibility = Visibility.Visible;

                    imageControl.Source = File.Exists(imagePath)
                        ? new BitmapImage(new Uri(imagePath, UriKind.Absolute))
                        : null;
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
            _stepEsperandoPick = current;

            lblPesoArt.Text = "TOMAR EL ARTICULO";

            txbArticulo.Text = $"PICK2LIGHT-{int.Parse(current.PartOrden) - 1}";

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
            txbPesoActual.Text = current.PartNoParte.ToUpper();
            ;

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
            SalidaPick2Orden(current.PartOrden, false);
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
            _estadoBasculasTimer.Tick += VerificarEstadoBasculasPick2;
            _estadoBasculasTimer.Start();
        }

        private void VerificarEstadoBasculasPick2(object sender, EventArgs e)
        {
            borderBascula1.Background = _stopBascula1 ? Brushes.Red : Brushes.Green;
            borderPick1.Background = pick0Estado ? Brushes.Red : Brushes.Green;
            borderPick2.Background = pick1Estado ? Brushes.Red : Brushes.Green;
            borderPick3.Background = pick2Estado ? Brushes.Red : Brushes.Green;
            borderPick4.Background = pick3Estado ? Brushes.Red : Brushes.Green;
            borderPick5.Background = pick4Estado ? Brushes.Red : Brushes.Green;
            borderPick6.Background = pick5Estado ? Brushes.Red : Brushes.Green;
            borderSensor0.Background = sensor0Completado ? Brushes.Green : Brushes.Red;
            borderSensor1.Background = sensor1Completado ? Brushes.Green : Brushes.Red;
        }

        #endregion

        #region SEALEVEL

        private void SalidaPick2Orden(string partOrden, bool activo)
        {
            var start = defaultSettings.InputPick2L0;

            if (int.TryParse(partOrden, out var offset) && offset > 0)
            {
                var index = start + (offset - 1);

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

        private async void IniciarSealevel()
        {
            if (ioScannerActivado)
                return;

            try
            {
                if (defaultSettings.CheckSealevelEthernet)
                {
                    var eth = new SeaLevelEthernet(defaultSettings.SealevelIP);
                    bool conectado = await Task.Run(() => eth.Connect()); // no bloquea la UI

                    if (!conectado)
                    {
                        ShowAlertError("NO SE LOGRO CONECTAR CON SEALEVEL ETHERNET");
                        return;
                    }

                    sealevel = eth;
                }
                else
                {
                    var serial = new SeaLevel(247, defaultSettings.PuertoSealevel, defaultSettings.BaudRateSea);
                    serial.Open();

                    if (!serial.IsOpen)
                    {
                        ShowAlertError("NO SE LOGRO CONECTAR CON SEALEVEL SERIAL");
                        return;
                    }

                    sealevel = serial;
                }

                ioInterface = new IOInterface(sealevel);
                ioScanner = new IOScanner(ioInterface);
                ioScanner.Tick += OnScannerTick;
                ioScanner.Start();
                ioScannerActivado = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al iniciar SeaLevel: {ex.Message}", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                ioScannerActivado = false;
            }
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
                var sensor0 = (inputsState & (1 << defaultSettings.InputSensor0)) != 0;
                var sensor1 = (inputsState & (1 << defaultSettings.InputSensor1)) != 0;
                var botonInspeccion = (inputsState & (1 << defaultSettings.InputBoton)) != 0;
                var pick0 = (inputsState & (1 << defaultSettings.InputPick2L0)) != 0;
                var pick1 = (inputsState & (1 << (defaultSettings.InputPick2L0 + 1))) != 0;
                var pick2 = (inputsState & (1 << (defaultSettings.InputPick2L0 + 2))) != 0;
                var pick3 = (inputsState & (1 << (defaultSettings.InputPick2L0 + 3))) != 0;
                var pick4 = (inputsState & (1 << (defaultSettings.InputPick2L0 + 4))) != 0;
                var pick5 = (inputsState & (1 << (defaultSettings.InputPick2L0 + 5))) != 0;


                if (sensor0 && !sensor0Activo) sensor0Completado = true;

                if (sensor1 && !sensor1Activo) sensor1Completado = true;

                if (sensor0Completado && sensor1Completado) _verificacionCompletado = true;

                if (botonInspeccion && !botonActivo && _activarBoton)
                {
                    if (_ignorarInput)
                        return;

                    _ignorarInput = true;
                    InspeccionarValidacionFunc();

                    await Task.Delay(2000);
                    _ignorarInput = false;
                }

                if (_esperandoPickToLight && !pick0Activo && pick0 && pick0Estado)
                {
                    pick0Estado = false;
                    await Task.Delay(2000);
                    DesactivarSalida(defaultSettings.InputPick2L0);
                }

                if (_esperandoPickToLight && !pick1Activo && pick1 && pick1Estado)
                {
                    pick1Estado = false;
                    await Task.Delay(2000);
                    DesactivarSalida(defaultSettings.InputPick2L0 + 1);
                }

                if (_esperandoPickToLight && !pick2Activo && pick2 && pick2Estado)
                {
                    pick2Estado = false;
                    await Task.Delay(2000);
                    DesactivarSalida(defaultSettings.InputPick2L0 + 2);
                }

                if (_esperandoPickToLight && !pick3Activo && pick3 && pick3Estado)
                {
                    pick3Estado = false;
                    await Task.Delay(2000);
                    DesactivarSalida(defaultSettings.InputPick2L0 + 3);
                }

                if (_esperandoPickToLight && !pick4Activo && pick4 && pick4Estado)
                {
                    pick4Estado = false;
                    await Task.Delay(2000);
                    DesactivarSalida(defaultSettings.InputPick2L0 + 4);
                }

                if (_esperandoPickToLight && !pick5Activo && pick5 && pick5Estado)
                {
                    pick5Estado = false;
                    await Task.Delay(2000);
                    DesactivarSalida(defaultSettings.InputPick2L0 + 5);
                }

                if (!pick0Estado && !pick1Estado && !pick2Estado && !pick3Estado && !pick4Estado && !pick5Estado &&
                    !_contieneCCAM && _inicioPicks)
                {
                    _esperandoPickToLight = false;
                    _pickCompletado = true;
                }

                if (_contieneCCAM && _esperandoPickToLight)
                {
                    var ccamIndex = _currentStepIndex + defaultSettings.InputPick2L0;
                    var ccamSensor = (inputsState & (1 << ccamIndex)) != 0;

                    if (ccamSensor)
                    {
                        DesactivarSalida(ccamIndex);

                        switch (_currentStepIndex)
                        {
                            case 0:
                                pick0Estado = false;
                                break;
                            case 1:
                                pick1Estado = false;
                                break;
                            case 2:
                                pick2Estado = false;
                                break;
                            case 3:
                                pick3Estado = false;
                                break;
                            case 4:
                                pick4Estado = false;
                                break;
                            case 5:
                                pick5Estado = false;
                                break;
                        }

                        _contieneCCAM = false;
                        _esperandoPickToLight = false;
                    }
                }


                sensor0Activo = sensor0;
                sensor1Activo = sensor1;
                botonActivo = botonInspeccion;
                pick0Activo = pick0;
                pick1Activo = pick1;
                pick2Activo = pick2;
                pick3Activo = pick3;
                pick4Activo = pick4;
                pick5Activo = pick5;
            });
        }

        #endregion

        #region LOGS
        
        private void RefreshStatistics()
        {
            if (registroBitacora == null)
            {
                lblCompletados.Content = "0";
                lblCompletados.IsEnabled = false;
            }
            else
            {
                lblCompletados.Content = registroBitacora.Completadas.ToString("N0", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                lblCompletados.IsEnabled = true;
            }
        }

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

        private void LogCompleteStep(SequenceStep step, string resultado, string matchedTag)
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var registro = new Completado
                    {
                        Fecha = DateTime.Now,
                        NoParte = ModeloData.NoModelo,
                        ModProceso = step.ModProceso,
                        Proceso = step.PartProceso,
                        PesoDetectado = double.TryParse(step.DetectedWeight, out var peso) ? peso : 0,
                        Estado = resultado,
                        Tag = matchedTag ?? ""
                    };

                    db.Completados.Add(registro);
                    db.SaveChanges();

                    if (registro.Tag != "" && registroBitacora != null) 
                        registroBitacora.Completadas++;

                    bitacora.Guardar(directorioBit);
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

                    registroBitacora.Rechazos++;
                    bitacora.Guardar(directorioBit);
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
                var error = false;

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
                    if (error) _stopBascula1 = true;
                }
            });
        }

        private void Bascula1_OnDataReady(object sender, BasculaEventArgs e)
        {
            if (_stopBascula1) return;

            if (_inicioZero)
                return;

            Dispatcher.Invoke(() =>
            {
                var weight = e.Value;
                pesoBascula = weight;
                var isStable = e.IsStable;
                PesoGeneral.Text = $"Peso: {weight-paso0:F5} kg";
                if (isStable)
                {
                    if (Math.Abs(weight - _lastWeight) < 0.0001)
                        _consecutiveCount++;
                    else
                        _consecutiveCount = 1;

                    _lastWeight = weight;


                    if (_currentStepIndex == 0 && !_inicioZero && !_zeroConfirmed)
                    {
                        if (Math.Abs(weight - paso0) <= 0.0025)
                        {
                            if (_consecutiveCount >= 1)
                            {
                                _zeroConfirmed = true;
                                _consecutiveCount = 0;

                                weight = weight - paso0;

                                var verificarArticulo = pasosFiltrados.Where(v => v.PartNoParte.Contains("VERIFY"));
                                var verificarArticulo2 = pasosFiltrados.Where(v => v.PartNoParte.Contains("VSEN2"));

                                if (!verificarArticulo.Any() && !_verificacionCompletado)
                                {
                                    sensor0Completado = true;
                                    borderSensor0.Visibility = Visibility.Hidden;
                                    lblVerificar.Visibility = Visibility.Hidden;
                                }

                                if (!verificarArticulo2.Any() && !_verificacionCompletado)
                                {
                                    sensor1Completado = true;
                                    borderSensor1.Visibility = Visibility.Hidden;
                                    lblVerificar1.Visibility = Visibility.Hidden;

                                    if (sensor0Completado && sensor1Completado)
                                        _verificacionCompletado = true;
                                }

                                if (Math.Abs(weight) <= 0.0025 && verificarArticulo.Any() && !_verificacionCompletado)
                                {
                                    _verificacionCompletado = false;
                                    borderSensor0.Visibility = Visibility.Visible;
                                    lblVerificar.Visibility = Visibility.Visible;
                                }

                                if (Math.Abs(weight) <= 0.0025 && verificarArticulo2.Any() && !_verificacionCompletado)
                                {
                                    sensor1Completado = false;
                                    _verificacionCompletado = false;
                                    borderSensor1.Visibility = Visibility.Visible;
                                    lblVerificar1.Visibility = Visibility.Visible;
                                }
                                Grd_Color.Background = Brushes.ForestGreen;

                                //if (!ModeloData.UsaPick2Light)
                                //    _ = ShowMensaje("PUEDE INICIAR", Brushes.Beige, 1000);

                                ProcessStableWeight(weight);
                            }
                        }
                        else
                        {
                            if (_consecutiveCount >= 1)
                            {
                                paso0 = weight;
                                _consecutiveCount = 0;
                            }
                        }


                    }

                    if (_zeroConfirmed && _consecutiveCount == 1 && ModeloData.UsaCamaraVision)
                    {
                        weight = weight - paso0;
                        ProcessStableWeight(weight);
                    }


                    if (_zeroConfirmed && _consecutiveCount == 1 && !ModeloData.UsaCamaraVision)
                    {
                        weight = weight - paso0;
                        ProcessStableWeightNoCam(weight);
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
                {
                    ShowAlertError("Error al cargar la secuencia, secuencia inexistente.");
                }
            }
            catch (Exception e)
            {
                ShowAlertError(e.ToString());
                throw;
            }
        }

        private async void ProcessStableWeight(double currentWeight)
        {
            if (!_zeroConfirmed) return;

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

            if (ModeloData.UsaPick2Light && !_esperandoPickToLight && !_contieneCCAM && !_pickCompletado)
            {
                _pickCompletado = false;
                pick0Estado = false;
                pick1Estado = false;
                pick2Estado = false;
                pick3Estado = false;
                pick4Estado = false;
                pick5Estado = false;
                _esperandoPickToLight = true;
                _inicioPicks = true;

                var listaSinCCAM = pasosAMostrar
                    .Where(p => !p.PartNoParte.Contains("CCAM"))
                    .ToList();

                borderPick1.Visibility = Visibility.Hidden;
                borderPick2.Visibility = Visibility.Hidden;
                borderPick3.Visibility = Visibility.Hidden;
                borderPick4.Visibility = Visibility.Hidden;
                borderPick5.Visibility = Visibility.Hidden;
                borderPick6.Visibility = Visibility.Hidden;
                lblPick_1.Visibility = Visibility.Hidden;
                lblPick_2.Visibility = Visibility.Hidden;
                lblPick_3.Visibility = Visibility.Hidden;
                lblPick_4.Visibility = Visibility.Hidden;
                lblPick_5.Visibility = Visibility.Hidden;
                lblPick_6.Visibility = Visibility.Hidden;

                var pickNumbers = listaSinCCAM
                    .Select(x => int.TryParse(x.PartOrden, out var orden) ? orden : -1)
                    .Where(orden => orden != -1)
                    .Select(orden => defaultSettings.InputPick2L0 + (orden - 1))
                    .ToList();

                foreach (var encender in pickNumbers)
                {
                    var relativeIndex = encender - defaultSettings.InputPick2L0;
                    var salidaPick = defaultSettings.InputPick2L0 + relativeIndex;

                    ActivarSalida(salidaPick);

                    switch (relativeIndex)
                    {
                        case 0:
                            pick0Estado = true;
                            break;
                        case 1:
                            pick1Estado = true;
                            break;
                        case 2:
                            pick2Estado = true;
                            break;
                        case 3:
                            pick3Estado = true;
                            break;
                        case 4:
                            pick4Estado = true;
                            break;
                        case 5:
                            pick5Estado = true;
                            break;
                    }
                }

                foreach (var paso in listaSinCCAM)
                    if (int.TryParse(paso.PartOrden, out var orden))
                    {
                        var index = orden - 1;
                        if (index >= 0 && index < borders.Count)
                        {
                            borders[index].Visibility = Visibility.Visible;
                            labels[index].Visibility = Visibility.Visible;
                        }
                    }
            }

            if (!ModeloData.UsaPick2Light)
            {
                _pickCompletado = true;
                _esperandoPickToLight = false;
                _inicioPicks = false;
            }

            var ccam = pasosAMostrar.Find(x => x.PartNoParte.Contains("CCAM"));

            if (ccam == null && !_verificacionIndividual && !_esperandoPickToLight && _verificacionCompletado)
            {
                var minTotal = pasosAMostrar.Sum(p => p.MinWeight);
                var maxTotal = pasosAMostrar.Sum(p => p.MaxWeight);

                if (currentWeight >= minTotal && currentWeight <= maxTotal && !_esperandoPickToLight)
                {
                    foreach (var step in pasosAMostrar)
                    {
                        if (!(FindName($"Part_Cantidad{int.Parse(step.PartOrden) - 1}") is TextBlock
                                cantidadTotalTxbBlock)) continue;
                        if (!(FindName(step.PartIndicator) is Rectangle totalIndicador) ||
                            !(FindName(step.PartPeso) is TextBlock pesoTotalTextBlock)) continue;

                        totalIndicador.Fill = Brushes.Green;
                        cantidadTotalTxbBlock.Text = "0";
                        pesoTotalTextBlock.Text = "OK";

                        step.DetectedWeight = "OK";

                        if (step == pasosAMostrar.Last())
                        {
                            //step.PartNoParte = ModeloData.NoModelo;
                            step.DetectedWeight = currentWeight.ToString();
                            (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(ModeloData.NoModelo);
                            codigo = $"{integrer}.{fraction}";
                            LogCompleteStep(step, "PESO TOTAL OK", codigo);
                        }
                    }

                    if (ModeloData.UsaCamaraVision)
                    {
                        ActivarCamaraValidacion();
                        return;
                    }

                    DesactivarSalida(defaultSettings.Piston);
                    ProcesarResetModelo();
                    lbx_Codes.Visibility = Visibility.Hidden;
                    SetImagesBox();
                    lblCompletados.Content = registroBitacora.Completadas;
                    ShowIniciar();
                    Dispatcher.Invoke(ShowPruebaCorrecta);
                    return;
                }
            }

            else if (ccam != null && !_verificacionIndividual && !_esperandoPickToLight && _verificacionCompletado)
            {
                var minTotalCCAM = pasosAMostrar
                    .Where(p => !p.PartNoParte.Contains("CCAM"))
                    .Sum(p => p.MinWeight);

                var maxTotalCCAM = pasosAMostrar
                    .Where(p => !p.PartNoParte.Contains("CCAM"))
                    .Sum(p => p.MaxWeight);

                var listaSinCCAM = pasosAMostrar
                    .Where(p => !p.PartNoParte.Contains("CCAM"))
                    .ToList();

                if (currentWeight >= minTotalCCAM && currentWeight <= maxTotalCCAM)
                {
                    foreach (var step in listaSinCCAM)
                    {
                        if (!(FindName($"Part_Cantidad{int.Parse(step.PartOrden) - 1}") is TextBlock
                                cantidadTotalTxbBlock)) continue;
                        if (!(FindName(step.PartIndicator) is Rectangle totalIndicador) ||
                            !(FindName(step.PartPeso) is TextBlock pesoTotalTextBlock)) continue;

                        totalIndicador.Fill = Brushes.Green;
                        cantidadTotalTxbBlock.Text = "0";
                        pesoTotalTextBlock.Text = "OK";

                        step.DetectedWeight = "OK";

                        if (step == listaSinCCAM.Last())
                        {
                            step.DetectedWeight = currentWeight.ToString();
                            LogCompleteStep(step, "PESO OK", "");
                        }
                    }

                    _accumulatedWeight = currentWeight;

                    var ccamIndex = pasosAMostrar.IndexOf(ccam);
                    _currentStepIndex = ccamIndex;

                    if (ModeloData.UsaCamaraVision)
                    {
                        _siguienteManual = true;
                        ActivarCamaraValidacion();
                        return;
                    }

                    _verificacionIndividual = true;
                }
            }

            if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock cantidadTextBlock))
                return;

            ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight,
                _validacion, int.Parse(cantidadTextBlock.Text));

            if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) ||
                !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

            if (!currentStep.PartNoParte.Contains("CCAM") && pieceWeight >= currentStep.MinWeight &&
                pieceWeight <= currentStep.MaxWeight)
            {
                var listaSinCCAM = pasosAMostrar
                    .Where(p => !p.PartNoParte.Contains("CCAM"))
                    .ToList();

                if (_currentStepIndex + 1 >= listaSinCCAM.Count && _esperandoPickToLight)
                    return;

                if ((_currentStepIndex + 1 >= pasosFiltrados.Count && _esperandoPickToLight) ||
                    (_currentStepIndex + 1 >= pasosFiltrados.Count && !_verificacionCompletado))
                    return;

                CompleteCurrentStep(currentStep, indicator, pesoTextBlock, cantidadTextBlock, currentWeight);

                if (_currentStepIndex + 1 < pasosFiltrados.Count)
                {
                    var siguientepaso = pasosFiltrados[_currentStepIndex + 1];

                    if (siguientepaso.PartNoParte.Contains("CCAM") && ModeloData.UsaCamaraVision)
                    {
                        _siguienteManual = true;
                        individual = true;
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

                    ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                        pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));
                    return;
                }

                if (ModeloData.UsaCamaraVision)
                {
                    var lista = ModeloData.ProgramaVision.Split(',');

                    await keyence.ChangeProgram(int.Parse(lista[1]));

                    ActivarCamaraValidacion();
                    return;
                }

                lbx_Codes.Visibility = Visibility.Hidden;
                DesactivarSalida(defaultSettings.Piston);
                ProcesarResetModelo();
                SetImagesBox();
                lblCompletados.Content = registroBitacora.Completadas;
                ShowIniciar();
                Dispatcher.Invoke(ShowPruebaCorrecta);
                return;
            }

            if (currentStep.PartNoParte.Contains("CCAM") && pieceWeight >= currentStep.MinWeight &&
                pieceWeight <= currentStep.MaxWeight && !_esperandoPickToLight && ModeloData.UsaCamaraVision)
            {
                indicator.Fill = Brushes.Green;
                pesoTextBlock.Text = $"{pieceWeight:F5} kg";
                currentStep.DetectedWeight = currentWeight.ToString();
                (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(ModeloData.NoModelo);
                codigo = $"{integrer}.{fraction}";
                _activarBoton = true;
                _manual = true;
                ActivarCamaraValidacion();
                return;
            }

            if (pieceWeight >= currentStep.MinWeight && pieceWeight <= currentStep.MaxWeight &&
                !ModeloData.UsaCamaraVision)
            {
                if (_currentStepIndex + 1 >= pasosFiltrados.Count && _esperandoPickToLight && ModeloData.UsaPick2Light)
                    return;

                CompleteCurrentStep(currentStep, indicator, pesoTextBlock, cantidadTextBlock, currentWeight);

                if (_currentStepIndex + 1 < pasosFiltrados.Count)
                {
                    var siguientepaso = pasosFiltrados[_currentStepIndex + 1];

                    if (siguientepaso.PartNoParte.Contains("CCAM") && ModeloData.UsaCamaraVision)
                    {
                        _siguienteManual = true;
                        individual = true;
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

                    ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                        pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));
                    return;
                }

                sensor0Completado = false;
                sensor1Completado = false;
                _verificacionCompletado = false;
                sensor0Completado = false;
                sensor1Completado = false;
                _siguienteManual = false;
                _manual = false;
                _stopBascula1 = false;
                _activarBoton = true;
                _siguienteManualCompletado = false;
                _zeroConfirmed = false;
                _inicioPicks = false;
                _pickCompletado = false;

                DesactivarSalida(defaultSettings.Piston);

                lbx_Codes.Visibility = Visibility.Hidden;
                _currentStepIndex = 0;
                _accumulatedWeight = 0;
                pieceWeight = 0;
                SetImagesBox();
                lblCompletados.Content = registroBitacora.Completadas;
                codigo = "";
                _inicioZero = true;
                ShowIniciar();
                Dispatcher.Invoke(ShowPruebaCorrecta);
                return;
            }

            indicator.Fill = Brushes.Red;
            pesoTextBlock.Text = $"{pieceWeight:F5} kg";
            currentStep.DetectedWeight = currentWeight.ToString("F5");
        }

        private void ProcessStableWeightNoCam(double currentWeight)
        {
            if (!_zeroConfirmed) return;

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

            var verificarArticulo = pasosAMostrar.Where(v => v.PartNoParte.Contains("VERIFY"));
            var verificarArticulo2 = pasosAMostrar.Where(v => v.PartNoParte.Contains("VSEN2"));

            if (Math.Abs(currentWeight) <= 0.0025 && verificarArticulo != null && !_verificacionCompletado)
            {
                _verificacionCompletado = false;
                borderSensor0.Visibility = Visibility.Visible;
                lblVerificar.Visibility = Visibility.Visible;
            }

            if (!verificarArticulo.Any())
            {
                sensor0Completado = true;
                borderSensor0.Visibility = Visibility.Hidden;
                lblVerificar.Visibility = Visibility.Hidden;
            }

            if (Math.Abs(currentWeight) <= 0.0025 && verificarArticulo2 != null && !_verificacionCompletado)
            {
                sensor1Completado = false;
                _verificacionCompletado = false;
                borderSensor1.Visibility = Visibility.Visible;
                lblVerificar1.Visibility = Visibility.Visible;
            }

            if (!verificarArticulo2.Any())
            {
                sensor1Completado = true;
                borderSensor1.Visibility = Visibility.Hidden;
                lblVerificar1.Visibility = Visibility.Hidden;
            }

            if (ModeloData.UsaPick2Light && !_esperandoPickToLight && !_pickCompletado)
            {
                _pickCompletado = false;
                pick0Estado = false;
                pick1Estado = false;
                pick2Estado = false;
                pick3Estado = false;
                pick4Estado = false;
                pick5Estado = false;
                _esperandoPickToLight = true;
                _inicioPicks = true;

                var listaSinCCAM = pasosAMostrar
                    .ToList();

                borderPick1.Visibility = Visibility.Hidden;
                borderPick2.Visibility = Visibility.Hidden;
                borderPick3.Visibility = Visibility.Hidden;
                borderPick4.Visibility = Visibility.Hidden;
                borderPick5.Visibility = Visibility.Hidden;
                borderPick6.Visibility = Visibility.Hidden;
                lblPick_1.Visibility = Visibility.Hidden;
                lblPick_2.Visibility = Visibility.Hidden;
                lblPick_3.Visibility = Visibility.Hidden;
                lblPick_4.Visibility = Visibility.Hidden;
                lblPick_5.Visibility = Visibility.Hidden;
                lblPick_6.Visibility = Visibility.Hidden;

                var pickNumbers = listaSinCCAM
                    .Select(x => int.TryParse(x.PartOrden, out var orden) ? orden : -1)
                    .Where(orden => orden != -1)
                    .Select(orden => defaultSettings.InputPick2L0 + (orden - 1))
                    .ToList();

                foreach (var encender in pickNumbers)
                {
                    var relativeIndex = encender - defaultSettings.InputPick2L0;
                    var salidaPick = defaultSettings.InputPick2L0 + relativeIndex;

                    ActivarSalida(salidaPick);

                    switch (relativeIndex)
                    {
                        case 0:
                            pick0Estado = true;
                            break;
                        case 1:
                            pick1Estado = true;
                            break;
                        case 2:
                            pick2Estado = true;
                            break;
                        case 3:
                            pick3Estado = true;
                            break;
                        case 4:
                            pick4Estado = true;
                            break;
                        case 5:
                            pick5Estado = true;
                            break;
                    }
                }

                foreach (var paso in listaSinCCAM)
                    if (int.TryParse(paso.PartOrden, out var orden))
                    {
                        var index = orden - 1;
                        if (index >= 0 && index < borders.Count)
                        {
                            borders[index].Visibility = Visibility.Visible;
                            labels[index].Visibility = Visibility.Visible;
                        }
                    }
            }

            if (!ModeloData.UsaPick2Light)
            {
                _pickCompletado = true;
                _esperandoPickToLight = false;
                _inicioPicks = false;
            }

            if (!_esperandoPickToLight)
            {
                var minTotal = pasosAMostrar.Sum(p => p.MinWeight);
                var maxTotal = pasosAMostrar.Sum(p => p.MaxWeight);

                if (currentWeight >= minTotal && currentWeight <= maxTotal && _verificacionCompletado)
                {
                    foreach (var step in pasosAMostrar)
                    {
                        if (!(FindName($"Part_Cantidad{int.Parse(step.PartOrden) - 1}") is TextBlock
                                cantidadTotalTxbBlock)) continue;
                        if (!(FindName(step.PartIndicator) is Rectangle totalIndicador) ||
                            !(FindName(step.PartPeso) is TextBlock pesoTotalTextBlock)) continue;

                        totalIndicador.Fill = Brushes.Green;
                        cantidadTotalTxbBlock.Text = "0";
                        pesoTotalTextBlock.Text = "OK";

                        step.DetectedWeight = "OK";

                        if (step == pasosAMostrar.Last())
                        {
                            //step.PartNoParte = ModeloData.NoModelo;
                            step.DetectedWeight = currentWeight.ToString();
                            (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(ModeloData.NoModelo);
                            codigo = $"{integrer}.{fraction}";
                            LogCompleteStep(step, "PESO TOTAL OK", codigo);
                        }
                    }

                    lbx_Codes.Visibility = Visibility.Hidden;
                    DesactivarSalida(defaultSettings.Piston);
                    ProcesarResetModelo();
                    SetImagesBox();
                    lblCompletados.Content = registroBitacora.Completadas;
                    ShowIniciar();
                    Dispatcher.Invoke(ShowPruebaCorrecta);
                    return;
                }
            }

            if (!(FindName($"Part_Cantidad{int.Parse(currentStep.PartOrden) - 1}") is TextBlock cantidadTextBlock))
                return;

            ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight,
                _validacion, int.Parse(cantidadTextBlock.Text));

            if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) ||
                !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

            if (pieceWeight >= currentStep.MinWeight && pieceWeight <= currentStep.MaxWeight)
            {
                if ((_currentStepIndex + 1 >= pasosFiltrados.Count && !_verificacionCompletado) ||
                    (_currentStepIndex + 1 >= pasosFiltrados.Count && _esperandoPickToLight))
                    return;

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

                lbx_Codes.Visibility = Visibility.Hidden;
                DesactivarSalida(defaultSettings.Piston);
                ProcesarResetModelo();
                SetImagesBox();
                lblCompletados.Content = registroBitacora.Completadas;
                ShowIniciar();
                Dispatcher.Invoke(ShowPruebaCorrecta);
                return;
            }

            indicator.Fill = Brushes.Red;
            pesoTextBlock.Text = $"{pieceWeight:F5} kg";
            currentStep.DetectedWeight = currentWeight.ToString("F5");
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
                        ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                            pieceWeight, _validacion, int.Parse(currentStep.PartCantidad));
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
            var auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                var newConfig = new ConfiguracionWindow();
                newConfig.Show();
                newConfig.Focus();

                newConfig.Closed += ConfiguracionWindow_Closed;

            }
        }

        private void ConfiguracionWindow_Closed(object sender, EventArgs e)
        {
            defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);
        }


        private void BtnCatalogos_Click(object sender, RoutedEventArgs e)
        {
            var auth = new AuthenticationWindow();

            if (auth.ShowDialog() == true)
            {
                _catalogos = new Catalogos();
                _catalogos.CambiosGuardados += ActualizarMainWindow;
                _catalogos.Show();
            }
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            //AuthenticationWindow auth = new AuthenticationWindow();

            //if (auth.ShowDialog() == true)
            //{
            if (Cbox_Modelo.SelectedItem == null)
            {
                ShowAlertError("No hay un modelo seleccionado.");
                return;
            }

            var modeloSeleccionado = ModeloData.NoModelo.ToUpper();

            if (ModeloExiste(modeloSeleccionado))
                ProcesarResetModelo();
            else
                ShowAlertError($"Modelo '{modeloSeleccionado}' no reconocido.");
            //}
        }

        private void BtnRechazo_Click(object sender, RoutedEventArgs e)
        {
            var auth = new AuthenticationWindow();

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

        private void Part_Imagen0_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            paso0 = pesoBascula;
        }
    }
}