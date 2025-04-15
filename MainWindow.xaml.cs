using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ClosedXML.Excel;
using log4net.Repository.Hierarchy;
using Scale_Program.Functions;
using Brushes = System.Windows.Media.Brushes;
using Rectangle = System.Windows.Shapes.Rectangle;

namespace Scale_Program
{
    public partial class MainWindow : Window
    {
        private readonly BasculaFunc bascula1;
        private readonly BasculaFunc bascula2;
        private readonly List<SequenceStep> valoresBolsas = new List<SequenceStep>();
        private Modelo ModeloData;
        private double _accumulatedWeight;
        private bool _activarSelladora;
        private AlertWindow _alertWindow;
        private Catalogos _catalogos = new Catalogos();
        private int _completedSequencesFerre;
        private int _consecutiveCount;
        private int _currentStepIndex;
        private ErrorMessageWindow _errorWindow;
        private bool _etapa1;
        private bool _etapa2;
        private bool _isInitializing;
        private double _lastWeight;
        private bool _startMeasurinBoxAndBags;
        private bool _stopBascula1 = true;
        private bool _stopBascula2 = true;
        private bool _validacion = false;
        private string codigo;
        private int contador;
        private Configuracion defaultSettings;
        private string fraction;
        private string integrer;
        public IOInterface ioInterface;
        public IOScanner ioScanner;
        private bool ioScannerActivado;
        private bool scannerReading;
        private List<SequenceStep> pasosFiltrados;
        private double pieceWeight;
        private int stepIndex = 0;
        private double runningWeight = 0.0;
        private bool sensorMasterActivo;
        private bool sensorProp65Activo;
        private bool sensorSelladoraActivo;
        private bool sensorUnitariaActivo;
        private List<SequenceStep> sequence;
        private string zpl;
        private DispatcherTimer _estadoBasculasTimer;

        private readonly string rutaImagenes = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes");

        public MainWindow()
        {
            InitializeComponent();

            defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);

            bascula1 = new BasculaFunc();
            bascula1.AsignarControles(Dispatcher);
            bascula1.OnDataReady += Bascula1_OnDataReady;

            bascula2 = new BasculaFunc();
            bascula2.AsignarControles(Dispatcher);
            bascula2.OnDataReady += Bascula2_OnDataReady;
        }

        private void Cbox_Modelo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                defaultSettings = Configuracion.Cargar(Configuracion.RutaArchivoConf);
                //IniciarSealevel();

                ModeloData = Cbox_Modelo.SelectedValue as Modelo ?? throw new Exception("Modelo no reconocido.");

                ProcesarModeloValido(ModeloData.NoModelo);

                bool conteoCajas = ModeloData.UsaConteoCajas;

                lblProgreso.Visibility = conteoCajas ? Visibility.Visible : Visibility.Hidden;
                lblConteoCajas.Visibility = conteoCajas ? Visibility.Visible : Visibility.Hidden;
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

            _startMeasurinBoxAndBags = false;
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

                using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Articulos");

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        try
                        {
                            var step = CrearSequenceStepDesdeFila(row, modeloModProceso);
                            if (step != null)
                            {
                                pasosFiltrados.Add(step);
                            }
                        }
                        catch (Exception ex)
                        {
                            ShowAlertError($"Error al procesar una fila para '{modeloSeleccionado}': {ex.Message}");
                        }
                    }
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

                using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Modelos");

                    var row = worksheet.RowsUsed()
                        .FirstOrDefault(r => r.Cell(1).GetString().ToUpper() == modeloSeleccionado);

                    if (row != null)
                    {
                        return row.Cell(2).GetValue<int>();
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

        private SequenceStep CrearSequenceStepDesdeFila(IXLRow row, int modeloModProceso)
        {
            var noParte = row.Cell(1).GetValue<string>();
            var modProceso = row.Cell(2).GetValue<int>();
            var proceso = row.Cell(3).GetValue<int>();
            var paso = row.Cell(4).GetValue<int>();
            var descripcion = row.Cell(5).GetValue<string>() ?? "";
            var pesoMin = row.Cell(6).GetValue<double>();
            var pesoMax = row.Cell(7).GetValue<double>();
            var cantidad = row.Cell(8).GetValue<int>();

            if (modProceso == modeloModProceso)
                return new SequenceStep
                {
                    MinWeight = pesoMin,
                    MaxWeight = pesoMax,
                    IsCompleted = false,
                    DetectedWeight = "",
                    Tag = "",
                    PartOrden = paso.ToString(),
                    GrdPart = $"Part{paso - 1}",
                    PartNoParte = noParte,
                    PartImagen = $"Part_Imagen{paso - 1}",
                    PartIndicator = $"Part_Indicator{paso - 1}",
                    PartPeso = $"Part_Peso{paso - 1}",
                    PartSecuencia = $"Part_Secuencia{paso - 1}",
                    PartProceso = proceso,
                    PartCantidad = cantidad.ToString(),
                    ModProceso = modProceso,
                    Descripcion = descripcion
                };

            return null;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            //OutputsOff();

            if (bascula1.GetPuerto() != null && bascula1.GetPuerto().IsOpen)
                bascula1.ClosePort();

            if (bascula2.GetPuerto() != null && bascula2.GetPuerto().IsOpen)
                bascula2.ClosePort();

            CerrarAlertas();

            Environment.Exit(0);
        }

        private void ResetSequence(List<SequenceStep> steps)
        {
            if (Cbox_Modelo.SelectedValue == null)
                throw new ArgumentException("No hay un modelo seleccionado.");

            var modeloSeleccionado = Cbox_Modelo.SelectedValue.ToString();

            if (ModeloExisteEnExcel(modeloSeleccionado))
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
            CargarModelosDesdeExcel();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ShowAlertError("RECUERDA CONFIGURAR LAS ENTRADAS Y EL CATALOGO ANTES DE EMPEZAR");
            IniciarMonitoreoEstadoBasculas();
            CargarModelosDesdeExcel();
        }

        private void CargarModelosDesdeExcel()
        {
            try
            {
                List<Modelo> modelos = new List<Modelo>();

                using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Modelos");

                    modelos = worksheet.RowsUsed()
                        .Skip(1)
                        .Select(row => new Modelo
                        {
                            NoModelo = row.Cell(1).GetString(),
                            ModProceso = row.Cell(2).GetValue<int>(),
                            Descripcion = row.Cell(3).GetString(),
                            UsaBascula1 = row.Cell(4).GetBoolean(),
                            UsaBascula2 = row.Cell(5).GetBoolean(),
                            UsaConteoCajas = row.Cell(6).GetBoolean(),
                            CantidadCajas = row.Cell(7).GetValue<int>(),
                            Etapa1 = row.Cell(8).GetString(),
                            Etapa2 = row.Cell(9).GetString(),
                            Activo = row.Cell(10).GetBoolean()
                        })
                        .Where(m => m.Activo)
                        .ToList();
                }

                Cbox_Modelo.ItemsSource = modelos;
                Cbox_Modelo.DisplayMemberPath = "NoModelo";
                Cbox_Modelo.SelectedValuePath = ".";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar modelos desde Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ModeloExisteEnExcel(string modelo)
        {
            try
            {
                using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Modelos");

                    return worksheet.RowsUsed()
                        .Skip(1)
                        .Any(row => row.Cell(1).GetString().ToUpper() == modelo.ToUpper());
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
                _stopBascula2 = true;

                _startMeasurinBoxAndBags = false;
                _consecutiveCount = 0;
                _accumulatedWeight = 0;
                _currentStepIndex = 0;
                lblProgreso.Content = 0;
                contador = 0;
                stepIndex = 0;
                runningWeight = 0.0;
                scannerReading = false;
                _validacion = false;
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

                var modeloSeleccionado = Cbox_Modelo.SelectedValue.ToString();

                if (!ModeloExisteEnExcel(modeloSeleccionado))
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
                ProcesarModeloValido(Cbox_Modelo.SelectedValue.ToString());
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

        private async void TxbScanner_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;

                var scannedCode = txbScanner.Text.Trim();
                var lastBag = valoresBolsas.FindLast(b => !b.IsCompleted);

                if (string.IsNullOrEmpty(scannedCode))
                {
                    await ShowMensaje("El código escaneado no puede estar vacío. Intente nuevamente.", Brushes.Red,
                        1500);
                    txbScanner.Clear();
                    return;
                }

                if (valoresBolsas.Count == 0)
                {
                    await ShowMensaje("No hay bolsas registradas para escanear.", Brushes.Red, 1500);
                    txbScanner.Clear();
                    return;
                }

                if (lastBag == null)
                {
                    await ShowMensaje("Todas las bolsas ya fueron escaneadas.", Brushes.Red, 1500);
                    txbScanner.Clear();
                    return;
                }

                if (scannedCode == lastBag.Tag)
                {
                    txbScanner.IsEnabled = false;
                    txbScanner.Visibility = Visibility.Hidden;
                    txbScanner.Clear();

                    if (Cbox_Modelo.Text == "ANTENNA KIT")
                        _stopBascula1 = false;

                    _activarSelladora = true;

                    if (Cbox_Modelo.Text == "ANTENNA KIT")
                        _currentStepIndex = pasosFiltrados.Count - 1;


                    lastBag.IsCompleted = true;
                    scannerReading = false;
                    txbScanner.KeyDown -= TxbScanner_KeyDown;

                    await ShowMensaje("SECUENCIA CORRECTA, PUEDE CERRAR LA BOLSA CON LA SELLADORA", Brushes.Green,
                        2500);
                }
                else
                {
                    txbScanner.Clear();
                    await ShowMensaje("CODIGO INCORRECTO", Brushes.Red, 1500);
                }
            }
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

        private void ReadScannerBolsa()
        {
            _stopBascula1 = true;
            scannerReading = true;

            lblPesoArt.Text = "ESCANEAR CÓDIGO DE LA BOLSA";
            txbArticulo.Text = "CÓDIGO";

            txbPesoMax.Visibility = Visibility.Hidden;
            txbPesoMin.Visibility = Visibility.Hidden;
            txbPesoActual.Visibility = Visibility.Hidden;

            lblPesoMax.Visibility = Visibility.Hidden;
            lblPesoMin.Visibility = Visibility.Hidden;
            lblPesoActual.Visibility = Visibility.Hidden;
            btnRechazo.Visibility = Visibility.Hidden;

            txbScanner.Visibility = Visibility.Visible;
            txbScanner.IsEnabled = true;
            txbScanner.Clear();
            txbScanner.Focus();
            txbScanner.KeyDown += TxbScanner_KeyDown;
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

        private void ShowAlertWindow()
        {
            _alertWindow = new AlertWindow();
            _alertWindow.ShowCompleteAndClose();
            _alertWindow.Show();
        }

        private void ShowScaner()
        {
            _alertWindow = new AlertWindow();
            _alertWindow.ShowScanner();
            _alertWindow.Show();
        }

        private void ShowAlertSensor()
        {
            if (_alertWindow != null && _alertWindow.IsVisible)
            {
                if (_alertWindow.lblPesoActual.IsVisible) _alertWindow.Close();
                return;
            }

            _alertWindow = new AlertWindow();
            _alertWindow.ShowNeedSensor();
            _alertWindow.Show();
        }

        private void ShowFerreSensor()
        {
            if (_alertWindow != null && _alertWindow.IsVisible)
            {
                if (_alertWindow.lblPesoActual.IsVisible) _alertWindow.Close();
                return;
            }

            _alertWindow = new AlertWindow();
            _alertWindow.ShowNeedSensorFerre();
            _alertWindow.Show();
        }

        private void ShowAlertaPeso(string name, double pesoMin, double pesoMax, double pesoActual, bool validacion)
        {
            if (!_validacion)
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


                txbArticulo.Text = name.ToUpper();
                txbPesoMax.Text = $"{pesoMax:F5} kg";
                txbPesoMin.Text = $"{pesoMin:F5} kg";
                txbPesoActual.Text = $"{pesoActual:F5} kg";

                grdValidacion.Visibility = Visibility.Visible;
                recValidacion.Visibility = Visibility.Visible;
            }
        }

        private void ShowAlertContinuidad()
        {
            if (!_validacion)
            {
                lblPesoArt.Text = "PRUEBA DE CONTINUIDAD";
                txbPesoMax.Visibility = Visibility.Hidden;
                txbPesoMin.Visibility = Visibility.Hidden;
                txbPesoActual.Visibility = Visibility.Visible;
                txbPesoActual.Text = "FAVOR DE REALIZAR LA PRUEBA";
                lblPesoMin.Visibility = Visibility.Hidden;
                lblPesoMax.Visibility = Visibility.Hidden;
                lblPesoActual.Visibility = Visibility.Hidden;
                txbScanner.Visibility = Visibility.Hidden;
                btnRechazo.Visibility = Visibility.Hidden;

                txbArticulo.Text = "";

                grdValidacion.Visibility = Visibility.Visible;
                recValidacion.Visibility = Visibility.Visible;
            }
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
            borderBascula2.Background = _stopBascula2 ? Brushes.Red : Brushes.Green;
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
                case "unitaria":
                    return defaultSettings.SalidaDispensadoraUnitaria;
                case "master":
                    return defaultSettings.SalidaDispensadoraMaster;
                case "prop65":
                    return defaultSettings.SalidaDispensadoraProp65;
                case "selladora":
                    return defaultSettings.SalidaSelladora;
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
                var sensorUnitaria = (inputsState & (1 << defaultSettings.EntradaSensorUnitaria)) != 0;
                var sensorProp65 = (inputsState & (1 << defaultSettings.EntradaSensorProp65)) != 0;
                var sensorMaster = (inputsState & (1 << defaultSettings.EntradaSensorMaster)) != 0;
                var sensorSelladora = (inputsState & (1 << defaultSettings.EntradaSensorSelladora)) != 0;

                if (sensorUnitaria && !sensorUnitariaActivo)
                    DesactivarSalida("unitaria");

                if (sensorMaster && !sensorMasterActivo)
                    DesactivarSalida("master");


                if (sensorProp65 && !sensorProp65Activo)
                    DesactivarSalida("prop65");

                if (sensorSelladora && sensorSelladoraActivo && _activarSelladora)
                {
                    ActivarSalida("selladora");

                    await Task.Delay(700);

                    DesactivarSalida("selladora");
                    _activarSelladora = false;
                }


                sensorUnitariaActivo = sensorUnitaria;
                sensorMasterActivo = sensorMaster;
                sensorProp65Activo = sensorProp65;
                sensorSelladoraActivo = sensorSelladora;
            });
        }

        #endregion

        #region LOGS

        private void LogFerreteriaStep(SequenceStep step, string resultado, string matchedTag)
        {
            try
            {
                using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Completados");
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cell(newRow, 2).Value = step.PartNoParte;
                    worksheet.Cell(newRow, 3).Value = step.ModProceso;
                    worksheet.Cell(newRow, 4).Value = step.PartProceso;
                    worksheet.Cell(newRow, 5).Value = step.DetectedWeight;
                    worksheet.Cell(newRow, 6).Value = resultado;
                    worksheet.Cell(newRow, 7).Value = matchedTag ?? "";

                    workbook.Save();
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
                using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Completados");
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cell(newRow, 2).Value = step.PartNoParte;
                    worksheet.Cell(newRow, 3).Value = step.ModProceso;
                    worksheet.Cell(newRow, 4).Value = step.PartProceso;
                    worksheet.Cell(newRow, 5).Value = step.DetectedWeight;
                    worksheet.Cell(newRow, 6).Value = "Rechazado";
                    worksheet.Cell(newRow, 7).Value = step.Tag ?? "";

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al registrar la pieza como Rechazada: {ex.Message}");
            }
        }

        #endregion

        #region Basculas

        private async void ReadInputBascula1()
        {
            await InicializarPuertoDeBasculaAsync(defaultSettings.PuertoBascula1, bascula1);
        }

        private async void ReadInputBascula2()
        {
            await InicializarPuertoDeBasculaAsync(defaultSettings.PuertoBascula2, bascula2);
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
                        _stopBascula2 = true;
                    }
                }
            });
        }

        private void Bascula1_OnDataReady(object sender, BasculaEventArgs e)
        {
            if (_stopBascula1 || scannerReading) return;

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

                    if (_consecutiveCount == 4)
                    {

                        if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && ModeloData.UsaConteoCajas || ModeloData.UsaBascula1 && !ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas)
                            ProcessSequenceFerreteria(weight);

                        else if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas)
                            ProcessStableWeightArmHarness(weight);
                    }
                }

            });
        }

        private void Bascula2_OnDataReady(object sender, BasculaEventArgs e)
        {
            if (_stopBascula2 || scannerReading) return;

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

                    if (_consecutiveCount == 2)
                    {
                        double tolerancia = 0.015;

                        if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && ModeloData.UsaConteoCajas && _etapa2)
                            ProcessCajaFerreteria(weight, tolerancia);

                        else if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas && _etapa2)
                            ProcessCajaEtapa2(weight, ref stepIndex, ref runningWeight, pasosFiltrados, valoresBolsas, tolerancia);

                        else if (!ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas && _etapa1)
                            ProcessSequenceFerreteria(weight);
                    }
                }
               
            });
        }

        private void SecuenciaASeguir(Modelo modeloseleccionado)
        {
            try
            {
                if (modeloseleccionado.UsaBascula1 && !modeloseleccionado.UsaBascula2 && !modeloseleccionado.UsaConteoCajas)
                {
                    //Secuencia de básculas individual - solo báscula 1

                    if (bascula1?.GetPuerto() == null || !bascula1.GetPuerto().IsOpen)
                        ReadInputBascula1();

                    _stopBascula1 = false;
                    _stopBascula2 = true;

                }
                else if (!modeloseleccionado.UsaBascula1 && modeloseleccionado.UsaBascula2 && !modeloseleccionado.UsaConteoCajas)
                {
                    //Secuencia de báscula individual - solo báscula 2
                    if (bascula2?.GetPuerto() == null || !bascula2.GetPuerto().IsOpen)
                        ReadInputBascula2();

                    _stopBascula1 = true;
                    _stopBascula2 = false;
                }
                else if (modeloseleccionado.UsaBascula1 && modeloseleccionado.UsaBascula2)
                {
                    //Secuencia de ambas básculas
                    if (bascula1?.GetPuerto() == null || !bascula1.GetPuerto().IsOpen)
                        ReadInputBascula1();

                    _stopBascula1 = false;
                    _stopBascula2 = true;
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
                            LogFerreteriaStep(currentStep, "OK", codigo);

                        if (_currentStepIndex == pasosFiltrados.Count - 2)
                        {
                            (zpl, integrer, fraction) =
                                ZebraPrinter.GenerateZplBody(Cbox_Modelo.SelectedValue.ToString());
                            codigo = $"{integrer}.{fraction}";
                            //Console.WriteLine(codigo);
                            //RawPrinterHelper.SendStringToPrinter(defaultSettings.ZebraName, zpl);

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
                                PartCantidad = currentStep.PartCantidad,
                                ModProceso = currentStep.ModProceso,
                                PartProceso = currentStep.PartProceso
                            });

                            //ReadScannerBolsa();
                        }

                        _currentStepIndex++;

                        if (_currentStepIndex < pasosFiltrados.Count)
                        {
                            currentStep = pasosFiltrados[_currentStepIndex];
                            pieceWeight = currentWeight - _accumulatedWeight;

                            if (ModeloData.NoModelo == "ANTENNA KIT" && pasosFiltrados[_currentStepIndex].PartNoParte == "Caja antenna 141A1528")
                            {
                                _ = ShowMensaje("PUEDES CERRAR LA BOLSA DE FERRETERIA CON LA SELLADORA", Brushes.Green,
                                    6000);
                                _activarSelladora = true;
                            }

                            ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                                pieceWeight,
                                _validacion, int.Parse(currentStep.PartCantidad));
                            return;
                        }

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

                        if (ModeloData.UsaConteoCajas)
                            lblProgreso.Content = contador;

                        else if (!ModeloData.UsaConteoCajas)
                            lblCompletados.Content = contador;


                        if (btnEtiquetaManual.Tag?.ToString() == "off")
                        {
                            if (Cbox_Modelo.Text == "123-0253-000")
                            {
                                ActivarSalida("prop65");

                                (zpl, integrer, fraction) =
                                    ZebraPrinter.GenerateZplBody(Cbox_Modelo.SelectedValue.ToString());
                                codigo = $"{integrer}.{fraction}";
                                RawPrinterHelper.SendStringToPrinter(defaultSettings.ZebraName, zpl);
                                LogFerreteriaStep(currentStep, "OK", codigo);
                            }
                            else
                            {
                                ActivarSalida("unitaria");
                                ActivarSalida("prop65");
                            }
                        }

                        codigo = "";
                        Dispatcher.Invoke(ShowPruebaCorrecta);
                    }
                }
            }
            else
            {
                indicator.Fill = Brushes.Red;
                pesoTextBlock.Text = $"{pieceWeight:F5} kg";
                currentStep.DetectedWeight = currentWeight.ToString("F5");
            }
        }

        private void ProcessStableWeightArmHarness(double currentWeight)
        {
            const int pasosPorPagina = 8;
            var totalPasos = pasosFiltrados.Count;
            var paginaActual = _currentStepIndex / pasosPorPagina;

            var pasosAMostrar = pasosFiltrados
                .Skip(paginaActual * pasosPorPagina)
                .Take(pasosPorPagina)
                .ToList();

            var indexEnPagina = _currentStepIndex % pasosPorPagina;

            if (_currentStepIndex >= totalPasos) return;

            var currentStep = pasosAMostrar[indexEnPagina];
            pieceWeight = currentWeight - _accumulatedWeight;

            var cantidadTextBlockName = currentStep.PartIndicator.Replace("Part_Indicator", "Part_Cantidad");

            if (!(FindName(cantidadTextBlockName) is TextBlock cantidadTextBlock)) return;

            ShowBolsasRestantes(currentStep.PartNoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight, _validacion, int.Parse(cantidadTextBlock.Text));

            if (!(FindName(currentStep.PartIndicator) is Rectangle indicator) || !(FindName(currentStep.PartPeso) is TextBlock pesoTextBlock)) return;

            var isLastStep = _currentStepIndex == pasosFiltrados.Count - 1;

            if (isLastStep)
            {
                if (pieceWeight >= currentStep.MinWeight && pieceWeight <= currentStep.MaxWeight)
                {
                    CompleteCurrentStep(currentStep, indicator, pesoTextBlock, cantidadTextBlock, currentWeight);

                    _currentStepIndex = 0;
                    _accumulatedWeight = 0;
                    pieceWeight = 0;
                    _stopBascula1 = true;
                    _etapa2 = true;
                    CargarSecuenciasPorModelo(Cbox_Modelo.SelectionBoxItem.ToString());
                    CargarProcesos();
                    SetValuesEtapas(sequence, 2);
                    ReindexarPasos(pasosFiltrados);
                    SetImagesBox();

                    if (bascula2.GetPuerto() == null || !bascula2.GetPuerto().IsOpen)
                        ReadInputBascula2();

                    _activarSelladora = true;

                    _ = ShowMensaje("FERRETERIA COMPLETA, PUEDE CERRAR LA BOLSA CON LA SELLADORA", Brushes.Green, 2500);

                    _stopBascula2 = false;
                    return;
                }
            }
            else if (pieceWeight >= currentStep.MinWeight && pieceWeight <= currentStep.MaxWeight)
            {
                CompleteCurrentStep(currentStep, indicator, pesoTextBlock, cantidadTextBlock, currentWeight);
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
                    if (!currentStep.IsCompleted)
                        LogFerreteriaStep(currentStep, "OK", null);

                    if (_currentStepIndex == pasosFiltrados.Count - 1)
                    {
                        (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(Cbox_Modelo.SelectedValue.ToString());
                        codigo = $"{integrer}.{fraction}";
                        //RawPrinterHelper.SendStringToPrinter(defaultSettings.ZebraName, zpl);
                        //Console.WriteLine(codigo);

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

                        //ReadScannerBolsa();
                        return;
                    }

                    _currentStepIndex++;

                    if (_currentStepIndex < pasosFiltrados.Count)
                    {
                        var pasosPorPagina = 8;

                        if (_currentStepIndex % pasosPorPagina == 0)
                        {
                            var paginaActual = _currentStepIndex / pasosPorPagina;

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
                    }
                }
            }
        }

        private void ProcessCajaEtapa2(double weight, ref int stepIndex, ref double runningWeight,
            List<SequenceStep> pasosFiltrados, List<SequenceStep> valoresBolsas, double tolerancia)
        {
            if (stepIndex >= pasosFiltrados.Count) return;

            var currentStep = pasosFiltrados[stepIndex];

            double minRange, maxRange;
            var pieceWeight = weight - runningWeight;

            if (currentStep.PartNoParte.Contains("Ferreteria"))
            {
                var sumBolsas = valoresBolsas.Sum(b => double.Parse(b.DetectedWeight));
                currentStep.DetectedWeight = sumBolsas.ToString("F5");
                currentStep.MinWeight = sumBolsas;
                currentStep.MaxWeight = sumBolsas;
                minRange = currentStep.MinWeight - (currentStep.MinWeight * tolerancia);
                maxRange = (currentStep.MaxWeight * tolerancia) + currentStep.MaxWeight;
            }
            else
            {
                minRange = currentStep.MinWeight - (currentStep.MinWeight * tolerancia);
                maxRange = (currentStep.MaxWeight * tolerancia) + currentStep.MaxWeight;

            }

            ShowBolsasRestantes(currentStep.PartNoParte, minRange, maxRange, pieceWeight, _validacion,int.Parse(currentStep.PartCantidad));

            if (pieceWeight >= minRange && pieceWeight <= maxRange)
            {
                var indicator = FindName(currentStep.PartIndicator) as Rectangle;
                var pesoTextBlock = FindName(currentStep.PartPeso) as TextBlock;
                var cantidadTextBlockName = currentStep.PartIndicator.Replace("Part_Indicator", "Part_Cantidad");
                var cantidadTextBlock = FindName(cantidadTextBlockName) as TextBlock;

                CompleteCurrentStepEtapa2(currentStep, indicator, pesoTextBlock, cantidadTextBlock, pieceWeight);

                runningWeight = weight;
                stepIndex++;

                if (stepIndex < pasosFiltrados.Count)
                {
                    currentStep = pasosFiltrados[stepIndex];
                    pieceWeight = weight - runningWeight;

                    if (currentStep.PartNoParte.Contains("Ferreteria"))
                    {                
                        var sumBolsas = valoresBolsas.Sum(b => double.Parse(b.DetectedWeight));
                        currentStep.DetectedWeight = sumBolsas.ToString("F5");
                        currentStep.MinWeight = sumBolsas;
                        currentStep.MaxWeight = sumBolsas;
                        minRange = currentStep.MinWeight - (currentStep.MinWeight * tolerancia);
                        maxRange = (currentStep.MaxWeight * tolerancia) + currentStep.MaxWeight;
                    }
                    else
                    {
                        minRange = currentStep.MinWeight - (currentStep.MinWeight * tolerancia);
                        maxRange = (currentStep.MaxWeight * tolerancia) + currentStep.MaxWeight;
                    }

                    ShowBolsasRestantes(currentStep.PartNoParte, minRange, maxRange, pieceWeight, _validacion,int.Parse(currentStep.PartCantidad));
                }

                if (stepIndex >= pasosFiltrados.Count)
                {
                    LogFerreteriaStep(currentStep, "OK", codigo);
                    EndFerreteriaG_ArmHarness();
                }
            }
            else
            {
                if (FindName(currentStep.PartIndicator) is Rectangle indicator) indicator.Fill = Brushes.Red;
                if (FindName(currentStep.PartPeso) is TextBlock pesoTextBlock) pesoTextBlock.Text = $"{pieceWeight:F5}  kg";
                if (FindName(currentStep.PartSecuencia) is TextBlock secuencia) secuencia.Text = $"Rango: {minRange:F5}/{maxRange:F5}";
                currentStep.DetectedWeight = pieceWeight.ToString("F5");
            }
        }

        private void CompleteCurrentStepEtapa2(SequenceStep currentStep, Rectangle indicator, TextBlock pesoTextBlock,
            TextBlock cantidadTextBlock, double currentWeight)
        {
            currentStep.DetectedWeight = currentWeight.ToString("F5");
            indicator.Fill = Brushes.Green;
            pesoTextBlock.Text = $"{currentWeight:F5} kg";

            if (int.TryParse(cantidadTextBlock.Text, out var cantidadRestante) && cantidadRestante > 0)
            {
                cantidadRestante--;
                cantidadTextBlock.Text = cantidadRestante.ToString();
                _accumulatedWeight = currentWeight;

                if (!(stepIndex + 1 >= pasosFiltrados.Count))
                    LogFerreteriaStep(currentStep, "OK", null);

                if (cantidadRestante == 0)
                    //_currentStepIndex++;
                    if (_currentStepIndex < pasosFiltrados.Count)
                    {
                        //var nextStep = pasosFiltrados[_currentStepIndex];
                        pieceWeight = currentWeight - _accumulatedWeight;

                        /*ShowBolsasRestantes(nextStep.Part_NoParte, nextStep.MinWeight, nextStep.MaxWeight,
                            pieceWeight, _validacion, int.Parse(nextStep.Part_Cantidad));*/
                    }
            }
        }

        private async void EndFerreteriaG()
        {
            await Dispatcher.InvokeAsync(ShowPruebaCorrecta);

            await Task.Delay(2000);
            if (btnEtiquetaManual.Tag?.ToString() == "off")
                ActivarSalida("master");

            _stopBascula1 = false;

            _startMeasurinBoxAndBags = false;
            _accumulatedWeight = 0;
            _currentStepIndex = 0;
            lblProgreso.Content = 0;
            _completedSequencesFerre++;
            lblCompletados.Content = _completedSequencesFerre;
            _validacion = false;
            valoresBolsas.Clear();
            _etapa1 = true;
            _etapa2 = false;
            contador = 0;
            txbPesoActual.Text = "0.0Kg";
            PesoGeneral.Text = "0.0Kg";

            HideAll();
            ProcesarModeloValido(Cbox_Modelo.SelectedValue.ToString());
            pasosFiltrados = ObtenerValoresProceso(Cbox_Modelo.SelectedValue.ToString(),
                1);
            ProcesarModeloYProceso();
            Cbox_Proceso.SelectedIndex = 0;
            SetValuesEtapas(sequence, 1);
            SetImagesBox();
        }

        private async void EndFerreteriaG_ArmHarness()
        {
            await Dispatcher.InvokeAsync(ShowPruebaCorrecta);

            await Task.Delay(2000);

            if (btnEtiquetaManual.Tag?.ToString() == "off")
            {
                ActivarSalida("unitaria");
                ActivarSalida("prop65");
            }

            _stopBascula1 = false;
            _stopBascula2 = true;

            _startMeasurinBoxAndBags = false;
            _accumulatedWeight = 0;
            _currentStepIndex = 0;
            lblProgreso.Content = 0;
            _completedSequencesFerre++;
            lblCompletados.Content = _completedSequencesFerre;
            _validacion = false;
            valoresBolsas.Clear();
            _etapa1 = true;
            _etapa2 = false;
            contador = 0;
            stepIndex = 0;
            runningWeight = 0.0;
            codigo = "";
            txbPesoActual.Text = "0.0Kg";
            PesoGeneral.Text = "0.0Kg";

            HideAll();
            ProcesarModeloValido(Cbox_Modelo.SelectedValue.ToString());
            sequence = ObtenerValoresProceso(Cbox_Modelo.SelectedValue.ToString(), 1);
            ProcesarModeloYProceso();
            Cbox_Proceso.SelectedIndex = 0;
            SetValuesEtapas(sequence, 1);
            SetImagesBox();
        }

        private void ProcessCajaFerreteria(double weight, double tolerancia)
        {
            if (pasosFiltrados.Count != 2)
            {
                _ = ShowMensaje("ESTA SECUENCIA SOLO LLEVA 2 ARTICULOS EN ETAPA 2, FAVOR DE ARREGLAR LA SECUENCIA EN CATALOGOS.",Brushes.Red,15000);
                _stopBascula1 = true;
                _stopBascula2 = true;
                return;
            }

            TextBlock cajaPeso = null;
            Rectangle cajaIndicador = null;

            var cajaVaciaStep = pasosFiltrados.FirstOrDefault();
            var cajaCompletaStep = pasosFiltrados.Last();

            var pesoTotalBolsas = valoresBolsas.Sum(s => double.Parse(s.DetectedWeight));

            var pieceWeight = weight - _accumulatedWeight;

            if (!_startMeasurinBoxAndBags)
            {
                ShowAlertaPeso("PESAR SOLO CAJA MASTER", cajaVaciaStep.MinWeight, cajaVaciaStep.MaxWeight, pieceWeight,
                    _validacion);
                var secuencia = FindName(cajaVaciaStep.PartSecuencia) as TextBlock;
                cajaIndicador = FindName(cajaVaciaStep.PartIndicator) as Rectangle;
                cajaPeso = FindName(cajaVaciaStep.PartPeso) as TextBlock;

                secuencia.Text = $"Rango: {cajaVaciaStep.MinWeight:F5} / {cajaVaciaStep.MaxWeight:F5}";

                if (pieceWeight >= cajaVaciaStep.MinWeight && pieceWeight <= cajaVaciaStep.MaxWeight)
                {
                    cajaVaciaStep.DetectedWeight = pieceWeight.ToString("F5");
                    LogFerreteriaStep(cajaVaciaStep, "CAJA VACÍA OK", null);

                    _startMeasurinBoxAndBags = true;

                    if (cajaPeso == null || cajaIndicador == null) return;

                    cajaIndicador.Fill = Brushes.Green;
                    cajaPeso.Text = $"{cajaVaciaStep.DetectedWeight:F5} kg";
                    _accumulatedWeight += pieceWeight;
                }
            }
            else
            {
                var pesoMinimo = pesoTotalBolsas - pesoTotalBolsas * tolerancia;
                var pesoMaximo = pesoTotalBolsas * tolerancia + pesoTotalBolsas;

                ShowAlertaPeso("PESAR CAJA MASTER CON FERRETERIAS", pesoMinimo, pesoMaximo, pieceWeight,
                    _validacion);
                var secuencia2 = FindName(cajaCompletaStep.PartSecuencia) as TextBlock;
                secuencia2.Text = $"Rango: {pesoMinimo:F5} / {pesoMaximo:F5}";
                if (pieceWeight >= pesoMinimo && pieceWeight <= pesoMaximo)
                {
                    cajaIndicador = FindName(cajaCompletaStep.PartIndicator) as Rectangle;
                    cajaPeso = FindName(cajaCompletaStep.PartPeso) as TextBlock;
                    var allTags = string.Join(", ",
                        valoresBolsas.Select(b => b.Tag).Where(tag => !string.IsNullOrEmpty(tag)));

                    _stopBascula2 = true;
                    if (cajaPeso != null && cajaIndicador != null)
                    {
                        cajaIndicador.Fill = Brushes.Green;
                        _accumulatedWeight += pieceWeight;
                        cajaCompletaStep.DetectedWeight = _accumulatedWeight.ToString("F5");
                        cajaPeso.Text = $"{cajaCompletaStep.DetectedWeight:F5} kg";
                    }

                    LogFerreteriaStep(cajaCompletaStep, "CAJA COMPLETA OK", allTags);
                    EndFerreteriaG();
                }
            }
        }

        private void ProcessSequenceFerreteria(double currentWeight)
        {
            _consecutiveCount = 0;

            if (contador == ModeloData.CantidadCajas && ModeloData.UsaConteoCajas)
            {
                _stopBascula1 = true;
                _etapa2 = true;
                CargarSecuenciasPorModelo(Cbox_Modelo.SelectionBoxItem.ToString());
                CargarProcesos();
                SetValuesEtapas(sequence, 2);
                ReindexarPasos(pasosFiltrados);
                SetImagesBox();

                if (bascula2.GetPuerto() == null || !bascula2.GetPuerto().IsOpen)
                    ReadInputBascula2();

                _stopBascula2 = false;
            }

            if (_etapa1)
                ProcessStableWeight(currentWeight);
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

                var modeloSeleccionado = Cbox_Modelo.Text.ToUpper();

                if (ModeloExisteEnExcel(modeloSeleccionado))
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

                var modeloSeleccionado = Cbox_Modelo.SelectedValue.ToString();

                if (!ModeloExisteEnExcel(modeloSeleccionado))
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

        private void BtnSelladora_Click(object sender, RoutedEventArgs e)
        {
            _activarSelladora = true;
        }

        private async void BtnEtiqueta_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(zpl))
            {
                RawPrinterHelper.SendStringToPrinter(defaultSettings.ZebraName, zpl);
                await ShowMensaje("Imprimiendo etiqueta", Brushes.LightGreen, 2000);
            }
            else
            {
                await ShowMensaje("No se ha generado ninguna etiqueta para imprimir.", Brushes.BlanchedAlmond, 2000);
            }
        }

        private void BtnEtiquetaManual_Click(object sender, RoutedEventArgs e)
        {
            if (btnEtiquetaManual.Tag?.ToString() == "off")
            {
                var imagePath = Path.Combine(rutaImagenes, "check.png");
        
                if (File.Exists(imagePath))
                    btnEtiquetaManual.Background = new ImageBrush(new BitmapImage(new Uri(imagePath, UriKind.Absolute)));

                btnEtiquetaManual.Tag = "on";
            }
            else
            {
                var imagePath = Path.Combine(rutaImagenes, "checkoff.png");
        
                if (File.Exists(imagePath))
                    btnEtiquetaManual.Background = new ImageBrush(new BitmapImage(new Uri(imagePath, UriKind.Absolute)));

                btnEtiquetaManual.Tag = "off";
            }
        }

        #endregion 
    }
}