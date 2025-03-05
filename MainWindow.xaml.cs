using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClosedXML.Excel;
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
        private bool _FerreteriaPart2;
        private bool _isInitializing;
        private double _lastWeight;
        private bool _pistonAbierto;
        private bool _sequenceFinished;
        private bool _startMeasurinBoxAndBags;
        private bool _stopBascula1;
        private bool _stopBascula2 = true;
        private bool _validacion;
        private List<SequenceStep> caja;
        private string codigo;
        private int contador;
        private int contadorTotal;
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

        private string rutaImagenes = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes");

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
            _FerreteriaPart2 = false;
            LedPart2.Background = Brushes.Red;
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

                ProcesarModeloYProceso(ModeloData.NoModelo);

                SecuenciaASeguir(ModeloData);
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al cargar la secuencia: {ex.Message}");
            }
        }

        private void SecuenciaASeguir(Modelo modeloseleccionado)
        {
            if (modeloseleccionado.UsaBascula1 && !modeloseleccionado.UsaBascula2 && !modeloseleccionado.UsaConteoCajas)
            {
                //Secuencia de bascula individual con solo bascula 1

                if (bascula1?.GetPuerto() == null || !bascula1.GetPuerto().IsOpen)
                    ReadInputBascula1();

                _stopBascula1 = false;
                _stopBascula2 = true;

            }
            else if (!modeloseleccionado.UsaBascula1 && modeloseleccionado.UsaBascula2 && !modeloseleccionado.UsaConteoCajas)
            {
                //Secuencia de bascula individual solo bascula 2
                if (bascula2?.GetPuerto() == null || !bascula2.GetPuerto().IsOpen)
                    ReadInputBascula2();

                _stopBascula1 = true;
                _stopBascula2 = false;
            }
            else if (modeloseleccionado.UsaBascula1 && modeloseleccionado.UsaBascula2 && !modeloseleccionado.UsaConteoCajas)
            {
                //Secuencia de bascula con ambas basculas sin conteo cajas
                if (bascula1?.GetPuerto() == null || !bascula1.GetPuerto().IsOpen)
                    ReadInputBascula1();

                _stopBascula1 = false;
                _stopBascula2 = true;
            }
            else if (modeloseleccionado.UsaBascula1 && modeloseleccionado.UsaBascula2 && modeloseleccionado.UsaConteoCajas)
            {
                //Secuencia de bascula con ambas bascula mas conteo de cajas
                if (bascula2?.GetPuerto() == null || !bascula2.GetPuerto().IsOpen)
                    ReadInputBascula1();

                _stopBascula1 = false;
                _stopBascula2 = true;
            }
            else
                ShowAlertError($"Error al cargar la secuencia, secuencia inexistente.");
        }

        private List<SequenceStep> ObtenerValoresProceso(string modeloSeleccionado, int procesoSeleccionado)
        {
            try
            {
                int modProceso = ObtenerModeloModProceso(modeloSeleccionado);
                if (modProceso == -1) return null;

                return sequence
                    .Where(s => s.Part_Proceso == procesoSeleccionado && s.ModProceso == modProceso)
                    .OrderBy(s => int.Parse(s.Part_Orden))
                    .ToList();
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al obtener la secuencia filtrada para '{modeloSeleccionado}': {ex.Message}");
                return null;
            }
        }

        private void ProcesarModeloYProceso(string modeloSeleccionado)
        {
            ResetSequence(pasosFiltrados);
            HideAll();
            SetValuesEtapas(sequence, 1);
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
                    Part_Orden = paso.ToString(),
                    grdPart = $"Part{paso - 1}",
                    Part_NoParte = noParte,
                    Part_Imagen = $"Part_Imagen{paso - 1}",
                    Part_Indicator = $"Part_Indicator{paso - 1}",
                    Part_Peso = $"Part_Peso{paso - 1}",
                    Part_Proceso = proceso,
                    Part_Cantidad = cantidad.ToString(),
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

        private void btnCerrarPeso_Click(object sender, RoutedEventArgs e)
        {
            _validacion = true;
            recValidacion.Visibility = Visibility.Hidden;
            grdValidacion.Visibility = Visibility.Hidden;
        }

        private void btnValidaciones_Click(object sender, RoutedEventArgs e)
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

                var indicator = FindName(step.Part_Indicator) as Rectangle;
                var pesoTextBlock = FindName(step.Part_Peso) as TextBlock;

                if (indicator != null)
                    indicator.Fill = Brushes.Red;

                if (pesoTextBlock != null)
                    pesoTextBlock.Text = "0.0Kgs";
            }

            SetValues(pasosFiltrados);
            _accumulatedWeight = 0.0;
            _currentStepIndex = 0;
        }

        private void Configuracion_btn_Click(object sender, RoutedEventArgs e)
        {
            var newConfig = new ConfiguracionWindow();
            newConfig.Show();
            newConfig.Focus();
        }

        private void btnCatalogos_Click(object sender, RoutedEventArgs e)
        {
            _catalogos = new Catalogos();
            _catalogos.CambiosGuardados += ActualizarMainWindow;
            _catalogos.Show();
        }

        private void ActualizarMainWindow()
        {
            CargarModelosDesdeExcel();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ShowAlertError("RECUERDA CONFIGURAR LAS ENTRADAS Y EL CATALOGO ANTES DE EMPEZAR");
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

                    // Leer todas las filas (omitiendo la primera fila de encabezados)
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
                        .Where(m => m.Activo) // Filtrar solo modelos activos
                        .ToList();
                }

                // Asignar los modelos al ComboBox (solo los activos)
                Cbox_Modelo.ItemsSource = modelos;
                Cbox_Modelo.DisplayMemberPath = "NoModelo";  // Mostrar solo el nombre del modelo
                Cbox_Modelo.SelectedValuePath = ".";  // Permitir seleccionar el objeto completo
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar modelos desde Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
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
                // Registrar el paso rechazado si corresponde
                if (_FerreteriaPart2)
                    RegistrarPasoRechazado(caja.FirstOrDefault());
                else if (_currentStepIndex < pasosFiltrados.Count)
                    RegistrarPasoRechazado(pasosFiltrados[_currentStepIndex]);

                // Reinicializar variables
                CerrarAlertas();
                _stopBascula1 = false;
                _stopBascula2 = true;
                _FerreteriaPart2 = false;
                _startMeasurinBoxAndBags = false;
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

                var procesoSeleccionado = int.Parse(Cbox_Proceso.SelectedValue.ToString());

                ObtenerValoresProceso(modeloSeleccionado, procesoSeleccionado);
                ProcesarModeloYProceso(modeloSeleccionado);
                SecuenciaASeguir(ModeloData);
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

        private void btnRechazo_Click(object sender, RoutedEventArgs e)
        {
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
                    ShowAlertError($"Se rechazó la pieza: {currentStep.Part_NoParte}, pieza Rechazada en FERRETERÍA");
                }
                else
                {
                    ShowAlertError("No hay una pieza válida para rechazar en FERRETERÍA.");
                }
            }
            catch (Exception ex)
            {
                ShowAlertError($"Error al registrar el rechazo en FERRETERÍA: {ex.Message}");
            }
        }


        private void btnSelladora_Click(object sender, RoutedEventArgs e)
        {
            _activarSelladora = true;
        }

        private async void btnEtiqueta_Click(object sender, RoutedEventArgs e)
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

        private void btnEtiquetaManual_Click(object sender, RoutedEventArgs e)
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

        #region BASCULAS

        // GUARDAR SE VA VERIFICAR EN UN FUTURO CON BASCULA 1 BASCULA 2 PARA DETECTAR A QUE SECUENCIA VA

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
                        //if (Cbox_Modelo.Text == "ANTENNA KIT" || Cbox_Modelo.Text == "123-0253-000") ProcessSequenceFerreteria(weight, isStable);

                        if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && ModeloData.UsaConteoCajas || ModeloData.UsaBascula1 && !ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas)
                        {
                            ProcessSequenceFerreteria(weight, isStable);
                        }

                        else if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas)
                        {
                            ProcessStableWeightArmHarness(weight, isStable);
                        }
                    }
                }

            });
        }

        private void ProcessStableWeightArmHarness(double currentWeight, bool isStable)
        {
            if (_FerreteriaPart2) return;

            var currentStep = pasosFiltrados[_currentStepIndex];
            pieceWeight = currentWeight - _accumulatedWeight;

            var indicator = FindName(currentStep.Part_Indicator) as Rectangle;
            var pesoTextBlock = FindName(currentStep.Part_Peso) as TextBlock;
            var cantidadTextBlockName = currentStep.Part_Indicator.Replace("Part_Indicator", "Part_Cantidad");
            var cantidadTextBlock = FindName(cantidadTextBlockName) as TextBlock;

            if (!isStable || cantidadTextBlock == null) return;

            ShowBolsasRestantes(currentStep.Part_NoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight, _validacion, int.Parse(cantidadTextBlock.Text));

            if (indicator == null || pesoTextBlock == null) return;

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
            pesoTextBlock.Text = "Fuera de rango";
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
                    if(!currentStep.IsCompleted)
                        LogFerreteriaStep(currentStep, "OK", null);

                    if (_currentStepIndex == pasosFiltrados.Count - 1)
                    {
                        (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(Cbox_Modelo.SelectedValue.ToString());
                        codigo = $"{integrer}.{fraction}";
                        //RawPrinterHelper.SendStringToPrinter(defaultSettings.ZebraName, zpl);
                        //Console.WriteLine(codigo);

                        valoresBolsas.Add(new SequenceStep
                        {
                            Part_NoParte = currentStep.Part_NoParte,
                            MinWeight = currentStep.MinWeight,
                            MaxWeight = currentStep.MaxWeight,
                            DetectedWeight = currentWeight.ToString(),
                            Tag = codigo,
                            IsCompleted = false,
                            Part_Indicator = currentStep.Part_Indicator,
                            Part_Peso = currentStep.Part_Peso,
                            Part_Orden = currentStep.Part_Orden,
                            Part_Cantidad = currentStep.Part_Cantidad
                        });

                        //ReadScannerBolsa();
                        return;
                    }

                    _currentStepIndex++;

                    if (_currentStepIndex < pasosFiltrados.Count)
                    {
                        currentStep = pasosFiltrados[_currentStepIndex];
                        pieceWeight = currentWeight - _accumulatedWeight;

                        ShowBolsasRestantes(currentStep.Part_NoParte, currentStep.MinWeight, currentStep.MaxWeight,
                            pieceWeight, _validacion, int.Parse(currentStep.Part_Cantidad));
                        return;
                    }
                }
            }
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
                        double tolerancia = 0.010;  // Ajusta según necesidad

                        if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && ModeloData.UsaConteoCajas && _etapa2)
                            ProcessCajaFerreteria(weight, isStable);

                        else if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas && _etapa2)
                            ProcessCajaEtapa2(weight, isStable, ref stepIndex, ref runningWeight, pasosFiltrados, valoresBolsas, tolerancia);

                        else if (!ModeloData.UsaBascula1 && ModeloData.UsaBascula2 && !ModeloData.UsaConteoCajas && _etapa1)
                        {
                            ProcessSequenceFerreteria(weight,isStable);
                        }
                    }
                }
               
            });
        }


        private void ProcessCajaEtapa2(double weight, bool isStable, ref int stepIndex, ref double runningWeight, 
            List<SequenceStep> pasosFiltrados, List<SequenceStep> valoresBolsas, double tolerancia)
        {
            if (!isStable) return;                     
            if (stepIndex >= pasosFiltrados.Count) return; 

            var currentStep = pasosFiltrados[stepIndex];

            double minRange, maxRange;
            double pieceWeight = weight - runningWeight;

            if (currentStep.Part_NoParte.Contains("Ferreteria"))
            {                
                double sumBolsas = valoresBolsas.Sum(b => double.Parse(b.DetectedWeight));
                currentStep.DetectedWeight = sumBolsas.ToString("F5");
                currentStep.MinWeight = sumBolsas;
                currentStep.MaxWeight = sumBolsas;
                minRange = runningWeight + currentStep.MinWeight;
                maxRange = runningWeight + currentStep.MaxWeight;
            }
            else
            {
                minRange = runningWeight + currentStep.MinWeight;
                maxRange = runningWeight + currentStep.MaxWeight;
            }

            minRange -= tolerancia;
            maxRange += tolerancia;

            ShowAlertaPeso(currentStep.Part_NoParte, minRange, maxRange, weight, _validacion);

            if (weight >= minRange && weight <= maxRange)
            {
                var indicator = FindName(currentStep.Part_Indicator) as Rectangle;
                var pesoTextBlock = FindName(currentStep.Part_Peso) as TextBlock;
                var cantidadTextBlockName = currentStep.Part_Indicator.Replace("Part_Indicator", "Part_Cantidad");
                var cantidadTextBlock = FindName(cantidadTextBlockName) as TextBlock;


                CompleteCurrentStepEtapa2(currentStep, indicator, pesoTextBlock, cantidadTextBlock, pieceWeight, "OK", codigo);

                runningWeight = weight;
                stepIndex++;

                if (stepIndex >= pasosFiltrados.Count)
                {
                    LogFerreteriaStep(currentStep, "OK", codigo);
                    EndFerreteriaG_ArmHarness();
                }


            }
            else
            {
                var indicator = FindName(currentStep.Part_Indicator) as Rectangle;
                var pesoTextBlock = FindName(currentStep.Part_Peso) as TextBlock;

                if (indicator != null) indicator.Fill = Brushes.Red;
                if (pesoTextBlock != null) pesoTextBlock.Text = "Fuera de rango";

                currentStep.DetectedWeight = pieceWeight.ToString("F5");
            }
        }


        private void CompleteCurrentStepEtapa2(SequenceStep currentStep, Rectangle indicator, TextBlock pesoTextBlock,
            TextBlock cantidadTextBlock, double currentWeight, string logMessage, string tags)
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
                {
                    //_currentStepIndex++;

                    if (_currentStepIndex < pasosFiltrados.Count)
                    {
                        //var nextStep = pasosFiltrados[_currentStepIndex];
                        pieceWeight = currentWeight - _accumulatedWeight;

                        /*ShowBolsasRestantes(nextStep.Part_NoParte, nextStep.MinWeight, nextStep.MaxWeight,
                            pieceWeight, _validacion, int.Parse(nextStep.Part_Cantidad));*/
                        return;
                    }
                }
            }
        }


        private async void ReadInputBascula1()
        {
            await Task.Run(() =>
            {
                var serialPort = new SerialPort(defaultSettings.PuertoBascula1, defaultSettings.BaudRateBascula12,
                    Parity.None, defaultSettings.DataBitsBascula12, StopBits.One);
                bascula1.AsignarPuertoBascula(serialPort);
                bascula1.OpenPort();
            });
        }

        private async void ReadInputBascula2()
        {
            await Task.Run(() =>
            {
                var serialPort = new SerialPort(defaultSettings.PuertoBascula2, defaultSettings.BaudRateBascula12,
                    Parity.None, defaultSettings.DataBitsBascula12, StopBits.One);
                bascula2.AsignarPuertoBascula(serialPort);
                bascula2.OpenPort();
            });
        }

        #endregion

        #region VISUALES

        private void HideParts(bool isHidden, List<SequenceStep> steps)
        {
            foreach (var partGrid in steps.Select(step => FindName(step.grdPart)).OfType<UIElement>())
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
                var imageControl = FindName(step.Part_Imagen) as Image;
                if (imageControl != null)
                {
                    var imagePath = Path.Combine(rutaImagenes, $"{step.Part_NoParte}.PNG");
                    var partCantidad = FindName($"Part_Cantidad{int.Parse(step.Part_Orden) - 1}") as TextBlock;

                    if (partCantidad != null)
                        partCantidad.Text = step.Part_Cantidad;

                    if (File.Exists(imagePath))
                        imageControl.Source = new BitmapImage(new Uri(imagePath, UriKind.Absolute));
                    else
                        imageControl.Source = null; // No se encontró la imagen, dejar vacío
                }

                var ledLabel = FindName(step.Part_Indicator) as UIElement;
                var ledPeso = FindName(step.Part_Peso) as UIElement;

                if (ledLabel != null) ledLabel.Visibility = Visibility.Visible;
                if (ledPeso != null) ledPeso.Visibility = Visibility.Visible;

                var partPesoIndex = step.Part_Peso.Replace("Part_Peso", "");
                var partNoTextBlockName = $"Part_NoParte{partPesoIndex}";

                var partNoTextBlock = FindName(partNoTextBlockName) as TextBlock;
                if (partNoTextBlock != null)
                    partNoTextBlock.Text = step.Part_NoParte;
            }
        }

        private void SetValuesEtapas(List<SequenceStep> steps, int etapa)
        {
            if (ModeloData.UsaBascula1 && ModeloData.UsaBascula2 || ModeloData.UsaBascula1 && !ModeloData.UsaBascula2)
                switch (etapa)
                {
                    case 1:
                        var etapa1 = steps.FirstOrDefault(step => step.Part_NoParte.Contains(ModeloData.Etapa1));

                        _etapa1 = true;
                        if (etapa1 != null)
                            pasosFiltrados = steps
                                .Where(s => int.Parse(s.Part_Orden) <= int.Parse(etapa1.Part_Orden))
                                .ToList();
                        break;

                    case 2:

                        var startStep = steps.FirstOrDefault(step => step.Part_NoParte.Contains(ModeloData.Etapa2));

                        _etapa2 = true;
                        if (startStep != null)
                            pasosFiltrados = steps
                                .Where(s => int.Parse(s.Part_Orden) >= int.Parse(startStep.Part_Orden))
                                .ToList();
                        break;

                    default:
                        throw new ArgumentException($"Etapa '{etapa}' no es válida.");
                }

            else if (!ModeloData.UsaBascula1 && ModeloData.UsaBascula2)
                switch (etapa)
                {
                    case 1:
                        var etapa1 = steps.FirstOrDefault(step => step.Part_NoParte.Contains(ModeloData.Etapa2));

                        _etapa1 = true;
                        if (etapa1 != null)
                            pasosFiltrados = steps
                                .Where(s => int.Parse(s.Part_Orden) <= int.Parse(etapa1.Part_Orden))
                                .ToList();
                        break;
                    default:
                        throw new ArgumentException($"Etapa '{etapa}' no es válida.");
                }

            foreach (var step in pasosFiltrados)
            {
                var imageControl = FindName(step.Part_Imagen) as Image;
                if (imageControl != null)
                {
                    var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes",
                        $"{step.Part_NoParte}.PNG");
                    var partCantidad = FindName($"Part_Cantidad{int.Parse(step.Part_Orden) - 1}") as TextBlock;

                    if (partCantidad != null)
                        partCantidad.Text = step.Part_Cantidad;

                    if (File.Exists(imagePath))
                        imageControl.Source = new BitmapImage(new Uri(imagePath, UriKind.Absolute));
                    else
                        imageControl.Source = null;
                }

                var ledLabel = FindName(step.Part_Indicator) as UIElement;
                var ledPeso = FindName(step.Part_Peso) as UIElement;

                if (ledLabel != null) ledLabel.Visibility = Visibility.Visible;
                if (ledPeso != null) ledPeso.Visibility = Visibility.Visible;

                var partPesoIndex = step.Part_Peso.Replace("Part_Peso", "");
                var partNoTextBlockName = $"Part_NoParte{partPesoIndex}";
                var partNoTextBlock = FindName(partNoTextBlockName) as TextBlock;

                if (partNoTextBlock != null)
                    partNoTextBlock.Text = step.Part_NoParte;
            }
        }


        private void ReindexarPasos(List<SequenceStep> pasos)
        {
            for (var i = 0; i < pasos.Count; i++)
            {
                var step = pasos[i];

                step.Part_Imagen = $"Part_Imagen{i}";
                step.Part_Indicator = $"Part_Indicator{i}";
                step.Part_Peso = $"Part_Peso{i}";
                step.grdPart = $"Part{i}";

                var imageControl = FindName(step.Part_Imagen) as Image;
                var pesoControl = FindName(step.Part_Peso) as TextBlock;
                var indicatorControl = FindName(step.Part_Indicator) as Rectangle;
                var gridControl = FindName(step.grdPart) as Grid;

                if (imageControl != null)
                {
                    imageControl.Visibility = Visibility.Visible;
                    imageControl.Source = null;
                }

                if (pesoControl != null)
                {
                    pesoControl.Visibility = Visibility.Visible;
                    pesoControl.Text = "0.0 kg";
                }

                if (indicatorControl != null)
                {
                    indicatorControl.Visibility = Visibility.Visible;
                    indicatorControl.Fill = Brushes.Red;
                }

                if (gridControl != null) gridControl.Visibility = Visibility.Visible;
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
                var imageControl = FindName(imageControlName) as Image;
                var textNoParte = FindName($"Part_NoParte{counter}") as TextBlock;
                var partCantidad = FindName($"Part_Cantidad{counter}") as TextBlock;
                var part = FindName($"Part{counter}") as Grid;
                var indicador = FindName($"Part_Indicator{counter}") as Rectangle;
                var peso = FindName($"Part_Peso{counter}") as TextBlock;

                if (indicador != null && peso != null)
                {
                    indicador.Fill = Brushes.Red;
                    peso.Text = "0.0 kg";
                }

                if (textNoParte != null)
                    textNoParte.Text = step.Part_NoParte;

                if (imageControl != null)
                {
                    var imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Imagenes", $"{step.Part_NoParte}.PNG");

                    if (partCantidad != null)
                        partCantidad.Text = step.Part_Cantidad;

                    if (part != null)
                        part.Visibility = Visibility.Visible;

                    if (File.Exists(imagePath))
                    {
                        imageControl.Source = new BitmapImage(new Uri(imagePath, UriKind.Absolute));
                    }
                    else
                    {
                        imageControl.Source = null;
                    }
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
            txbScanner.KeyDown += txbScanner_KeyDown;
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
                TitleText = "Error",
                Message = mensaje
            };

            _errorWindow.Show();
        }

        private async Task ShowMensaje(string mensaje, Brush color, int time)
        {
            await ShowCustomMessage(mensaje, color, time);
        }

        public static async Task ShowCustomMessage(string message, Brush color, int time)
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
                    throw new ArgumentException($"Tipo de dispensadora '{tipo}' no válido.");
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
                    worksheet.Cell(newRow, 2).Value = step.Part_NoParte;
                    worksheet.Cell(newRow, 3).Value = step.ModProceso;
                    worksheet.Cell(newRow, 4).Value = step.Part_Proceso;
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

        private void LogFinal()
        {
            using (var workbook = new XLWorkbook(_catalogos.filePathExcel))
            {
                var worksheet = workbook.Worksheet("Completados");

                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                var newRow = lastRow + 1;

                worksheet.Cell(newRow, 1).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                worksheet.Cell(newRow, 2).Value = "INFORMACIÓN FINAL";
                worksheet.Cell(newRow, 3).Value = $"FERRETERÍA: {_completedSequencesFerre}";
                worksheet.Cell(newRow, 4).Value = $"PIEZAS TOTALES: {_completedSequencesFerre}";

                workbook.Save();
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
                    worksheet.Cell(newRow, 2).Value = step.Part_NoParte;
                    worksheet.Cell(newRow, 3).Value = step.ModProceso;
                    worksheet.Cell(newRow, 4).Value = step.Part_Proceso;
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

        #region FERRETERIA

        private void ProcessCajaFerreteria(double weight, bool isStable)
        {
            var cajaStep = pasosFiltrados.LastOrDefault();
            if (cajaStep == null) return;


            var pesoTotalBolsas = valoresBolsas.Sum(s => double.Parse(s.DetectedWeight));
            var listaTags = valoresBolsas.Select(b => b.Tag).Where(tag => !string.IsNullOrEmpty(tag)).ToList();
            var pesoMinimo = pesoTotalBolsas + cajaStep.MinWeight + (0.01 * (pesoTotalBolsas + cajaStep.MinWeight));
            var pesoMaximo = pesoTotalBolsas + cajaStep.MaxWeight + (0.01 * (pesoTotalBolsas + cajaStep.MaxWeight));

            if (isStable)
            {
                if (!_startMeasurinBoxAndBags)
                {
                    ShowAlertaPeso("PESAR SOLO CAJA MASTER", cajaStep.MinWeight, cajaStep.MaxWeight, weight,
                        _validacion);

                    if (weight >= cajaStep.MinWeight && weight <= cajaStep.MaxWeight)
                    {
                        cajaStep.DetectedWeight = weight.ToString();
                        LogFerreteriaStep(cajaStep, "CAJA VACÍA OK", null);

                        _startMeasurinBoxAndBags = true;
                    }
                }
                else
                {
                    ShowAlertaPeso("PESAR CAJA MASTER CON CAJAS UNITARIAS", pesoMinimo, pesoMaximo, weight,
                        _validacion);

                    if (weight >= pesoMinimo && weight <= pesoMaximo)
                    {
                        cajaStep.DetectedWeight = weight.ToString("F5");
                        var allTags = string.Join(", ", listaTags);
                        LogFerreteriaStep(cajaStep, "CAJA COMPLETA OK", allTags);

                        _stopBascula2 = true;
                        EndFerreteriaG();
                    }
                }
            }

            var cajaPeso = FindName("Part_Peso0") as TextBlock;
            if (cajaPeso != null)
                cajaPeso.Text = weight.ToString();
        }

        private void ProcessSequenceFerreteria(double currentWeight, bool isStable)
        {
            _consecutiveCount = 0;

            //Contador valor caja + modeloseleccionado.UsaConteoCajas

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

            //if ((Cbox_Modelo.SelectionBoxItem.ToString() == "ANTENNA KIT" || Cbox_Modelo.SelectionBoxItem.ToString() == "123-0253-000") && _etapa1)
            if (_etapa1)
                ProcessStableWeight(currentWeight, isStable);
            //ProcessStableWeight(currentWeight, isStable);
        }

        private void ProcessStableWeight(double currentWeight, bool isStable)
        {
            if (_FerreteriaPart2) return;

            var currentStep = pasosFiltrados[_currentStepIndex];
            pieceWeight = currentWeight - _accumulatedWeight;

            var indicator = FindName(currentStep.Part_Indicator) as Rectangle;
            var pesoTextBlock = FindName(currentStep.Part_Peso) as TextBlock;
            var cantidadTextBlock = FindName($"Part_Cantidad{int.Parse(currentStep.Part_Orden) - 1}") as TextBlock;

            if (!isStable || cantidadTextBlock == null) return;

            ShowBolsasRestantes(currentStep.Part_NoParte, currentStep.MinWeight, currentStep.MaxWeight,
                pieceWeight,
                _validacion, int.Parse(cantidadTextBlock.Text));

            if (indicator == null || pesoTextBlock == null) return;

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
                        if (!currentStep.IsCompleted && Cbox_Modelo.Text == "ANTENNA KIT")
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
                                Part_NoParte = currentStep.Part_NoParte,
                                MinWeight = currentStep.MinWeight,
                                MaxWeight = currentStep.MaxWeight,
                                DetectedWeight = currentWeight.ToString(),
                                Tag = codigo,
                                IsCompleted = false,
                                Part_Indicator = currentStep.Part_Indicator,
                                Part_Peso = currentStep.Part_Peso,
                                Part_Orden = currentStep.Part_Orden,
                                Part_Cantidad = currentStep.Part_Cantidad,
                                ModProceso = currentStep.ModProceso,
                                Part_Proceso = currentStep.Part_Proceso
                            });

                            //ReadScannerBolsa();
                        }

                        _currentStepIndex++;

                        if (_currentStepIndex < pasosFiltrados.Count)
                        {
                            currentStep = pasosFiltrados[_currentStepIndex];
                            pieceWeight = currentWeight - _accumulatedWeight;

                            ShowBolsasRestantes(currentStep.Part_NoParte, currentStep.MinWeight, currentStep.MaxWeight,
                                pieceWeight,
                                _validacion, int.Parse(currentStep.Part_Cantidad));
                            return;
                        }

                        var lastBag = valoresBolsas.LastOrDefault();
                        if (lastBag != null)
                        {
                            lastBag.Part_NoParte = currentStep.Part_NoParte;
                            lastBag.MinWeight = currentStep.MinWeight;
                            lastBag.MaxWeight = currentStep.MaxWeight;
                            lastBag.DetectedWeight = currentWeight.ToString("F5");
                            lastBag.Part_Indicator = currentStep.Part_Indicator;
                            lastBag.Part_Peso = currentStep.Part_Peso;
                            lastBag.Part_Orden = currentStep.Part_Orden;
                            lastBag.Part_Cantidad = currentStep.Part_Cantidad;
                            lastBag.Part_Proceso = currentStep.Part_Proceso;
                        }

                        _currentStepIndex = 0;
                        _accumulatedWeight = 0;
                        pieceWeight = 0;
                        SetImagesBox();
                        contador++;

                        if(ModeloData.UsaConteoCajas)
                            lblProgreso.Content = contador;

                        else if (!ModeloData.UsaConteoCajas)
                            lblCompletados.Content = contador;


                        if (btnEtiquetaManual.Tag?.ToString() == "off")
                        {
                            if (Cbox_Modelo.Text == "123-0253-000")
                            {
                                ActivarSalida("prop65");

                                (zpl, integrer, fraction) = ZebraPrinter.GenerateZplBody(Cbox_Modelo.SelectedValue.ToString());
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
                pesoTextBlock.Text = "Fuera de rango";
                currentStep.DetectedWeight = currentWeight.ToString("F5");
            }
        }

        private async void txbScanner_KeyDown(object sender, KeyEventArgs e)
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
                    txbScanner.KeyDown -= txbScanner_KeyDown;

                    await ShowMensaje("SECUENCIA CORRECTA, PUEDE CERRAR LA BOLSA CON LA SELLADORA", Brushes.Green, 2500);
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
                    .Select(s => s.Part_Proceso)
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

        private async void EndFerreteriaG()
        {
            await Dispatcher.InvokeAsync(ShowPruebaCorrecta);

            await Task.Delay(2000);
            if (btnEtiquetaManual.Tag?.ToString() == "off")
                ActivarSalida("master");

            _stopBascula1 = false;

            _FerreteriaPart2 = false;
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

            HideAll();
            ProcesarModeloValido(Cbox_Modelo.SelectedValue.ToString());
            pasosFiltrados = ObtenerValoresProceso(Cbox_Modelo.SelectedValue.ToString(),
                1);
            ProcesarModeloYProceso(Cbox_Modelo.SelectedValue.ToString());
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

            _FerreteriaPart2 = false;
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

            HideAll();
            ProcesarModeloValido(Cbox_Modelo.SelectedValue.ToString());
            sequence = ObtenerValoresProceso(Cbox_Modelo.SelectedValue.ToString(), 1);
            ProcesarModeloYProceso(Cbox_Modelo.SelectedValue.ToString());
            Cbox_Proceso.SelectedIndex = 0;
            SetValuesEtapas(sequence, 1);
            SetImagesBox();
        }

        #endregion
    }
}