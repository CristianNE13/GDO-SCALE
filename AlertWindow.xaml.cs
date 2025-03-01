using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Brushes = System.Windows.Media.Brushes;

namespace Scale_Program
{
    /// <summary>
    ///     Interaction logic for AlertWindow.xaml
    /// </summary>
    public partial class AlertWindow : Window
    {
        public enum AlertResult
        {
            Reset,
            Continue
        }

        public AlertWindow()
        {
            InitializeComponent();
            MouseDown += AlertWindow_MouseDown;
        }

        public bool ShowCompletion { get; set; }
        public bool ShowSensorNeeded { get; set; }
        public bool ShowPiezaNeeded { get; set; }
        public string PiezaName { get; set; }
        public double PiezaPesoMin { get; set; }
        public double PiezaPesoMax { get; set; }
        public double PiezaPeso { get; set; }

        /// <summary>
        ///     Gets or set the result
        /// </summary>
        public AlertResult Result { get; private set; }

        private void AlertWindow_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void OnResetButtonClick(object sender, RoutedEventArgs e)
        {
            Result = AlertResult.Reset;
            Close();
        }

        private void OnContinueButtonClick(object sender, RoutedEventArgs e)
        {
            Result = AlertResult.Continue;
            Close();
        }

        public void ShowCompleteAndClose()
        {
            ShowCompletion = true;
        }

        public void ShowNeedSensor()
        {
            ShowSensorNeeded = true;
        }

        public void ShowPieza(string name, double pesoMin, double pesoMax, double piezaPeso)
        {
            PiezaName = name;
            PiezaPesoMin = pesoMin;
            PiezaPesoMax = pesoMax;
            PiezaPeso = piezaPeso;
            ShowPiezaNeeded = true;
        }

        public void ShowNeedSensorFerre()
        {
            lblSecuencia.Content = "INTRODUCE LA BOLSA A LA CAJA";
            lblAlert.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Hidden;
            btnContinuar.Visibility = Visibility.Hidden;
            grdAlertWindow.Background = Brushes.LightGreen;
        }

        private async void ShowCompleteAndCloseInternal()
        {
            lblSecuencia.Visibility = Visibility.Hidden;
            ImgAlerta.Visibility = Visibility.Hidden;
            lblProceso.Content = "SECUENCIA COMPLETADA";
            lblAlert.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Hidden;
            btnContinuar.Visibility = Visibility.Hidden;
            grdAlertWindow.Background = Brushes.LightGreen;

            await Task.Delay(3000);
            Close();
        }

        public async void ShowPruebaCorrecta()
        {
            lblSecuencia.Visibility = Visibility.Hidden;
            lblProceso.Content = "SECUENCIA CORRECTA";
            lblProceso.Visibility = Visibility.Visible;
            ImgAlerta.Visibility = Visibility.Hidden;
            lblAlert.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Hidden;
            btnContinuar.Visibility = Visibility.Hidden;
            grdAlertWindow.Background = Brushes.LightGreen;

            await Task.Delay(3000);
            Close();
        }

        public async void ShowScanner()
        {
            lblSecuencia.Visibility = Visibility.Hidden;
            lblProceso.Content = "SCANEO CORRECTO";
            lblProceso.Visibility = Visibility.Visible;
            ImgAlerta.Visibility = Visibility.Hidden;
            lblAlert.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Hidden;
            btnContinuar.Visibility = Visibility.Hidden;
            grdAlertWindow.Background = Brushes.LightGreen;

            await Task.Delay(4000);
            Close();
        }

        private void ShowNeedSensorInternal()
        {
            lblSecuencia.Content = "Insertar la flecha";
            lblAlert.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Hidden;
            btnContinuar.Visibility = Visibility.Hidden;
        }

        private void ShowPiezaInternal()
        {
            ImgAlerta.Visibility = Visibility.Hidden;
            grdAlertWindow.Background = Brushes.Yellow;

            lblAlert.Content = "FAVOR DE PESAR EL ARTICULO";
            lblSecuencia.Content = PiezaName;

            lblPesoMin.Visibility = Visibility.Visible;
            lblPesoMax.Visibility = Visibility.Visible;
            lblPesoActual.Visibility = Visibility.Visible;

            btnContinuar.Visibility = Visibility.Hidden;
            btnReset.Visibility = Visibility.Hidden;

            lblPesoMin.Content = $"Peso Min: {PiezaPesoMin}kg";
            lblPesoMax.Content = $"Peso Max: {PiezaPesoMax}kg";
            lblPesoActual.Content = $"Peso Actual:{PiezaPeso:F3}kg";
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (ShowCompletion)
                ShowCompleteAndCloseInternal();
            else if (ShowSensorNeeded)
                ShowNeedSensorInternal();
            else if (ShowPiezaNeeded) ShowPiezaInternal();
        }
    }
}