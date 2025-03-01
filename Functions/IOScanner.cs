using System;
using System.Diagnostics;
using System.Threading;

namespace Scale_Program.Functions
{
    public sealed class IOScanner
    {
        private const int TickPeriod = 50; // Intervalo en milisegundos
        private readonly Stopwatch freeRunningTimer; // Temporizador interno para medición de tiempos
        private readonly IOInterface io; // Interfaz para manejar el hardware
        private Timer tickTimer; // Temporizador para escaneo periódico

        // Constructor
        public IOScanner(IOInterface io)
        {
            this.io = io;
            freeRunningTimer = new Stopwatch();
        }

        // Sensores específicos
        public string PortName { get; set; } // Nombre del puerto utilizado
        public bool IsSensor1Activated { get; private set; } // Estado del sensor 1
        public bool IsSensor2Activated { get; private set; } // Estado del sensor 2
        public int InputSensor1 { get; set; } // Número del sensor 1

        public int InputSensor2 { get; set; } // Número del sensor 2
        public int InputSensor3 { get; set; }

        public bool IsRunning()
        {
            return tickTimer != null;
        }

        // Evento para notificar cambios en las entradas
        public event EventHandler<uint> Tick;

        /// <summary>
        ///     Inicia el escaneo periódico de entradas.
        /// </summary>
        public void Start()
        {
            if (tickTimer != null) Stop();

            tickTimer = new Timer(OnTimerTick);
            tickTimer.Change(0, TickPeriod); // Configura el temporizador para iniciar de inmediato
        }

        /// <summary>
        ///     Detiene el escaneo periódico de entradas.
        /// </summary>
        public void Stop()
        {
            if (tickTimer != null)
            {
                tickTimer.Dispose();
                tickTimer = null;

                Thread.Sleep(TickPeriod * 2); // Asegura que no haya tareas pendientes
            }

            freeRunningTimer.Stop();
            IsSensor1Activated = false;
            IsSensor2Activated = false;
        }

        /// <summary>
        ///     Verifica la conexión con el dispositivo IO.
        /// </summary>
        /// <returns>True si está conectado, false de lo contrario.</returns>
        public bool Verify()
        {
            try
            {
                io.ReadAllInputs(); // Verifica lectura de entradas
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        ///     Lógica ejecutada periódicamente por el temporizador.
        /// </summary>
        private void OnTimerTick(object state)
        {
            try
            {
                // Lee las entradas del IOInterface
                var inputsState = io.ReadAllInputs();

                // Dispara el evento `Tick` para notificar cambios
                Tick?.Invoke(this, inputsState);

                // Actualiza los estados de los sensores
                UpdateSensorStates(inputsState);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en OnTimerTick: {ex.Message}");
            }
            finally
            {
                // Reinicia el temporizador
                tickTimer?.Change(TickPeriod, Timeout.Infinite);
            }
        }

        /// <summary>
        ///     Actualiza los estados de los sensores configurados.
        /// </summary>
        private void UpdateSensorStates(uint inputsState)
        {
            // Verifica el estado del sensor 1
            if (InputSensor1 >= 0 && InputSensor1 < IOInterface.MaxInputs)
                IsSensor1Activated = (inputsState & (1 << InputSensor1)) != 0;

            // Verifica el estado del sensor 2
            if (InputSensor2 >= 0 && InputSensor2 < IOInterface.MaxInputs)
                IsSensor2Activated = (inputsState & (1 << InputSensor2)) != 0;
        }
    }
}