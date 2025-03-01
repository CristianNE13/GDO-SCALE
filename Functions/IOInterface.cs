using System;
using log4net;

namespace Scale_Program.Functions
{
    public sealed class IOInterface : IDisposable
    {
        private const int IOBaudRate = 115200; // Configuración del puerto serie
        public const int MaxInputs = 16; // Número máximo de entradas
        public const int MaxOutputs = 16; // Número máximo de salidas
        private readonly ILog Logger = LogManager.GetLogger("IOInterface");

        private readonly SeaLevel seaLevel;

        /// <summary>
        ///     Inicializa una nueva instancia de la clase <see cref="IOInterface" />.
        /// </summary>
        /// <param name="portName">Nombre del puerto a usar</param>
        /// <exception cref="IOException">El puerto está en un estado inválido</exception>
        /// <exception cref="UnauthorizedAccessException">Acceso denegado al puerto</exception>
        public IOInterface(string portName)
        {
            seaLevel = new SeaLevel(247, portName, IOBaudRate);
            seaLevel.Open();
        }

        /// <summary>
        ///     Libera los recursos utilizados por la clase.
        /// </summary>
        public void Dispose()
        {
            seaLevel.Dispose();
        }

        /// <summary>
        ///     Obtiene el estado de todas las entradas del dispositivo IO.
        /// </summary>
        /// <returns>Estado de todas las entradas como un entero sin signo</returns>
        public uint ReadAllInputs()
        {
            uint value = 0;

            try
            {
                value = seaLevel.ReadDiscreteInputs(0, MaxInputs);
            }
            catch (Exception ex)
            {
                Logger.Error("ReadAllInputs", ex);
            }

            return value;
        }

        /// <summary>
        ///     Obtiene el estado de una entrada específica.
        /// </summary>
        /// <param name="inputNumber">Número de la entrada (0 o 1)</param>
        /// <returns>True si está activa, false de lo contrario</returns>
        public bool ReadSingleInput(int inputNumber)
        {
            if (inputNumber < 0 || inputNumber >= MaxInputs) throw new ArgumentOutOfRangeException(nameof(inputNumber));

            var status = false;

            try
            {
                status = seaLevel.ReadDiscreteInputs(inputNumber, 1) != 0;
            }
            catch (Exception ex)
            {
                Logger.Error("ReadSingleInput", ex);
            }

            return status;
        }

        /// <summary>
        ///     Establece el estado de una salida específica.
        /// </summary>
        /// <param name="outputNumber">Número de la salida (0 a 15)</param>
        /// <param name="active">True para activar, false para desactivar</param>
        public void WriteSingleOutput(int outputNumber, bool active)
        {
            if (outputNumber < 0 || outputNumber >= MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(outputNumber));

            try
            {
                seaLevel.SetSingleCoilState(outputNumber, active);
            }
            catch (Exception ex)
            {
                Logger.Error("WriteSingleOutput", ex);
            }
        }

        /// <summary>
        ///     Verifica la conexión con el dispositivo.
        /// </summary>
        /// <returns>True si la verificación es exitosa, false de lo contrario</returns>
        public bool Verify()
        {
            try
            {
                ReadAllInputs();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void WriteMultipleOutputs(int startOutput, ushort value, int count)
        {
            if (startOutput < 0 || startOutput >= MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(startOutput),
                    "El índice de salida inicial está fuera del rango.");

            if (startOutput + count > MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(count),
                    "La cantidad de salidas supera el número máximo permitido.");

            try
            {
                seaLevel.WriteMultipleCoils(startOutput, value, count);
            }
            catch (Exception ex)
            {
                Logger.Error("WriteMultipleOutputs", ex);
            }
        }
    }
}