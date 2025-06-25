using System;
using log4net;

namespace Scale_Program.Functions
{
    public sealed class IOInterface : IDisposable
    {
        public const int MaxInputs = 16; // Número máximo de entradas
        public const int MaxOutputs = 16; // Número máximo de salidas
        private readonly ILog Logger = LogManager.GetLogger("IOInterface");
        private readonly ISeaLevelDevice seaLevel;

        public IOInterface(ISeaLevelDevice device)
        {
            seaLevel = device ?? throw new ArgumentNullException(nameof(device));
        }

        public void Dispose()
        {
            seaLevel.Dispose();
        }

        public uint ReadAllInputs()
        {
            try
            {
                return seaLevel.ReadDiscreteInputs(0, MaxInputs);
            }
            catch (Exception ex)
            {
                Logger.Error("ReadAllInputs", ex);
                return 0;
            }
        }

        public bool ReadSingleInput(int inputNumber)
        {
            if (inputNumber < 0 || inputNumber >= MaxInputs)
                throw new ArgumentOutOfRangeException(nameof(inputNumber));

            try
            {
                return seaLevel.ReadDiscreteInputs(inputNumber, 1) != 0;
            }
            catch (Exception ex)
            {
                Logger.Error("ReadSingleInput", ex);
                return false;
            }
        }

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

        public void WriteMultipleOutputs(int startOutput, ushort value, int count)
        {
            if (startOutput < 0 || startOutput >= MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(startOutput));

            if (startOutput + count > MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(count));

            try
            {
                seaLevel.WriteMultipleCoils(startOutput, value, count);
            }
            catch (Exception ex)
            {
                Logger.Error("WriteMultipleOutputs", ex);
            }
        }

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
    }
}