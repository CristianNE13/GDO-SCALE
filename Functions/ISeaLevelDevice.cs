using System;

namespace Scale_Program.Functions
{
    public interface ISeaLevelDevice : IDisposable
    {
        void SetSingleCoilState(int index, bool active);
        void WriteMultipleCoils(int startIndex, int value, int count);
        uint ReadDiscreteInputs(int startIndex, int count);
        void Dispose();
    }
}