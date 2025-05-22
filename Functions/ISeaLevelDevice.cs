using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scale_Program.Functions
{
    public interface ISeaLevelDevice : IDisposable
    {
        void SetSingleCoilState(int index, bool active);
        void WriteMultipleCoils(int startIndex, int value, int count);
        uint ReadDiscreteInputs(int startIndex, int count);
    }

}
