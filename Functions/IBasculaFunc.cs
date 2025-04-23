using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace Scale_Program.Functions
{
    public interface IBasculaFunc
    {
        void AsignarPuertoBascula(SerialPort port);
        void AsignarControles(Dispatcher dispatcher);
        void OpenPort();
        void ClosePort();
        void EnviarZero();
        SerialPort GetPuerto();
        void EnviarComandoABascula(string comando);
        event EventHandler<BasculaEventArgs> OnDataReady;
    }

}
