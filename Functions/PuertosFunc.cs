using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;

namespace Scale_Program.Functions
{
    public class PuertosFunc
    {
        private const int DefaultBaudRate = 9600;

        public List<string> GetPorts()
        {
            List<string> ports = null;
            ports = SerialPort.GetPortNames().ToList();
            return ports;
        }

        public SerialPort InicializarSerialPort(string puerto)
        {
            return InicializarSerialPort(puerto, DefaultBaudRate);
        }

        public SerialPort InicializarSerialPort(string puerto, int baudRate)
        {
            return new SerialPort(puerto, baudRate);
        }
    }
}