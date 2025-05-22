using System;
using System.Net.Sockets;
using Modbus.Device; // NSModbus4

namespace Scale_Program.Functions
{
    public class SeaLevelEthernet : ISeaLevelDevice
    {
        private const int MaxInputs = 16;
        private const int MaxOutputs = 16;

        private readonly string _ip;
        private readonly int _port;
        private readonly byte _slaveId;
        private TcpClient _tcpClient;
        private IModbusMaster _modbusMaster;
        private bool _isConnected;
        public bool IsConnected => _isConnected;

        public SeaLevelEthernet(string ipAddress, int slaveId, int port)
        {
            _ip = ipAddress;
            _port = port;
            _slaveId = (byte)slaveId;
        }

        public SeaLevelEthernet(string ipAddress, int slaveId)
        {
            _ip = ipAddress;
            _port = 502; //Modbus default port
            _slaveId = (byte)slaveId;
        }

        public SeaLevelEthernet(string ipAddress)
        {
            _ip = ipAddress;
            _port = 502; //Modbus default port
            _slaveId = 247;
        }

        public void Connect()
        {
            if (_isConnected)
                throw new InvalidOperationException("La conexión ya está abierta.");

            _tcpClient = new TcpClient(_ip, _port);
            _modbusMaster = ModbusIpMaster.CreateIp(_tcpClient);
            _isConnected = true;
        }

        public void Dispose()
        {
            if (_tcpClient != null && _tcpClient.Connected)
            {
                _tcpClient.Close();
                _isConnected = false;
            }
        }


        public void SetSingleCoilState(int outputNumber, bool state)
        {
            if (outputNumber < 0 || outputNumber >= MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(outputNumber));

            _modbusMaster.WriteSingleCoil(_slaveId, (ushort)outputNumber, state);
        }

        public void WriteMultipleCoils(int startIndex, int value, int count)
        {
            if (count <= 0 || count > MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(count));

            if (startIndex < 0 || startIndex + count > MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(startIndex));

            var bits = new bool[count];
            for (int i = 0; i < count; i++)
                bits[i] = (value & (1 << i)) != 0;

            _modbusMaster.WriteMultipleCoils(_slaveId, (ushort)startIndex, bits);
        }

        public uint ReadDiscreteInputs(int startIndex, int count)
        {
            if (startIndex < 0 || startIndex >= MaxInputs)
                throw new ArgumentOutOfRangeException(nameof(startIndex));

            if (startIndex + count > MaxInputs)
                throw new ArgumentOutOfRangeException(nameof(count));

            var inputs = _modbusMaster.ReadInputs(_slaveId, (ushort)startIndex, (ushort)count);
            ushort result = 0;
            for (int i = 0; i < inputs.Length; i++)
            {
                if (inputs[i])
                    result |= (ushort)(1 << i);
            }
            return result;
        }

        public ushort ReadCoils(int startIndex, int count)
        {
            if (startIndex < 0 || startIndex >= MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(startIndex));

            if (startIndex + count > MaxOutputs)
                throw new ArgumentOutOfRangeException(nameof(count));

            var coils = _modbusMaster.ReadCoils(_slaveId, (ushort)startIndex, (ushort)count);
            ushort result = 0;
            for (int i = 0; i < coils.Length; i++)
            {
                if (coils[i])
                    result |= (ushort)(1 << i);
            }
            return result;
        }
    }
}