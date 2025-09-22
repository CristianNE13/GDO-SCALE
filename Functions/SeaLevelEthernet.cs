using System;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows;
using Modbus.Device;

namespace Scale_Program.Functions
{
    public class SeaLevelEthernet : ISeaLevelDevice
    {
        private const int MaxInputs = 16;
        private const int MaxOutputs = 16;

        private readonly string _ip;
        private readonly object _modbusLock = new object();
        private readonly int _port;
        private readonly byte _slaveId;
        private IModbusMaster _modbusMaster;
        private TcpClient _tcpClient;

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
            _slaveId = 247; //default id
        }

        public bool IsConnected { get; private set; }

        public void Dispose()
        {
            try
            {
                _tcpClient?.Close();
                _tcpClient?.Dispose();
                _tcpClient = null;
                _modbusMaster = null;
                IsConnected = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al liberar recursos: {ex.Message}", "ERROR");
            }
        }

        public uint ReadDiscreteInputs(int startIndex, int count)
        {
            try
            {
                return EjecutarConReconexion(() =>
                {
                    var inputs = _modbusMaster.ReadInputs(_slaveId, (ushort)startIndex, (ushort)count);
                    uint result = 0;
                    for (var i = 0; i < inputs.Length; i++)
                        if (inputs[i])
                            result |= (uint)(1 << i);
                    return result;
                }, "ReadDiscreteInputs");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en ReadDiscreteInputs: {ex.Message}", "ERROR");
                return 0;
            }
        }

        public void SetSingleCoilState(int coilIndex, bool state)
        {
            try
            {
                EjecutarConReconexion(() =>
                {
                    _modbusMaster.WriteSingleCoil(_slaveId, (ushort)coilIndex, state);
                    return true;
                }, "SetSingleCoilState");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en SetSingleCoilState: {ex.Message}", "ERROR");
            }
        }

        public void WriteMultipleCoils(int startIndex, int value, int count)
        {
            try
            {
                EjecutarConReconexion(() =>
                {
                    var bits = new bool[count];
                    for (var i = 0; i < count; i++)
                        bits[i] = (value & (1 << i)) != 0;

                    _modbusMaster.WriteMultipleCoils(_slaveId, (ushort)startIndex, bits);
                    return true;
                }, "WriteMultipleCoils");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en WriteMultipleCoils: {ex.Message}", "ERROR");
            }
        }

        public bool Connect()
        {
            try
            {
                _tcpClient = new TcpClient(_ip, _port);

                if (_tcpClient.Connected)
                {
                    _modbusMaster = ModbusIpMaster.CreateIp(_tcpClient);
                    IsConnected = true;
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al conectar SeaLevel: {ex.Message}");
                IsConnected = false;
                return false;
            }
        }


        private T EjecutarConReconexion<T>(Func<T> operacion, string contexto)
        {
            lock (_modbusLock)
            {
                try
                {
                    return operacion();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[{contexto}] Error: {ex.Message}. Intentando reconexión...");

                    Dispose();

                    try
                    {
                        Connect();
                        return operacion();
                    }
                    catch (Exception reconEx)
                    {
                        Console.WriteLine($"[{contexto}] Falla en reconexión: {reconEx.Message}");
                        throw;
                    }
                }
            }
        }
    }
}