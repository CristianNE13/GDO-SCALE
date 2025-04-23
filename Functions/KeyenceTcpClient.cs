using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace Scale_Program.Functions
{
    public class KeyenceTcpClient : IDisposable
    {
        private TcpClient _client;
        private NetworkStream _stream;

        public string IpAddress { get; }
        public int Port { get; }

        public KeyenceTcpClient(string ipAddress, int port = 8500)
        {
            IpAddress = ipAddress;
            Port = port;
        }

        public async Task<bool> ConnectAsync()
        {
            if (_client != null && _client.Connected)
                return true;

            try
            {
                _client = new TcpClient();
                await _client.ConnectAsync(IpAddress, Port);
                _stream = _client.GetStream();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al conectar: {ex.Message}");
                return false;
            }
        }

        public async Task<string> SendCommandAsync(string command)
        {
            if (_stream == null || !_client.Connected)
                throw new InvalidOperationException("No hay conexión activa con la cámara.");

            try
            {
                byte[] commandBytes = Encoding.ASCII.GetBytes(command + "\r");
                await _stream.WriteAsync(commandBytes, 0, commandBytes.Length);

                byte[] buffer = new byte[1024];
                int bytesRead = await _stream.ReadAsync(buffer, 0, buffer.Length);
                return Encoding.ASCII.GetString(buffer, 0, bytesRead);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al enviar comando: {ex.Message}");
                return ex.Message;
            }
        }

        public Task<string> SendTrigger()
        { 
            var resultado = SendCommandAsync("T2");
            return resultado;
        }

        public Task<string> ChangeProgram(int program)
        {
            var resultado = SendCommandAsync($"PW,{program}");
            return resultado;
        }

        public void Disconnect()
        {
            _stream?.Dispose();
            _client?.Close();
            _client = null;
            _stream = null;
        }

        public List<string> Formato(string entrada)
        {
            var partes = entrada.Split(',');
            var resultado = new List<string>();

            if (partes.Length < 4 || partes[0] != "RT")
            {
                resultado.Add(entrada);
                return resultado;
            }

            string estadoGlobal = partes[2];
            resultado.Add($"Área: {estadoGlobal}");

            for (int i = 3; i + 2 < partes.Length; i += 3)
            {
                string herramienta = partes[i];
                string resultadoHerramienta = partes[i + 1];
                string tasa = partes[i + 2].TrimStart('0');

                if (string.IsNullOrEmpty(tasa))
                    tasa = "0";

                int numeroHerramienta = int.TryParse(herramienta, out var h) ? h + 1 : i;
                resultado.Add($"Herramienta {numeroHerramienta}: {resultadoHerramienta} - {tasa}");
            }

            return resultado;
        }


        public void Dispose()
        {
            Disconnect();
        }
    }
}