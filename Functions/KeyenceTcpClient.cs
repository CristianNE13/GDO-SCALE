using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Scale_Program.Functions
{
    public class KeyenceTcpClient : IDisposable
    {
        private TcpClient _client;
        private NetworkStream _stream;
        private readonly int _timeout;
        private CancellationTokenSource _monitorTokenSource;
        private readonly int _monitorInterval = 8000;
        private readonly int _reconnectDelay = 3000; 
        public event Action<int> OnCameraStatusChanged;

        public string IpAddress { get; }
        public int Port { get; }

        public KeyenceTcpClient(string ipAddress, int port = 8500, int timeout = 5000)
        {
            IpAddress = ipAddress;
            Port = port;
            _timeout = timeout;
        }

        public async Task<bool> ConnectAsync()
        {
            if (_client != null && _client.Connected)
                return true;

            try
            {
                _client = new TcpClient();
                var connectTask = _client.ConnectAsync(IpAddress, Port);
                if (await Task.WhenAny(connectTask, Task.Delay(_timeout)) == connectTask)
                {
                    _stream = _client.GetStream();
                    return true;
                }
                else
                {
                    throw new TimeoutException("Tiempo de conexión excedido.");
                }
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

                using (var cts = new CancellationTokenSource(_timeout))
                {
                    byte[] buffer = new byte[1024];
                    int bytesRead = await _stream.ReadAsync(buffer, 0, buffer.Length, cts.Token);
                    return Encoding.ASCII.GetString(buffer, 0, bytesRead);
                }
            }
            catch (OperationCanceledException)
            {
                return "Tiempo de respuesta excedido.";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al enviar comando: {ex.Message}");
                return ex.Message;
            }
        }

        public Task<string> SendTrigger()
        { 
            return SendCommandAsync("T2");
        }

        public Task<string> ChangeProgram(int? program)
        {
            if (program == null)
                return Task.FromResult("El programa no puede ser nulo.");

            return SendCommandAsync($"PW,{program}");
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
            StopMonitoring();

            try
            {
                _stream?.Close();
                _stream?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al cerrar el stream: {ex.Message}");
            }

            try
            {
                _client?.Close();
                _client = null;
                _stream = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al cerrar la conexión: {ex.Message}");
            }
        }

        public async Task<int?> CheckActiveProgram()
        {
            try
            {
                string response = await SendCommandAsync("PR");
                if (string.IsNullOrWhiteSpace(response))
                    return null;

                var partes = response.Split(',');
                if (partes.Length < 2)
                    return null;

                if (int.TryParse(partes[1].Trim(), out int program))
                    return program;

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al verificar el programa activo: {ex.Message}");
                return null;
            }
        }

        public void StartMonitoring()
        {
            if (_monitorTokenSource != null && !_monitorTokenSource.IsCancellationRequested)
            {
                Console.WriteLine("Monitoreo ya está activo.");
                return;
            }

            _monitorTokenSource = new CancellationTokenSource();

            if (_client != null && _client.Connected)
                UpdateCameraStatus(2); 
            else
                UpdateCameraStatus(0);

            Task.Run(async () => await MonitorConnection(_monitorTokenSource.Token));
        }



        public void StopMonitoring()
        {
            if (_monitorTokenSource != null && !_monitorTokenSource.IsCancellationRequested)
                _monitorTokenSource.Cancel();
        }

        private async Task MonitorConnection(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    int? activeProgram = await CheckActiveProgram();

                    if (activeProgram == null)
                    {
                        Console.WriteLine("Conexión perdida. Intentando reconectar...");
                        UpdateCameraStatus(1);

                        await Task.Delay(_reconnectDelay, token);

                        bool reconnected = await ConnectAsync();
                        if (!reconnected)
                        {
                            Console.WriteLine("No se pudo reconectar a la cámara.");
                            UpdateCameraStatus(0);
                        }
                        else
                        {
                            Console.WriteLine("Reconexión exitosa.");
                            UpdateCameraStatus(2);
                        }
                    }
                    else
                    {
                        UpdateCameraStatus(2);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error en el monitoreo de conexión: {ex.Message}");
                    UpdateCameraStatus(0);
                }

                await Task.Delay(_monitorInterval, token);
            }
        }


        private void UpdateCameraStatus(int status)
        {
            OnCameraStatusChanged?.Invoke(status);
        }

    }
}