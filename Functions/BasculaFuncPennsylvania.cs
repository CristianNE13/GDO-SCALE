using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Threading;
using log4net;

namespace Scale_Program.Functions
{
    public sealed class BasculaFuncPennsylvania : IBasculaFunc
    {
        public event EventHandler<BasculaEventArgs> OnDataReady;

        private const int MaxLogEntries = 100;
        private const int EntriesToKeep = 10;
        private static readonly ILog Logger = LogManager.GetLogger("BasculaFunc");

        private static readonly List<LogEntry> log = new List<LogEntry>(MaxLogEntries);
        private readonly int EOLCount = 1;
        private readonly string grams = "gr";
        private readonly string kilograms = "kg";
        private readonly string pounds = "lb";

        private readonly StringBuilder sb = new StringBuilder();
        private readonly ManualResetEvent waitHandle;

        public Dispatcher dispatcher;

        private double lastWeight;
        private int newLineCount;
        public Regex regex;

        public SerialPort sPort;

        public BasculaFuncPennsylvania()
        {
            waitHandle = new ManualResetEvent(true);
        }

        public string Puerto => sPort == null ? string.Empty : sPort.PortName;

        public static IList<LogEntry> LogEntries => log.AsReadOnly();

        public void AsignarControles(Dispatcher dispatcher)
        {
            this.dispatcher = dispatcher;
        }

        public SerialPort GetPuerto()
        {
            return sPort;
        }

        public void OpenPort()
        {
            if (sPort != null)
                if (!sPort.IsOpen)
                {
                    regex = new Regex(
                        @"^(?<status>(@|B|D|F|A|E))\s*(Gross)\s*(?<weight>(-)?([0-9]+\.[0-9]+))\s*(?<units>(lb|kg|gr))$"
                    );

                    sPort.DataReceived += OnDataReceived;

                    waitHandle.Set();


                    sPort.Open();
                    sPort.DiscardInBuffer();
                }
        }

        public void ClosePort()
        {
            if (sPort != null)
                try
                {
                    if (sPort.IsOpen)
                    {
                        waitHandle.WaitOne();
                        sPort.Close();
                    }
                }
                finally
                {
                    sPort.Dispose();
                    sPort = null;
                }
        }

        private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                waitHandle.Set();

                var port = sender as SerialPort;

                HandleIncomingData(port);
            }
            catch (Exception ex)
            {
                Logger.Error("OnDataReceived", ex);
            }
            finally
            {
                waitHandle.Set();
            }
        }

        private void HandleIncomingData(SerialPort port)
        {
            var availableData = 0;
            int count;
            int rxByte;

            if (port != null && port.IsOpen)
            {
                availableData = port.BytesToRead;

                count = 0;

                while (count < availableData)
                {
                    rxByte = port.ReadByte();
                    count++;

                    if (rxByte == 10)
                    {
                        newLineCount++;

                        if (newLineCount == EOLCount)
                        {
                            if (sb.Length > 0)
                            {
                                ProcessBuffer();
                                sb.Clear();
                            }
                        }
                        else
                        {
                            sb.AppendLine();
                        }
                    }
                    else if (rxByte == 13)
                    {
                    }
                    else
                    {
                        newLineCount = 0;
                        sb.Append(Convert.ToChar(rxByte));
                    }
                }
            }
        }

        private void ProcessBuffer()
        {
            var match = regex.Match(sb.ToString());
            if (match.Success)
            {
                var weight = Convert.ToDouble(match.Groups["weight"].Value);
                var units = match.Groups["units"].Value;
                var status = match.Groups["status"].Value;

                if (string.Compare(units, pounds, true) == 0)
                    weight *= 0.453592;

                else if (string.Compare(units, grams, true) == 0)
                    weight *= 0.001;

                else if (string.Compare(units, kilograms, true) == 0)
                {
                    // Already in kilograms
                }

                var isStable = status.ToUpper() == "@" || status.ToUpper() == "A" || status.ToUpper() == "D" ||
                               status.ToUpper() == "E";

                {
                    if (dispatcher != null)
                    {
                        var e = new BasculaEventArgs();
                        e.Value = weight;
                        e.IsStable = isStable;

                        dispatcher.BeginInvoke((Action)(() => { RaiseEvent(e); }
                            )
                        );
                    }

                    lastWeight = weight;
                }
            }

            LogMessage(sb.ToString());
        }


        private void RaiseEvent(BasculaEventArgs e)
        {
            OnDataReady?.Invoke(this, e);
        }


        public void AsignarPuertoBascula(SerialPort newSerialPort)
        {
            sPort = newSerialPort;
            sPort.DataReceived += OnDataReceived;
        }

        public void EnviarComandoABascula(string comando)
        {
            if (sPort.IsOpen)
            {
                sPort.Write(comando + "\r\n");

                LogMessage(comando);
            }
        }

        public void EnviarZero()
        {
            if (sPort.IsOpen)
            {
                sPort.Write("ZRO" + "\r\n");

                LogMessage("ZRO");
            }
        }

        private static void LogMessage(string message)
        {
            var now = DateTime.Now;

            if (log.Count == MaxLogEntries)
            {
                var toRemove = log.Count - EntriesToKeep;

                for (var i = 0; i < toRemove; i++)

                    log.RemoveAt(0);
            }

            log.Add(
                new LogEntry
                {
                    EntryDate = now,
                    Entry = message
                }
            );
        }
    }
}