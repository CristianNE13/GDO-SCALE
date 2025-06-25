using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Threading;

namespace Scale_Program.Functions
{
    public sealed class BasculaFuncGFC : IBasculaFunc
    {
        //The maximum size in characters to keep in the log
        private const int MaxLogEntries = 100;
        private const int EntriesToKeep = 10;

        private static readonly List<LogEntry> log = new List<LogEntry>(MaxLogEntries);
        private readonly int EOLCount = 3;

        private readonly string Grams = "gr";

        private readonly StringBuilder sb = new StringBuilder();
        private readonly ManualResetEvent waitHandle;

        public Dispatcher dispatcher;


        private double lastWeight;
        private int newLineCount;
        public Regex regex;

        public SerialPort sPort;

        public BasculaFuncGFC()
        {
            waitHandle = new ManualResetEvent(true);
        }

        public string Puerto => sPort == null ? string.Empty : sPort.PortName;

        public event EventHandler<BasculaEventArgs> OnDataReady;

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
                    //build regex validator
                    regex = new Regex(
                        @"^(?<status>(ST|US)),(GS|NT),\s+(?<weight>(-)?([0-9]+\.[0-9]+))\s+(?<units>(kg|gr|lb))"
                    );
                    sPort.DataReceived += OnDataReceived;

                    waitHandle.Set();

                    sPort.Open();
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
                sPort.Write("C" + "\r\n");
                Thread.Sleep(100);
                sPort.Write("T" + "\r\n");

                LogMessage("Z");
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
                Console.WriteLine("OnDataReceived " + ex.Message);
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

                    if (rxByte == '\n')
                    {
                        newLineCount++;
                        if (newLineCount == EOLCount)
                        {
                            //if we have to work with
                            if (sb.Length > 0)
                            {
                                //process incoming data
                                ProcessBuffer();

                                //clear buffer
                                sb.Clear();
                            }
                        }
                        else
                        {
                            //only include one eol
                            sb.AppendLine();
                        }
                    }
                    else if (rxByte == '\r')
                    {
                        //skip it
                    }
                    else
                    {
                        //restart count
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

                if (string.Compare(units, Grams, true) == 0)
                    //convert to kilos
                    weight = weight / 1000;


                //convertir kilos a libras//


                //End convertion

                var isStable = status.ToUpper() == "ST";

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

        private static void LogMessage(string message)
        {
            var now = DateTime.Now;

            if (log.Count == MaxLogEntries)
            {
                var toRemove = log.Count - EntriesToKeep;

                for (var i = 0; i < toRemove; i++)
                    //remove older entries
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