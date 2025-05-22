using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Text;

namespace Scale_Program.Functions
{
    public class SeaLevel : ISeaLevelDevice
    {
        public void Dispose()
        {
            if (serialPort != null)
            {
                if (serialPort.IsOpen)
                    //calling close also disposes the object
                    serialPort.Close();
                serialPort = null;
            }
        }

        #region Constants

        //Timeout in milliseconds
        private const int ReadTimeOut = 100;
        private const int WriteTimeOut = 150;


        //Response header includes <address><function><byte count>
        private const int ResponseHeaderLength = 3;
        private const int ChecksumLength = 2;
        private const int MaxInputs = 16;
        private const int MaxOuputs = 16;


        //Read Discrete Inputs
        //Modbus Request Packet
        // 
        //Description             Value        Length
        //-------------------------------------------
        //Function                0x02           1
        //Start Address High       ?             1
        //Start Address Low        ?             1
        //Quantity High            ?             1
        //Quantity Low             ?             1
        //Checksum                 ?             2
        // 
        //Modbus Response
        // 
        //Description             Value        Length
        //-------------------------------------------
        //Function                0x02           1
        //Byte Count              N              1
        //Input Status            ?              N
        //Checksum                ?              2
        private const byte ReadDiscretInputsCommand = 0x02;


        //Write Singe Coil 
        //Modbus Request Packet
        //       
        //Description          Value       Length
        //----------------------------------------
        //Function             0x05           1
        //Output Address High  ?              1
        //Output Address Low   ?              1
        //Output Value         0x00 or 0xFF   1
        //                     0x00           1
        //Checksum             ?              2
        // 
        //Modbus Response Packet
        // 
        //Description           Value       Length
        //----------------------------------------
        //Function              0x05           1
        //Output Address High    ?             1
        //Output Address Low     ?             1
        //Output Value           0x00 or 0xff  1
        //                       0x00          1
        //Checksum               ?             2
        private const byte WriteSingleCoilCommand = 0x05;


        //Write Multiple Coil 
        //Modbus Request Packet
        //       
        //Description            Value        Length
        //------------------------------------------
        //Function               0x0f           1
        //Starting Address High   ?             1
        //Starting Address Low    ?             1
        //Quantity High           ?             1
        //Quantity Low            ?             1
        //Byte Count              N             1
        //Output Values           ?             N
        //Checksum                ?             2
        // 
        //Modbus Response Packet
        // 
        //Description           Value        Length
        //-----------------------------------------
        //Function              0x0f            1
        //Starting Address High   ?             1
        //Starting Address Low    ?             1
        //Quantity High           ?             1
        //Quantity Low            ?             1
        //Checksum                ?             2
        private const byte WriteMultipleCoilsCommand = 0x0F;

        #endregion

        #region Private Members

        private SerialPort serialPort;
        private readonly StreamWriter streamWriter;

        #endregion

        #region Properties

        /// <summary>
        /// </summary>
        public string PortName => serialPort.PortName;

        /// <summary>
        /// </summary>
        public int BaudRate => serialPort.BaudRate;

        /// <summary>
        /// </summary>
        public byte ModbusAddress { get; }

        public bool IsOpen => serialPort.IsOpen;

        #endregion

        #region Constructor

        public SeaLevel(int address, SerialPort serialPort, Stream logOuput)
        {
            if (serialPort == null) throw new ArgumentNullException("serialPort");

            this.serialPort = serialPort;
            this.serialPort.WriteTimeout = WriteTimeOut;
            this.serialPort.ReadTimeout = ReadTimeOut;

            ModbusAddress = (byte)address;

            //if log required initialize stream writer
            if (logOuput != null) streamWriter = new StreamWriter(logOuput);
        }

        public SeaLevel(int address, SerialPort serialPort)
            : this(address, serialPort, null)
        {
        }

        public SeaLevel(int address, string portName, int baudRate)
            : this(address, new SerialPort(portName, baudRate), null)
        {
        }

        #endregion

        #region Public Methods

        public void Open()
        {
            if (!serialPort.IsOpen) serialPort.Open();
        }

        public void Close()
        {
            if (serialPort.IsOpen) serialPort.Close();
        }

        /// <summary>
        ///     Sets or clears one or more outputs
        /// </summary>
        /// <param name="startIndex"></param>
        /// <param name="value"></param>
        /// <param name="count"></param>
        public void WriteMultipleCoils(int startIndex, int value, int count)
        {
            if (count == 0) throw new ArgumentOutOfRangeException("count", "Must be greater than 0");

            if (startIndex >= MaxOuputs) throw new ArgumentOutOfRangeException("startIndex");

            if (startIndex + count > MaxOuputs) throw new ArgumentOutOfRangeException("count");

            var message = new List<byte>();

            message.Add(ModbusAddress);
            message.Add(WriteMultipleCoilsCommand);
            message.Add(
                (byte)((startIndex >> 8) & 0x00ff)
            );
            message.Add(
                (byte)(startIndex & 0x00ff)
            );

            message.Add(
                (byte)((count >> 8) & 0x00ff)
            );

            message.Add(
                (byte)(count & 0x00ff)
            );

            var numOfBytes = count / 8;
            if (count % 8 > 0) numOfBytes += 1;

            message.Add((byte)numOfBytes);

            var values = BitConverter.GetBytes(value);

            //write in LSB to MSB order
            for (var i = 0; i < numOfBytes; i++) message.Add(values[i]);

            var response = new byte[8];

            lock (serialPort)
            {
                WriteMessage(message.ToArray(), 0, message.Count);

                if (ReadMessage(response, 0, 8, ReadTimeOut) < 0)
                    //time-out elapses prior receiving the response
                    throw new SeaLevelTimeoutException();
            }
        }

        /// <summary>
        ///     Sets or clears a single output
        /// </summary>
        /// <param name="index">zero index based of output</param>
        /// <param name="active">true for set, false to clear the output</param>
        /// <exception cref="SeaLevelTimeoutException">The device did not respond before the time-out period elapses</exception>
        public void SetSingleCoilState(int index, bool active)
        {
            var message = new List<byte>();

            message.Add(ModbusAddress);
            message.Add(WriteSingleCoilCommand);
            message.Add(
                (byte)((index >> 8) & 0x00FF)
            );
            message.Add(
                (byte)(index & 0x00FF)
            );

            if (active)
                message.Add(0xFF);
            else
                message.Add(0x00);

            message.Add(0x00);

            var response = new byte[message.Count + 2];

            //serialize access to this part
            lock (serialPort)
            {
                WriteMessage(message.ToArray(), 0, message.Count);

                if (ReadMessage(response, 0, response.Length, ReadTimeOut) < 0)
                    //time-out elapses prior receiving the response
                    throw new SeaLevelTimeoutException();
            }
        }

        /// <summary>
        ///     Reads one or more discrete inputs
        /// </summary>
        /// <param name="startIndex">zero based index of the input</param>
        /// <param name="count">The number of inputs to read</param>
        /// <returns>The state of <see cref="count" /> inputs  starting from <see cref="startIndex" /></returns>
        /// <exception cref="ArgumentOutOfRangeException">
        ///     Either StartIndex is greater or equal to number of inputs in the device or the combination of both parameters
        ///     exceeds
        ///     the maximun number of inputs
        /// </exception>
        /// <exception cref="SeaLevelTimeoutException">The device did not respond before the time-out period elapses</exception>
        public uint ReadDiscreteInputs(int startIndex, int count)
        {
            //perfom input parameters validation
            if (startIndex < 0 || startIndex >= MaxInputs) throw new ArgumentOutOfRangeException("startIndex");

            if (startIndex + count > MaxInputs) throw new ArgumentOutOfRangeException("count");

            var message = new List<byte>();

            message.Add(ModbusAddress);
            message.Add(ReadDiscretInputsCommand);

            message.Add(
                (byte)((startIndex >> 8) & 0x00ff)
            );
            message.Add(
                (byte)(startIndex & 0x00ff)
            );
            message.Add(
                (byte)((count >> 8) & 0x00ff)
            );
            message.Add(
                (byte)(count & 0x00ff)
            );

            int responseLength;

            //calculate the number of bytes required to fit the result
            var resultByteCount = count / 8;
            if (count % 8 != 0) resultByteCount += 1;

            responseLength = ResponseHeaderLength + resultByteCount;

            //add 2 bytes to include the length of the checksum
            responseLength += ChecksumLength;

            var response = new byte[responseLength];

            lock (serialPort)
            {
                WriteMessage(message.ToArray(), 0, message.Count);

                if (ReadMessage(response, 0, responseLength, ReadTimeOut) < 0)
                    //time-out elapses prior receiving the response
                    throw new SeaLevelTimeoutException();
            }

            uint result = 0x0000;
            int byteCount = response[2];
            for (var j = 0; j < byteCount; j++)
            {
                uint value = response[j + 3];

                result |= value << (j * 8);
            }

            return result;
        }

        public override string ToString()
        {
            return string.Format("Port:{0}, BaudRate:{1}, Address:{2}",
                PortName,
                BaudRate,
                ModbusAddress
            );
        }

        #endregion

        #region Private Methods

        private int CalculateChecksum(byte[] data, int offset, int count)
        {
            if (data == null)
                throw new InvalidOperationException("CalculateChecksum: El campo de datos no puede ser nulo.");

            if (data.Length < offset + count)
                throw new InvalidOperationException(
                    string.Format("CalculateChecksum: Indice fuera de rango ({0} < {1})", data.Length, offset + count)
                );

            var crc = 0xffff;
            var carryFlag = 0;

            for (int i = 0, j = offset; i < count; i++, j++)
            {
                crc = crc ^ data[j];

                for (var k = 0; k < 8; k++)
                {
                    carryFlag = crc & 0x01;
                    crc = crc >> 1;
                    if (carryFlag == 0x01)
                        crc = crc ^ 0xA001;
                }
            }

            return crc & 0xffff;
        }

        private void WriteMessage(byte[] data, int offset, int count)
        {
            //create local buffer of data length + 2 bytes for checksum
            var message = new byte[count + 2];

            var checksum = CalculateChecksum(data, 0, count);

            //copy message to local array
            Array.Copy(data, message, count);

            //append checksum to the message
            message[count] = (byte)(checksum & 0xff);
            message[count + 1] = (byte)(checksum >> 8);

            try
            {
                serialPort.DiscardInBuffer();

                serialPort.Write(message, offset, message.Length);

                if (streamWriter != null)
                {
                    var sb = new StringBuilder();

                    sb.Append(" -> ");

                    for (var i = 0; i < count; i++)
                    {
                        if (i > 0) sb.Append(" ");

                        sb.AppendFormat("{0:X2}", data[offset + i]);
                    }

                    LogData(sb.ToString());
                }
            }
            catch (TimeoutException)
            {
                throw new SeaLevelTimeoutException();
            }
        }

        private int ReadMessage(byte[] data, int offset, int count, int timeOut)
        {
            var isDone = false;
            var result = 0;
            var bufferIndex = offset;

            var startTime = DateTime.Now;
            do
            {
                var bytesReceived = serialPort.BytesToRead;

                if (bytesReceived > 0)
                    for (var i = 0; i < bytesReceived; i++)
                    {
                        //abort if buffer is full 
                        if (bufferIndex == data.Length)
                        {
                            isDone = true;
                            result = 1;
                            break;
                        }

                        data[bufferIndex++] = (byte)serialPort.ReadByte();

                        //if read up to count bytes exit loop
                        if (bufferIndex - offset == count)
                        {
                            isDone = true;
                            break;
                        }
                    }

                //check if timeout elapsed prior completion
                if (isDone == false && (DateTime.Now - startTime).TotalMilliseconds >= timeOut)
                {
                    result = -1;
                    isDone = true;
                }
            } while (isDone == false);

            if (streamWriter != null)
            {
                var sb = new StringBuilder();

                sb.Append(" <- ");

                if (isDone == false)
                    sb.AppendLine(string.Format(" TIMEOUT. {0} of {1} received", bufferIndex - offset, count));
                else
                    for (var i = offset; i < bufferIndex; i++)
                    {
                        if (i > 0) sb.Append(" ");

                        sb.AppendFormat("{0:X2}", data[i]);
                    }

                LogData(sb.ToString());
            }

            return result;
        }

        private void LogData(string message)
        {
            var now = DateTime.Now;
            try
            {
                streamWriter.Write("[");
                streamWriter.Write(now.ToString("HH:mm:ss.fff"));
                streamWriter.Write("] ");
                streamWriter.WriteLine(message);
                streamWriter.Flush();
            }
            catch
            {
                //ignore exception, writing to a log is not critic in this case
            }
        }

        #endregion
    }
}