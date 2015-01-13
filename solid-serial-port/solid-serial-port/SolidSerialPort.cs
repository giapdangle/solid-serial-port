using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SolidSerialPort
{
    public interface ISolidSerialPort : IDisposable
    {
        string PortName { get; }
        string ReadLine();
        void WriteLine(string text);
        bool IsOpen { get; }
        int ReadByte();
        void Write(byte[] buffer, int offset, int count);
        int ReadTimeout { get; set; }
        int WriteTimeout { get; set; }
    }

	// Wrapper around SerialPort
    public class SolidSerialPort : ISolidSerialPort
    {
        readonly SerialPort _port;
        readonly Stream _internalSerialStream;

        public SolidSerialPort(SerialPortConfig portConfig)
        {
            if (portConfig == null) throw new ArgumentNullException("portConfig");

            // http://zachsaw.blogspot.com/2010/07/net-serialport-woes.html
            SerialPortFixer.Execute(portConfig.Name);

            var port = new SerialPort(
                portConfig.Name,
                portConfig.BaudRate,
                portConfig.Parity,
                portConfig.DataBits,
                portConfig.StopBits)
            {
                RtsEnable = portConfig.RtsEnable,
                DtrEnable = portConfig.DtrEnable,
                ReadTimeout = 5000,
                WriteTimeout = 5000
            };
            port.Open();

            try
            {
                this._internalSerialStream = port.BaseStream;
                this._port = port;
                this._port.DiscardInBuffer();
                this._port.DiscardOutBuffer();
            }
            catch (Exception ex)
            {
                Stream internalStream = this._internalSerialStream;

                if (internalStream == null)
                {
                    FieldInfo field = typeof(SerialPort).GetField(
                        "internalSerialStream",
                        BindingFlags.Instance | BindingFlags.NonPublic);

                    // This will happen if the SerialPort class is changed
                    // in future versions of the .NET Framework
                    if (field == null)
                    {
                        /*
                        this.Log.WarnFormat(
                            "An exception occured while creating the serial port adaptor, "
                            + "the internal stream reference was not acquired and we were unable "
                            + "to get it using reflection. The serial port may not be accessible "
                            + "any further until the serial port object finalizer has been run: {0}",
                            ex);
                        */
                        throw;
                    }

                    internalStream = (Stream)field.GetValue(port);
                }

                /*
                this.Log.DebugFormat(
                    "An error occurred while constructing the serial port adaptor: {0}", ex);
                */
                SafeDisconnect(port, internalStream);
                throw;
            }
        }
        public bool IsOpen { get { return _port.IsOpen; } }
        public int ReadTimeout { get; set; }
        public int WriteTimeout { get; set; }

        public int ReadByte()
        {
            return _port.ReadByte();
        }

        public void Write(byte[] buffer, int offset, int count)
        {
            _port.Write(buffer, offset, count);
        }

        public string PortName
        {
            get { return this._port.PortName; }
        }

        public string ReadLine()
        {
            return this._port.ReadTo(Environment.NewLine);
        }

        public void WriteLine(string text)
        {
            this._port.Write(text);
            this._port.Write("\r");
        }

        public void Dispose()
        {
            this.Dispose(true);
        }

        protected void Dispose(bool disposing)
        {
            SafeDisconnect(this._port, this._internalSerialStream);

            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }

        /// <summary>
        /// Safely closes a serial port and its internal stream even if
        /// a USB serial interface was physically removed from the system
        /// in a reliable manner.
        /// </summary>
        /// <param name="port"></param>
        /// <param name="internalSerialStream"></param>
        /// <remarks>
        /// The <see cref="SerialPort"/> class has 3 different problems in disposal
        /// in case of a USB serial device that is physically removed:
        /// 
        /// 1. The eventLoopRunner is asked to stop and <see cref="SerialPort.IsOpen"/> 
        /// returns false. Upon disposal this property is checked and closing 
        /// the internal serial stream is skipped, thus keeping the original 
        /// handle open indefinitely (until the finalizer runs which leads to the next problem)
        /// 
        /// The solution for this one is to manually close the internal serial stream.
        /// We can get its reference by <see cref="SerialPort.BaseStream" />
        /// before the exception has happened or by reflection and getting the 
        /// "internalSerialStream" field.
        /// 
        /// 2. Closing the internal serial stream throws an exception and closes 
        /// the internal handle without waiting for its eventLoopRunner thread to finish, 
        /// causing an uncatchable ObjectDisposedException from it later on when the finalizer 
        /// runs (which oddly avoids throwing the exception but still fails to wait for 
        /// the eventLoopRunner).
        /// 
        /// The solution is to manually ask the event loop runner thread to shutdown
        /// (via reflection) and waiting for it before closing the internal serial stream.
        /// 
        /// 3. Since Dispose throws exceptions, the finalizer is not suppressed.
        /// 
        /// The solution is to suppress their finalizers at the beginning.
        /// </remarks>
        static void SafeDisconnect(SerialPort port, Stream internalSerialStream)
        {
            GC.SuppressFinalize(port);
            GC.SuppressFinalize(internalSerialStream);

            ShutdownEventLoopHandler(internalSerialStream);

            try
            {
                /*
                s_Log.DebugFormat("Disposing internal serial stream");
                 */
                internalSerialStream.Close();
            }
            catch (Exception ex)
            {
                /*
                s_Log.DebugFormat(
                    "Exception in serial stream shutdown of port {0}: {1}", port.PortName, ex);
                 */
            }

            try
            {
                /*
                s_Log.DebugFormat("Disposing serial port");
                 */
                port.Close();
            }
            catch (Exception ex)
            {
                /*
                s_Log.DebugFormat("Exception in port {0} shutdown: {1}", port.PortName, ex);
                 */
            }
        }

        static void ShutdownEventLoopHandler(Stream internalSerialStream)
        {
            try
            {
                /*
                s_Log.DebugFormat("Working around .NET SerialPort class Dispose bug");
                 */

                FieldInfo eventRunnerField = internalSerialStream.GetType()
                    .GetField("eventRunner", BindingFlags.NonPublic | BindingFlags.Instance);

                if (eventRunnerField == null)
                {
                    /*
                    s_Log.WarnFormat(
                        "Unable to find EventLoopRunner field. "
                        + "SerialPort workaround failure. Application may crash after "
                        + "disposing SerialPort unless .NET 1.1 unhandled exception "
                        + "policy is enabled from the application's config file.");
                     */
                }
                else
                {
                    object eventRunner = eventRunnerField.GetValue(internalSerialStream);
                    Type eventRunnerType = eventRunner.GetType();

                    FieldInfo endEventLoopFieldInfo = eventRunnerType.GetField(
                        "endEventLoop", BindingFlags.Instance | BindingFlags.NonPublic);

                    FieldInfo eventLoopEndedSignalFieldInfo = eventRunnerType.GetField(
                        "eventLoopEndedSignal", BindingFlags.Instance | BindingFlags.NonPublic);

                    FieldInfo waitCommEventWaitHandleFieldInfo = eventRunnerType.GetField(
                        "waitCommEventWaitHandle", BindingFlags.Instance | BindingFlags.NonPublic);

                    if (endEventLoopFieldInfo == null
                        || eventLoopEndedSignalFieldInfo == null
                        || waitCommEventWaitHandleFieldInfo == null)
                    {
                        /*
                        s_Log.WarnFormat(
                            "Unable to find the EventLoopRunner internal wait handle or loop signal fields. "
                            + "SerialPort workaround failure. Application may crash after "
                            + "disposing SerialPort unless .NET 1.1 unhandled exception "
                            + "policy is enabled from the application's config file.");
                         */
                    }
                    else
                    {
                        /*
                        s_Log.DebugFormat(
                            "Waiting for the SerialPort internal EventLoopRunner thread to finish...");
                         */
                        
                        var eventLoopEndedWaitHandle =
                            (WaitHandle)eventLoopEndedSignalFieldInfo.GetValue(eventRunner);
                        var waitCommEventWaitHandle =
                            (ManualResetEvent)waitCommEventWaitHandleFieldInfo.GetValue(eventRunner);

                        endEventLoopFieldInfo.SetValue(eventRunner, true);

                        // Sometimes the event loop handler resets the wait handle
                        // before exiting the loop and hangs (in case of USB disconnect)
                        // In case it takes too long, brute-force it out of its wait by 
                        // setting the handle again.
                        do
                        {
                            waitCommEventWaitHandle.Set();
                        } while (!eventLoopEndedWaitHandle.WaitOne(2000));

                        /*
                        s_Log.DebugFormat("Wait completed. Now it is safe to continue disposal.");
                         */
                    }
                }
            }
            catch (Exception ex)
            {
                /*
                s_Log.ErrorFormat(
                    "SerialPort workaround failure. Application may crash after "
                    + "disposing SerialPort unless .NET 1.1 unhandled exception "
                    + "policy is enabled from the application's config file: {0}",
                    ex);
                 */
            }
        }
    }
}
