using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidSerialPort
{
    public class SerialPortConfig
    {
        public string Name { get; private set; }
        public int BaudRate { get; private set; }
        public int DataBits { get; private set; }
        public StopBits StopBits { get; private set; }
        public Parity Parity { get; private set; }
        public bool DtrEnable { get; private set; }
        public bool RtsEnable { get; private set; }

        public SerialPortConfig(
            string name,
            int baudRate,
            int dataBits,
            StopBits stopBits,
            Parity parity,
            bool dtrEnable,
            bool rtsEnable)
        {
            if (String.IsNullOrWhiteSpace(name)) throw new ArgumentNullException("name");

            this.RtsEnable = rtsEnable;
            this.BaudRate = baudRate;
            this.DataBits = dataBits;
            this.StopBits = stopBits;
            this.Parity = parity;
            this.DtrEnable = dtrEnable;
            this.Name = name;
        }

        public override string ToString()
        {
            return String.Format(
                "{0} (Baud: {1}/DataBits: {2}/Parity: {3}/StopBits: {4}/{5})",
                this.Name,
                this.BaudRate,
                this.DataBits,
                this.Parity,
                this.StopBits,
                this.RtsEnable ? "RTS" : "No RTS");
        }
    }

}
