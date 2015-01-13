using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SolidSerialPort;

namespace SolidSerialPort.Tests
{
    [TestClass]
    public class SerialPortFixerTests
    {
        [TestMethod]
        public void InvalidPortFailure()
        {
            try
            {
                SerialPortFixer.Execute("");
                Assert.Fail();
            }
            catch (Exception)
            {
            }
        }

        [TestMethod]
        public void ValidPortSuccess()
        {
            var ports = System.IO.Ports.SerialPort.GetPortNames();
            if (ports.Length > 0)
            {
                try
                {
                    SerialPortFixer.Execute(ports[0]);
                }
                catch (Exception)
                {
                    Assert.Fail();
                }
            }
            else
            {
                Assert.Inconclusive();
            }
        }
    }
}
