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
    }
}
