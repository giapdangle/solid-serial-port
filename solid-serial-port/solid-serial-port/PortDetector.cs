
using System;
using System.Threading;
using System.Linq;
using System.IO.Ports;
using System.Collections.Generic;
using System.Management;

namespace SolidSerialPort
{
    public static class PortDetector
    {
        public enum VID_PID_Selector
        {
            TI_CC2540
        }
        private static readonly Dictionary<VID_PID_Selector, string> VID_PID_SelectorStrings = new Dictionary<VID_PID_Selector,string>()
        {
            { VID_PID_Selector.TI_CC2540, @"USB\VID_0451&PID_16AA" }
        };

		private static bool canUseWMI = true;

        public static List<string> DetectCandidatePorts(VID_PID_Selector portSelector)
		{
			List<string> portNames = new List<string>();

            // First time we will try to use wmi
			if (canUseWMI)
			{
                try 
	            {
				    var instances = new ManagementClass("Win32_SerialPort").GetInstances();
				    foreach (ManagementObject port in instances)
				    {
                        if (port.Properties["PNPDeviceID"].Value.ToString().Contains(VID_PID_SelectorStrings[portSelector]))
						    portNames.Add(port.Properties["DeviceID"].Value.ToString());// deviceid
				    }
	            }
			    catch (System.Runtime.InteropServices.COMException ex)
			    {
				    return null; // not sure what happened 
			    }
			    catch (System.Management.ManagementException ex)
			    {
				    if (ex.ErrorCode == ManagementStatus.ProviderLoadFailure)
				    {
                        // This has been observed in gatekeeper before, retry
					    canUseWMI = false;
                        return DetectCandidatePorts(portSelector);
				    }

                    // Unknown case
				    return null;
			    }
			    catch (Exception ex)
			    {
				    throw ex;
			    }
			}
			else
			{
                // If we can't use wmi, the user must choose
				portNames = SerialPort.GetPortNames().ToList();
			}

			return portNames;
		}
	}
}
