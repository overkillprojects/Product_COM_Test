using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Product_COM_Test
{
    public class ComPortProperties
    {
        public ComPortProperties(string ComPortName, string VID, string PID, string FriendlyName)
        {
            this.ComPortName = ComPortName;
            this.VID = VID;
            this.PID = PID;
            this.FriendlyName = FriendlyName;
        }

        public string ComPortName { get; set; }
        public string VID { get; set; }
        public string PID { get; set; }
        public string FriendlyName { get; set; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder(ComPortName);

            if(!string.IsNullOrWhiteSpace(FriendlyName))
            {
                sb.Append($": {FriendlyName}");
            }

            if (!(string.IsNullOrWhiteSpace(VID) && string.IsNullOrWhiteSpace(PID)))
            {
                sb.Append($" ({VID}/{PID})");
            }

            return sb.ToString();
        }
    }

    public class ComPortManager
    {
        // Same as GetPortProperties(), but only returns the com port name
        public static List<string> GetPortNames() => GetPortProperties().Select((x) => x.ComPortName).ToList();

        // Same as GetPortProperties(String VID, String PID), but only returns the com port name
        public static List<string> GetPortNames(String VID, String PID) => GetPortProperties(VID, PID).Select((x) => x.ComPortName).ToList();

        // Same as GetPortProperties(String VID, String PID), but gets all connected COM ports (ignores VID/PID)
        // You can use this function to display the name of the device (like "FTDI Serial Port") rather than just "COM1"
        public static List<ComPortProperties> GetPortProperties() => ComPortNames("^VID_([0-9a-fA-F]{4}).PID_([0-9a-fA-F]{4})");

        /// <summary>
        /// Gets the properties of the COM ports of all devices with the given VID and PID
        /// From https://stackoverflow.com/questions/10350340/identify-com-port-using-vid-and-pid-for-usb-device-attached-to-x64
        /// </summary>
        /// <param name="VID">The Vendor ID, eg "0403" for FTDI</param>
        /// <param name="PID">The Product ID, eg "6001" for the FT232 Serial UART</param>
        public static List<ComPortProperties> GetPortProperties(String VID, String PID) => ComPortNames($"^VID_({VID}).PID_({PID})");

        // Find com ports which match a certain VID/PID by pattern
        // Note: Regex pattern must have two groups, one for the VID and one for the PID!
        // Note: This code is only needed if you want to get the descriptive name of the com port, otherwise
        //       you can just use the GetConnectedComPortNames() function.
        private static List<ComPortProperties> ComPortNames(String pattern)
        {
            Regex _rx = new Regex(pattern, RegexOptions.IgnoreCase);
            List<ComPortProperties> comports = new List<ComPortProperties>();

            RegistryKey rk1 = Registry.LocalMachine;
            RegistryKey rk2 = rk1.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum");

            foreach (String s3 in rk2.GetSubKeyNames())
            {
                //rk3 is something like "USB" or "STORAGE" .... various other names as well
                RegistryKey rk3 = rk2.OpenSubKey(s3);
                foreach (String s in rk3.GetSubKeyNames())
                {
                    Match regex_match_result = _rx.Match(s);
                    if (regex_match_result.Success)
                    {
                        //rk4 is something like VID_2341&PID_0042
                        RegistryKey rk4 = rk3.OpenSubKey(s);
                        foreach (String s2 in rk4.GetSubKeyNames())
                        {
                            try
                            {
                                // RK5 is something like "5573731323135120B022"
                                RegistryKey rk5 = rk4.OpenSubKey(s2);

                                //Get the friendly name
                                string friendlyName = String.Empty;
                                try {
                                    // if key doesn't exist (ret null), default to String.Empty
                                    friendlyName = (string)rk5.GetValue("FriendlyName") ?? String.Empty; 
                                } catch (Exception) { }

                                // The "Device Parameters" field stores the COM port name
                                RegistryKey rk6 = rk5.OpenSubKey("Device Parameters");
                                string portName = (string)rk6.GetValue("PortName");

                                // Registry lists 'ghost' com ports which are not connected to the computer (they were connected in the past).
                                // Check that the com port is not a ghost.
                                if (!String.IsNullOrEmpty(portName) && GetConnectedComPortNames().Contains(portName))
                                {
                                    comports.Add(
                                        new ComPortProperties(
                                            ComPortName: (string)rk6.GetValue("PortName"),
                                            VID: regex_match_result.Groups[1].ToString(),
                                            PID: regex_match_result.Groups[2].ToString(),
                                            FriendlyName: friendlyName
                                        )
                                    );
                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"Com Port Manager: Got Malformed registry key at LocalMachine\\SYSTEM\\CurrentControlSet\\Enum\\{s3}\\{s}\\{s2}\n{e}");
                            }
                        }
                    }
                }
            }
            return comports;
        }

        // Find any connected com ports, regardless of manufacturer
        private static List<string> GetConnectedComPortNames()
        {
            string[] names = SerialPort.GetPortNames();

            // Fix bug in windows when querying connected com port names
            // Sometimes there will be extra characters after the com port name. 
            // In my case, I get "COM06\0" then garbage characters, so I look for a null terminator
            // This may not always fix the issue.
            // See https://stackoverflow.com/questions/32040209/serialport-getportnames-returns-incorrect-port-names/32227459
            List<string> fixed_names = new List<string>();
            foreach (string name in names)
            {
                int i;
                for (i = 0; i < name.Length; i++)
                {
                    int charVal = name[i];
                    if (charVal == 0)
                    {
                        break;
                    }
                }

                fixed_names.Add(name.Substring(0, i));
            }

            return fixed_names;
        }
    }
}