using System;
using System.Diagnostics;
using System.Threading;
using System.IO.Ports;
using System.Management;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace ArduinoToFoobar
{
    /*
     * This is a hidden console application used to read serial data from an Arduino nano clone.
     * 
     * The data received will be treated as 'commands' to control music playback within a 
     * program called foobar2000. 
     * 
     * Every 30 seconds this application checks if the device is connected which is 
     * useful if it has been unplugged and plugged in again, as the connection will be reestablished.
     * 
     * SETUP:
     * 1. Open Device Manager
     * 2. Double click the 'USB-SERIAL CH340' listed under 'Ports'
     * 3. Go to the 'Details' tab and copy the 'Class GUID' property value
     * 4. Replace 'ExpectedClassGuid' string with this value
     */

    class Program
    {
        const string ExpectedClassGuid = "{4d36e978-e325-11ce-bfc1-08002be10318}";
        const string ExpectedDisplayName = "USB-SERIAL CH340 (COM";
        const string foobarExecutable = "\"C:\\Program Files (x86)\\foobar2000\\foobar2000.exe\" ";

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static SerialPort device = null;
        static string COM = string.Empty;

        static ProcessStartInfo processStartInfo;
        static Process process;

        static void Main()
        {
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, 0); // Passing a Zero to hide this console window

            var timer = new Timer(ConnectionWatchdog, null, 0, 30000);

            // Keep program alive
            Console.ReadLine();
        }

        /// <summary>
        /// Watchdog method called during each timer interval to provide a
        /// constant connection to the device.
        /// </summary>
        /// <param name="state"></param>
        private static void ConnectionWatchdog(object state)
        {
            if (device == null)
            {
                // Find correct device and connect to it
                if (FindSerialDeviceByName())
                {
                    ConnectToSerialDevice();
                }
            }
            else if (!device.IsOpen) // if lost connection to device
            {
                DeviceCleanup();
            }
        }

        /// <summary>
        /// Enumerates plug and play hardware connected to computer.
        /// Identifies specific device to use based on an expected ClassGUID and Device Name.
        /// </summary>
        /// <returns>bool indicating whether expected device was found</returns>
        private static bool FindSerialDeviceByName()
        {
            ManagementObjectSearcher search = new ManagementObjectSearcher("root\\CIMV2",
                "SELECT * FROM Win32_PnPEntity WHERE ClassGuid = \"" + ExpectedClassGuid + "\"");

            foreach (ManagementObject objectFound in search.Get())
            {
                string displayName = objectFound["Caption"].ToString();

                if (displayName.Contains(ExpectedDisplayName))
                {
                    Regex rx = new Regex(@"(COM\d+)(?!\()");
                    COM = rx.Match(displayName).Value;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Attempts to open a serial port connection to enable communication with the device.
        /// </summary>
        private static void ConnectToSerialDevice()
        {
            try
            {
                device = new SerialPort(COM, 9600, Parity.None, 8, StopBits.One);
                device.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
                device.Open();
                Debug.WriteLine("Successful connection to serial device " + COM);
            }
            catch (Exception)
            {
                DeviceCleanup();
                Debug.WriteLine("Failed to connect to serial device " + COM);
            }
        }
        /// <summary>
        /// Event that is fired when serial data is received from the device.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                SerialPort port = sender as SerialPort;
                ParseDeviceData(port.ReadExisting()); // Read all bytes available
            }
            catch (Exception)
            {
                Debug.WriteLine("Error when receiving serial data.");
                DeviceCleanup();
            }
        }

        /// <summary>
        /// Parse the serial data into usable commands for the music application.
        /// </summary>
        /// <param name="data"></param>
        private static void ParseDeviceData(string dataList)
        {
            //Spaces in Program Path +parameters:
            //CMD /C ""c:\Program Files\demo.cmd"" Parameter1 Param2

            //Spaces in Program Path +parameters with spaces:
            //CMD /K ""c:\batch files\demo.cmd" "Parameter 1 with space" "Parameter2 with space""

            // For each command found in the dataList
            foreach (string cmd in dataList.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                switch (cmd) // Expected data {UP, DOWN, LEFT, RIGHT, A, B, C, RESET, START}
                {
                    case "UP":
                        Debug.WriteLine("Volume Up");
                        ExecuteCommand(foobarExecutable + "\"/command:Up\"");
                        break;

                    case "DOWN":
                        Debug.WriteLine("Volume Down");
                        ExecuteCommand(foobarExecutable + "\"/command:Down\"");
                        break;

                    case "LEFT":
                        Debug.WriteLine("Seek Backwards");
                        ExecuteCommand(foobarExecutable + "\"/command:Back by 30 seconds\"");
                        break;

                    case "RIGHT":
                        Debug.WriteLine("Seek Forwards");
                        ExecuteCommand(foobarExecutable + "\"/command:Ahead by 30 seconds\"");
                        break;

                    case "A":
                        Debug.WriteLine("Previous song");
                        ExecuteCommand(foobarExecutable + "\"/prev\"");
                        break;

                    case "B":
                        Debug.WriteLine("Play/Pause");
                        ExecuteCommand(foobarExecutable + "\"/playpause\"");
                        break;

                    case "C":
                        Debug.WriteLine("Next song");
                        ExecuteCommand(foobarExecutable + "\"/next\"");
                        break;

                    case "RESET":
                        Debug.WriteLine("Delete current song");
                        ExecuteCommand(foobarExecutable + "\"/playing_command:Delete file\""); // invokes the specified context menu command on currently played track
                        break;

                    case "START":
                        Debug.WriteLine("Open/Show foobar2000");
                        ExecuteCommand(foobarExecutable);
                        break;

                    default:
                        Debug.WriteLine("Unknown command received: " + cmd);
                        break;
                }
            }
        }

        /// <summary>
        /// Dispose of any device related resources.
        /// </summary>
        private static void DeviceCleanup()
        {
            Debug.WriteLine("Lost connection to device.");

            if (device != null)
            {
                device.DataReceived -= new SerialDataReceivedEventHandler(DataReceivedHandler);
                if (device.IsOpen) device.Close();
                device.Dispose();
                device = null;
            }
        }

        /// <summary>
        /// Runs a console command through CMD without displaying a GUI.
        /// </summary>
        /// <param name="command"></param>
        public static void ExecuteCommand(string command)
        {
            try
            {
                processStartInfo = new ProcessStartInfo("cmd", "/C " + "\"" + command + "\"")
                {
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (process = new Process())
                {
                    process.StartInfo = processStartInfo;
                    process.Start();
                }
            }
            catch (Exception)
            {
                Debug.WriteLine("Failed to send command.");
            }
        }
    }
}