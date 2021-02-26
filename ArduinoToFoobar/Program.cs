using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace ArduinoToFoobar
{

    /*
     * This is a hideable console application used to read serial data from an Arduino clone.
     * The data received will be treated as 'commands' to control music playback within a 
     * program called foobar2000. Every 10 seconds this application checks if the device 
     * is connected which is useful if it has been unplugged and plugged in again, as 
     * the connection will be reestablished.
     */

    class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static System.Threading.Timer t;
        static string ExpectedClassGuid;
        static string ExpectedDisplayName;
        static string COM;
        static SerialPort device;
        

        static void Main()
        {
            //HideConsoleWindow();

            COM = string.Empty;
            device = null;

            // Determined using Device Manager to check device properties
            ExpectedClassGuid   = "{4d36e978-e325-11ce-bfc1-08002be10318}";
            ExpectedDisplayName = "USB-SERIAL CH340 (COM";

            // Check the connection state every 10 seconds
            t = new System.Threading.Timer(ConnectionWatchdog, null, 0, 10000);

            // Keep program alive
            Console.ReadLine();

            //TODO: intercept the closing of console window with ConsoleCancelEventHandler etc
        }

        /// <summary>
        /// Hides the console window using a function built into Windows OS
        /// </summary>
        private static void HideConsoleWindow()
        {
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, 0); // Passing a Zero will hide the window
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
                Console.WriteLine("Successful connection to serial device " + COM);
            }
            catch (Exception)
            {
                DeviceCleanup();
                Console.WriteLine("Failed to connect to serial device " + COM);
            }
        }

        /// <summary>
        /// Dispose of any device related resources.
        /// </summary>
        private static void DeviceCleanup()
        {
            Console.WriteLine("Lost connection to device.");

            if (device != null)
            {
                device.DataReceived -= new SerialDataReceivedEventHandler(DataReceivedHandler);
                if (device.IsOpen) device.Close();
                device.Dispose();
                device = null;
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
                var data = (sender as SerialPort).ReadLine().Trim();
                ParseDeviceData(data);
            }
            catch(Exception)
            {
                Console.WriteLine("Error when receiving serial data.");
                DeviceCleanup();
            }
        }

        /// <summary>
        /// Parse the serial data into usable commands for the music application.
        /// </summary>
        /// <param name="data"></param>
        private static void ParseDeviceData(string data)
        {
            // Location of foobar2000 music application
            string program = "\"C:\\Program Files (x86)\\foobar2000\\foobar2000.exe\" ";

            //Spaces in Program Path +parameters:
            //CMD /C ""c:\Program Files\demo.cmd"" Parameter1 Param2

            //Spaces in Program Path +parameters with spaces:
            //CMD /K ""c:\batch files\demo.cmd" "Parameter 1 with space" "Parameter2 with space""

            switch (data) // Expected data {UP, DOWN, LEFT, RIGHT, A, B, C, RESET, START}
            {
                case "UP":
                    Console.WriteLine("Volume Up");
                    ExecuteCommand(program + "\"/command:Up\"");
                    break;

                case "DOWN":
                    Console.WriteLine("Volume Down");
                    ExecuteCommand(program + "\"/command:Down\"");
                    break;

                case "LEFT":
                    Console.WriteLine("Seek Backwards");
                    ExecuteCommand(program + "\"/command:Back by 30 seconds\"");
                    break;

                case "RIGHT":
                    Console.WriteLine("Seek Forwards");
                    ExecuteCommand(program + "\"/command:Ahead by 30 seconds\"");
                    break;

                case "A":
                    Console.WriteLine("Previous song");
                    ExecuteCommand(program + "\"/prev\"");
                    break;

                case "B":
                    Console.WriteLine("Play/Pause");
                    ExecuteCommand(program + "\"/playpause\"");
                    break;

                case "C":
                    Console.WriteLine("Next song");
                    ExecuteCommand(program + "\"/next\"");
                    break;

                case "RESET":
                    Console.WriteLine("Delete current song");
                    //ExecuteCommand(program + "\"/playing_command:<context menu command>\""); // invokes the specified context menu command on currently played track
                    break;

                case "START":
                    Console.WriteLine("Open/Show foobar2000");
                    ExecuteCommand(program);
                    break;

                default:
                    Console.WriteLine("Unknown command received: " + data);
                    break;
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
                ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd", "/C " + "\"" + command + "\"")
                {
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (Process process = new Process())
                {
                    process.StartInfo = processStartInfo;
                    process.Start();
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Failed to send command.");
            }
        }
    }
}