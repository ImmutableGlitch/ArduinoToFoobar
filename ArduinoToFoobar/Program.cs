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

    /*
     * TODO: 
     * -read config from JSON or XML file on startup
     *      hardware settings {portName, baudRate, parity, dataBits, stopBits}
     *      misc {connectionInterval, logPath}
     *      profiles {profileName, targetExecutable, }
     */

    class Program
    {
        enum MapperProfile
        {
            foobar,
            video
        }

        static MapperProfile profile;

        static SerialPort device = null;
        static string portName = "";

        static string expectedClassGuid   = ""; //{4d36e978-e325-11ce-bfc1-08002be10318}
        static string expectedDisplayName = ""; //USB-SERIAL CH340 (COM5)
        static string targetExecutable    = ""; //C:\Program Files (x86)\foobar2000\foobar2000.exe

        static int connectionInterval = 30000;


        static Process process;
        static ProcessStartInfo processStartInfo;

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static void Main()
        {
            //////////[Temporary test]
            profile = MapperProfile.foobar;
            targetExecutable = @"""C:\Program Files (x86)\foobar2000\foobar2000.exe""";
            portName = "COM13";
            //////////
            
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, 0); // Passing a Zero to hide this console window

            var timer = new Timer(ConnectionWatchdog, null, 0, connectionInterval); // Check connection every x seconds

            Console.ReadLine(); // Keep program alive
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
                //if (FindSerialDeviceByName()) // Find correct device
                //{
                    ConnectToSerialDevice(); // Connect to it
                //}
            }
            else if (!device.IsOpen) // if lost connection to device
            {
                DeviceCleanup(); // Close and dispose serial port
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
                "SELECT * FROM Win32_PnPEntity WHERE ClassGuid = \"" + expectedClassGuid + "\"");

            foreach (ManagementObject objectFound in search.Get())
            {
                string displayName = objectFound["Caption"].ToString();

                if (displayName.Contains(expectedDisplayName))
                {
                    Regex rx = new Regex(@"(COM\d+)(?!\()"); //TODO: check if working for double digit number such as COM13
                    portName = rx.Match(displayName).Value;
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
                device = new SerialPort(portName, 9600, Parity.None, 8, StopBits.One);
                device.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
                device.Open();
                Debug.WriteLine("Successful connection to serial device " + portName);
            }
            catch (Exception)
            {
                DeviceCleanup();
                Debug.WriteLine("Failed to connect to serial device " + portName);
            }
        }

        /// <summary>
        /// Dispose of any device related resources.
        /// </summary>
        private static void DeviceCleanup()
        {
            if (device != null)
            {
                Debug.WriteLine("Closing connection to device.");

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

            // For each command found in the dataList
            foreach (string cmd in dataList.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                if(profile == MapperProfile.foobar)
                {
                    switch (cmd) // Expected data {UP, DOWN, LEFT, RIGHT, A, B, C, RESET, START}
                    {
                        case "UP":
                            Debug.WriteLine("Volume Up");
                            ExecuteCommand("\"/command:Up\"");
                            break;

                        case "DOWN":
                            Debug.WriteLine("Volume Down");
                            ExecuteCommand("\"/command:Down\"");
                            break;

                        case "LEFT":
                            Debug.WriteLine("Seek Backwards");
                            ExecuteCommand("\"/command:Back by 30 seconds\"");
                            break;

                        case "RIGHT":
                            Debug.WriteLine("Seek Forwards");
                            ExecuteCommand("\"/command:Ahead by 30 seconds\"");
                            break;

                        case "A":
                            Debug.WriteLine("Previous song");
                            ExecuteCommand("\"/prev\"");
                            break;

                        case "B":
                            Debug.WriteLine("Play/Pause");
                            ExecuteCommand("\"/playpause\"");
                            break;

                        case "C":
                            Debug.WriteLine("Next song");
                            ExecuteCommand("\"/next\"");
                            break;

                        case "RESET":
                            Debug.WriteLine("Delete current song");
                            ExecuteCommand("\"/playing_command:Delete file\""); // invokes the specified context menu command on currently played track
                            break;

                        case "START":
                            Debug.WriteLine("Open/Show foobar2000");
                            ExecuteCommand("");
                            break;

                        default:
                            Debug.WriteLine("Unknown command received: " + cmd);
                            break;

                    }
                }
                else if (profile == MapperProfile.video)
                {
                    Debug.WriteLine(cmd);

                    switch (cmd)
                    {
                        case "scrollUp":
                            ExecuteCommand("nircmdc sendmouse wheel 120");
                            break;

                        case "scrollDown":
                            ExecuteCommand("nircmdc sendmouse wheel -120");
                            break;

                        case "flickRight":
                            ExecuteCommand("nircmdc sendkey right press");
                            break;

                        case "flickLeft":
                            ExecuteCommand("nircmdc sendkey left press");
                            break;

                        default:
                            Debug.WriteLine("Unknown command received: " + cmd);
                            break;
                    }
                }
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
                processStartInfo = new ProcessStartInfo("cmd", $" /C \"{targetExecutable} {command}\" ") // cmd /C "exePath args"
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