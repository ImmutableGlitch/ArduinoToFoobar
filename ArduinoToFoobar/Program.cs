using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace ArduinoToFoobar
{

    /*
     * This is a hidden console application ran in the background which reads serial data 
     * from an Arduino clone. The data received will be treated as 'commands' to control music 
     * playback within a program called foobar2000. This software still works if the
     * device is unplugged and plugged in again.
     */

    class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static System.Threading.Timer t;
        static string HARDWARE_ID;
        static string COM;
        static SerialPort device;

        static void Main()
        {
            HideConsoleWindow();

            HARDWARE_ID = @"USB\VID_1A86&PID_7523\5&E658374&0&14";
            COM = string.Empty;
            device = null;
            t = new System.Threading.Timer(ConnectionWatchdog, null, 0, 60000);

            Console.ReadLine(); // Keep program alive
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
                if (FindDeviceByID())
                {
                    ConnectToSerialDevice();
                }
            }
            else if (!device.IsOpen) // if connection to existing device is closed
            {
                DeviceCleanup();
            }
        }

        /// <summary>
        /// Enumerates hardware connected to computer.
        /// Identifies specific device to use based on a hardware ID.
        /// </summary>
        /// <returns>bool indicating whether expected device was found</returns>
        private static bool FindDeviceByID()
        {
            //using System.Management;
            List<ManagementObject> devices;

            try
            {
                string query = "SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0";// get only devices that are working properly."
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
                devices = searcher.Get().Cast<ManagementObject>().ToList();
                searcher.Dispose();
            }
            catch (Exception)
            {
                MessageBox.Show("Error from listing devices!");
                return false;
            }

            object deviceCaption;
            string caption;

            foreach (ManagementObject dev in devices)
            {
                deviceCaption = dev["Caption"];
                if (deviceCaption != null)
                {
                    caption = deviceCaption.ToString();
                    if (caption.Contains("(COM"))
                    {
                        //Debug.WriteLine("Caption: " + caption + "\nDeviceID: " + dev["DeviceID"] + "\n\n");

                        if (dev["DeviceID"].ToString() == HARDWARE_ID)
                        {
                            // "USB-SERIAL CH340 (COM30)"
                            // should equal COM30 after regex
                            Regex rx = new Regex(@"(COM\d+)(?!\()");
                            COM = rx.Match(caption).Value;
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Attempts to open a serial port connection to enable communication.
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
        /// Dispose of any device related resources.
        /// </summary>
        private static void DeviceCleanup()
        {
            Debug.WriteLine("Cleaning up device resources.");

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
                //var data = device.ReadLine().Trim();
                ParseDeviceData(data);
            }
            catch(Exception)
            {
                Debug.WriteLine("Error when receiving serial data.");
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
            // does not accept @"C:\Program Files (x86)\foobar2000\foobar2000.exe"
            string program = "\"C:\\Program Files (x86)\\foobar2000\\foobar2000.exe\"";

            switch (data)
            {
                // Expected data {AHEAD, BACK, UP, DOWN, PLAYPAUSE, PREVIOUS, NEXT}

                case "AHEAD":
                    Debug.WriteLine("Ahead");
                    ExecuteCommand(program + " /command:Ahead by 5 seconds");
                    break;

                case "BACK":
                    Debug.WriteLine("Back");
                    ExecuteCommand(program + " /command:Back by 5 seconds");
                    break;

                case "UP":
                    Debug.WriteLine("Up");
                    ExecuteCommand(program + " /command:Up");
                    break;

                case "DOWN":
                    Debug.WriteLine("Down");
                    ExecuteCommand(program + " /command:Down");
                    break;

                case "PLAYPAUSE":
                    Debug.WriteLine("Play/Pause");
                    ExecuteCommand(program + " /playpause");
                    break;

                case "PREVIOUS":
                    Debug.WriteLine("Previous");
                    ExecuteCommand(program + " /prev");
                    break;

                case "NEXT":
                    Debug.WriteLine("Next");
                    ExecuteCommand(program + " /next");
                    break;

                default:
                    Debug.WriteLine("Unknown command received: " + data);
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
                ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd", "/c " + command)
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
                Debug.WriteLine("Failed to send command.");
            }
        }
    }
}