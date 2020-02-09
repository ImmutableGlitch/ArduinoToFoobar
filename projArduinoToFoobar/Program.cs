using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using System.Management;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace projArduinoToFoobar
{

    /*
     * This is a hidden program ran in the background which reads serial data 
     * from an Arduino with the data being 'commands'
     *
     * The commands received will allow the control of music playback within foobar2000
     */

    class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static SerialPort port;
        static string COM;

        static void Main(string[] args)
        {
            HideCurrentWindow();

            if (FindArduinoByID(@"USB\VID_1A86&PID_7523\5&E658374&0&14")) { 
                ConnectToPort();
            }

            Console.ReadLine();
        }

        public static void HideCurrentWindow()
        {
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, 0); // Passing a Zero will hide the window
        }

        public static bool FindArduinoByID(string arduinoID)
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
            catch (Exception ex)
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
                        //Console.WriteLine("Caption: {0}\nDeviceID: {1}\n\n", caption, dev["DeviceID"]);

                        if (dev["DeviceID"].ToString() == arduinoID) // Current Arduino clone to use
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

        public static void ConnectToPort()
        {
            Debug.WriteLine("Beginning connection to serial port");

            bool connected = false;

            // Try to connect to the serial port
            // Allow user to retry connection or abort on error
            do
            {
                try
                {
                    //string[] ports = SerialPort.GetPortNames();
                    //string COM = ports[2];
                    //Debug.WriteLine(COM);

                    port = new SerialPort(COM, 9600, Parity.None, 8, StopBits.One);
                    // Create event handler
                    port.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
                    port.Open();
                    connected = true;
                }
                catch (Exception)
                {
                    var choice = MessageBox.Show("Connection to COM Port failed.", "Error", MessageBoxButtons.RetryCancel);

                    if (choice == DialogResult.Cancel)
                    {
                        return;
                    }
                }

            } while (connected == false);
        }

        private static void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort port = (SerialPort)sender;

            string data = port.ReadLine();
            Debug.WriteLine("Receiving data from port: " + data);

            ManageSerialData(data.Trim());
        }

        private static void ManageSerialData(string data)
        {
            // Location of foobar2000 application
            //does not accept @"C:\Program Files (x86)\foobar2000\foobar2000.exe"
            string foo = "\"C:\\Program Files (x86)\\foobar2000\\foobar2000.exe\"";

            switch (data)
            {
                // Expected data is a command:
                // ahead, back, up, down, play, prev, next

                //case "ahead":
                //    Console.WriteLine("Ahead");
                //     /command:"Ahead by 5 seconds"
                //    break;

                //case "back":
                //    Console.WriteLine("Back");
                //     /command:"Back by 5 seconds"
                //    break;

                case "UP":
                    Console.WriteLine("Up");
                    ExecuteCommand(foo + " /command:Up");
                    break;

                case "DOWN":
                    Console.WriteLine("Down");
                    ExecuteCommand(foo + " /command:Down");
                    break;

                case "PLAYPAUSE":
                    Console.WriteLine("PlayPause");
                    ExecuteCommand(foo + " /playpause");
                    break;

                case "PREVIOUS":
                    Console.WriteLine("Prev");
                    ExecuteCommand(foo + " /prev");
                    break;

                case "NEXT":
                    Console.WriteLine("Next");
                    ExecuteCommand(foo + " /next");
                    break;

                default:
                    //Console.WriteLine("Unknown command");
                    Console.WriteLine(data);
                    break;
            }
        }

        public static void ExecuteCommand(string command)
        {
            // Runs a command through CMD with no GUI
            ProcessStartInfo procStartInfo = new ProcessStartInfo("cmd", "/c " + command)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process proc = new Process())
            {
                proc.StartInfo = procStartInfo;
                proc.Start();
            }
        }
    }
}