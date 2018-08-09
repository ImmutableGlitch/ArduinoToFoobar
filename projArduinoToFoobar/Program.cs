using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO.Ports;
using System.Runtime.InteropServices;

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

        static void Main(string[] args)
        {
            HideCurrentWindow();

            ConnectToPort();

            Console.ReadLine();
        }

        public static void HideCurrentWindow()
        {
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, 0); // Passing a Zero will hide the window
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
                    string[] ports = SerialPort.GetPortNames();
                    string COM = ports[2];
                    Debug.WriteLine(COM);

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

            ManageSerialData(data);
        }

        private static void ManageSerialData(string data)
        {
            // Location of foobar2000 application
            string foo = "\"C:\\Program Files (x86)\\foobar2000\\foobar2000.exe\""; //does not accept @"C:\Program Files (x86)\foobar2000\foobar2000.exe"

            /* ---- OLD ----
            * 
            * Expected data is in the format 0,0,0 which is three comma separated integers //
            * string[] splitData = data.Split(',');
            * serialData = Array.ConvertAll(splitData, int.Parse);
            */

            switch (data)
            {
                // Expected data is a command:
                // ahead, back, up, down, play, prev, next
                
                case "ahead":
                    // /command:"Ahead by 5 seconds"
                    break;

                case "back":
                    // /command:"Back by 5 seconds"
                    break;

                case "up":
                    // /command:"Volume up"
                    break;

                case "down":
                    // /command:"Volume down"
                    break;

                case "play":
                    ExecuteCommand(foo + " /playpause");
                    break;

                case "prev":
                    ExecuteCommand(foo + " /prev");
                    break;

                case "next":
                    ExecuteCommand(foo + " /next");
                    break;

                default:
                    Debug.WriteLine("Unknown command received");
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