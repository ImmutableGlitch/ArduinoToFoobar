using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Runtime.InteropServices;

namespace projArduinoToFoobar
{

    /*
     * This is a hidden program ran in the background which reads serial data from an Arduino.
     * When a button is pressed on the Arduino it sends a serial command which is read by this program.
     * The commands will allow the control of music playback within foobar2000
     */

    class Program
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        static SerialPort port;

        static void Main(string[] args)
        {
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            // Passing a Zero will hide the window
            ShowWindow(winHandle, 0);
            
            Debug.WriteLine("Listening to messages");

            try
            {
                // Use third COM port and create event handler
                string[] ports = SerialPort.GetPortNames();
                string COM = ports[2];
                Debug.WriteLine(COM);
                port = new SerialPort(COM, 9600, Parity.None, 8, StopBits.One);
                port.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
                port.Open();
            }
            catch (Exception)
            {
                Debug.WriteLine("COM Port error");
                return;
            }

            Console.ReadLine();
        }

        private static void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort port = (SerialPort)sender;
            string data = port.ReadLine();
            Debug.WriteLine(data);

            // Arduino sent PP with carriage return
            if (data != String.Empty && data.Equals("PP\r"))
            {
                string foo = "\"E:\\Data\\Programs\\Installers\\Fresh_Install\\Windows_Settings\\Program Files (x86)\\foobar2000\\foobar2000.exe\"";
                ExecuteCommand(foo + " /playpause");
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