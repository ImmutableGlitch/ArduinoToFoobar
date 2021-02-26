# ArduinoToFoobar
This is a hideable console application used to read serial data from an Arduino clone.
The data received will be treated as 'commands' to control music playback within a 
program called foobar2000. Every 10 seconds this application checks if the device 
is connected which is useful if it has been unplugged and plugged in again, as 
the connection will be reestablished.

### Read before running...
The foobar2000 file path within the source code may need edited along with the ClassGUID 
and Device Name for your Arduino. The application can also be hidden within the source code.
The ParseDeviceData method needs customised for your specific usage.