# VB6 Send Files Via SerialPort
### A little example how we can send files via RS232 to another PC or device.
This example shows how to send binary data to another PC or device via RS-232.
To send the file this app opens it in binary mode and send **byte by byte** every 50 ms. The other device should check serial buffer every 50 ms. To avoid more than 1 byte in serial buffer, ...inBufferSize... property is established to 1 byte.


