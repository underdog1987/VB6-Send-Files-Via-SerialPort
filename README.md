# VB6 Send Files Via Serial Port
### A little example how we can send files via RS232 to another PC or device.
This example shows how to send binary data to another PC or device via RS-232.
To send the file this app opens it in binary mode and send **byte by byte** every 50 ms. The other device should check serial buffer every 50 ms. To avoid more than 1 byte in serial buffer, _inBufferSize_ property is established to 1 byte.
The max file size to send is determinated by sizeof(integer) due every byte of the file is places in an array element and the max size of an array is determinated by sizeof(integer) too.
#### Test the file in a single computer
You need a PC with available serial port. Just connect pins 2 and 3 (RX/TX).

