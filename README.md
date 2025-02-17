Microsoft uses a protocol to allow an application to serve reatime data to other applications, most notably Excel.
This project is an attempt to create a universal client via a PowerShell script

The idea here is to create a COM object that is both an RTD client and an HTTP server.

The code should listen for HTTP clients that connect and ask for an RTD topic, then pass the request to the RTD server.  It should then relay the response to the client.

I am the world's worst coder, so any help is appreciated. Thank you.
