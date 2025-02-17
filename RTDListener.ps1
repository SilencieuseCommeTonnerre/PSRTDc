# Load the necessary COM object for RTD Server.
Add-Type -AssemblyName 'System.Runtime.InteropServices'

Add-Type -TypeDefinition @"
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Runtime.InteropServices;

public class RTD_Com_Server {
    [DllImport("ole32.dll")]
    public static extern int CoInitialize(IntPtr pvReserved);
    [DllImport("ole32.dll")]
    public static extern void CoUninitialize();
   
    private object rtdServerObj;
    private Type rtdType;

    public RTD_Com_Server() {
        CoInitialize(IntPtr.Zero);
        rtdType = Type.GetTypeFromProgID("ToS.RTD");
        if (rtdType != null) {
            rtdServerObj = Activator.CreateInstance(rtdType);
        }
    }

    public string GetData(string topic) {
        if (rtdServerObj != null) {
            try {
                return (string)rtdType.InvokeMember("Topic", System.Reflection.BindingFlags.InvokeMethod, null, rtdServerObj, new object[] { topic });
            } catch {
                return "Error: Invalid topic.";
            }
        } else {
            return "Error: RTD server object not initialized.";
        }
    }

    ~RTD_Com_Server() {
        CoUninitialize();
    }
}
"@

# Create RTD server instance
$rtdComServer = [RTD_Com_Server]::new()

# Function to handle each client connection
function Handle-Client {
    param ($client)
    Write-Host "Handling new client connection..."
    $stream = $client.GetStream()

    while ($client.Connected) {
        $buffer = New-Object byte[] 1024
        $bytesRead = $stream.Read($buffer, 0, $buffer.Length)
       
        if ($bytesRead -gt 0) {
            $data = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $bytesRead).Trim()
            if ($data) {
                Write-Host "Received RTD request for topic: $data"
                $response = $rtdComServer.GetData($data)
                $responseBytes = [System.Text.Encoding]::UTF8.GetBytes($response)
                $stream.Write($responseBytes, 0, $responseBytes.Length)
            }
        }
    }

    $stream.Dispose()
    $client.Dispose()
    Write-Host "Client connection closed."
}

# Setup TCP Listener
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, 13000)

# Handle Ctrl+C (SIGINT) for graceful shutdown
$script:cancelEvent = $false

function OnCancel {
    $script:cancelEvent = $true
    Write-Host "Ctrl+C detected, shutting down..."
}

# Trap Ctrl+C signal
$null = Register-EngineEvent -SourceIdentifier ConsoleControlC -Action { OnCancel }

try {
    $listener.Start()
    Write-Host "TCP Listener started on port 13000..."

    while (-not $script:cancelEvent) {
        if ($listener.Pending()) {
            $client = $listener.AcceptTcpClient()
            Write-Host "Client connected from $($client.Client.RemoteEndPoint)..."
            Start-Job -ScriptBlock { param ($client) Handle-Client $client } -ArgumentList $client
        }
    }
} catch {
    Write-Host "An error occurred: $_"
} finally {
    $listener.Stop()
    Write-Host "TCP Listener stopped."
    Unregister-Event -SourceIdentifier ConsoleControlC
}
