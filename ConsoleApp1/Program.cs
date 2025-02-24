using System.IO.Ports;
using System.Management;
using System.Collections.Generic;
using System.Linq;

class PortLister
{
    static void Main()
    {
        var allPorts = new HashSet<(string port, string description, string status)>();

        // Method 1: SerialPort.GetPortNames()
        Console.WriteLine("\nChecking SerialPort.GetPortNames():");
        try
        {
            foreach (string port in SerialPort.GetPortNames())
            {
                Console.WriteLine($"Found port: {port}");
                allPorts.Add((port, "Via SerialPort.GetPortNames()", "Listed"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting SerialPort.GetPortNames(): {ex.Message}");
        }

        // Method 2: WMI Query for PnP Serial Ports
        Console.WriteLine("\nChecking WMI for PnP Serial Ports:");
        try
        {
            using (var searcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_PnPEntity WHERE Caption LIKE '%(COM%'"))
            using (var collection = searcher.Get())
            {
                foreach (var device in collection)
                {
                    string caption = device["Caption"]?.ToString() ?? "";
                    string status = device["Status"]?.ToString() ?? "Unknown";

                    if (caption.Contains("(COM"))
                    {
                        string port = "COM" + caption.Split('(', ')')[1].Split('M')[1];
                        Console.WriteLine($"Found device: {caption} - Status: {status}");
                        allPorts.Add((port, caption, status));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error querying WMI: {ex.Message}");
        }

        // Method 3: Direct WMI query for Serial Ports
        Console.WriteLine("\nChecking WMI for Serial Ports:");
        try
        {
            using (var searcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_SerialPort"))
            using (var collection = searcher.Get())
            {
                foreach (var device in collection)
                {
                    string name = device["DeviceID"]?.ToString() ?? "";
                    string description = device["Description"]?.ToString() ?? "";
                    string status = device["Status"]?.ToString() ?? "Unknown";

                    Console.WriteLine($"Found device: {name} - {description} - Status: {status}");
                    allPorts.Add((name, description, status));
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error querying Serial Ports WMI: {ex.Message}");
        }

        // Summary of all the ports found from the three methods
        Console.WriteLine("\nSummary of all COM ports found:");
        Console.WriteLine("--------------------------------");
        foreach (var port in allPorts.OrderBy(p => p.port))
        {
            Console.WriteLine($"Port: {port.port}");
            Console.WriteLine($"Description: {port.description}");
            Console.WriteLine($"Status: {port.status}");
            Console.WriteLine("--------------------------------");
        }

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}

