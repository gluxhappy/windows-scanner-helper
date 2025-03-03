using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WIA;

namespace WindowPrinterHelper
{
    class WmiScanner
    {

        public void list()
        {
            listScannersAsync();
        }

        public void scan(Dictionary<string, string> scannerFilterProperties, Dictionary<string, string> feederProperties, Dictionary<string, string> settingProperties, string format, string output)
        {
            DeviceManager deviceManager = new DeviceManager();
            Device? scanner = findScanner(scannerFilterProperties);
            if (scanner == null)
            {
                Console.WriteLine("No scanner found.");
                return;
            }
            Item? item = findFeeder(scanner, feederProperties);
            if (item == null)
            {
                Console.WriteLine("No feeder found.");
                return;
            }
            if (settingProperties.Count > 0)
            {
                foreach (var setting in settingProperties)
                {
                    item.Properties[setting.Key].set_Value(setting.Value);
                }
            }
            string scanProfile = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
            switch (format)
            {
                case "jpg":
                    scanProfile = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"; // FormatID.wiaFormatJPEG;
                    break;
                case "png":
                    scanProfile = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
                    break;
                default:
                    Console.WriteLine("Invalid format.");
                    break;
            }
            ImageFile imageFile = (ImageFile)item.Transfer(scanProfile);

            // Delete the output file if it already exists
            if (File.Exists(output))
            {
                File.Delete(output);
            }

            // Save the image as a JPEG
            string filePath = output;
            imageFile.SaveFile(filePath);
        }

        private Device findScanner(Dictionary<string, string> scannerFilterProperties)
        {
            
            if (scannerFilterProperties == null || scannerFilterProperties.Count == 0)
            {
                return new WIA.CommonDialog().ShowSelectDevice(WiaDeviceType.ScannerDeviceType, true, false);
            }
            DeviceManager deviceManager = new DeviceManager();
            Device? scanner = null;
            foreach (DeviceInfo deviceInfo in deviceManager.DeviceInfos)
            {
                if (deviceInfo.Type == WiaDeviceType.ScannerDeviceType)
                {
                    bool found = true;
                    foreach (var filter in scannerFilterProperties)
                    {
                        if (deviceInfo.Properties[filter.Key].get_Value().ToString() != filter.Value)
                        {
                            found = false;
                            break;
                        }
                    }
                    if (found)
                    {
                        scanner = deviceInfo.Connect();
                        break;
                    }
                }
            }
            return scanner;
        }

        private Item? findFeeder(Device scanner, Dictionary<string, string>? feederProperties)
        {
            if (feederProperties == null || feederProperties.Count == 0)
            {
                return scanner.Items[1];
            }
            Item? feeder = null;
            foreach (Item item in scanner.Items)
            {
                bool found = true;
                foreach (var filter in feederProperties)
                {
                    if (item.Properties[filter.Key].get_Value().ToString() != filter.Value)
                    {
                        found = false;
                        break;
                    }
                }
                if (found)
                {
                    feeder = item;
                    break;
                }
            }
            return feeder;
        }

        private void listScannersAsync()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("Call with arguments to scan:");
            stringBuilder.AppendLine("    --scanner   used to find the scanner, can use multiple times, i.e. --scanner=Name:HP530, \"--scanner=Private Productnumber:4SB25A\"");
            stringBuilder.AppendLine("    --feeder    used to filter the document provider of the scanner, can use multiple times, i.e. --feeder=Planar:0, \"--feeder=Item Name:Scan\"");
            stringBuilder.AppendLine("    --setting   used to specify the settings of the scanner, for example, set the resolution, \"--setting=Horizontal Resolution:100\" \"--setting=Vertical Resolution:100\"");
            stringBuilder.AppendLine("    --output    specify a place to save the scan result image, default scan.jpg");
            stringBuilder.AppendLine("    --format    jpg, png, default jpg");
            stringBuilder.AppendLine("");
            stringBuilder.AppendLine("Available scanners and feeders on this host are:");
            DeviceManager deviceManager = new DeviceManager();
            foreach (DeviceInfo deviceInfo in deviceManager.DeviceInfos)
            {
                if (deviceInfo.Type == WiaDeviceType.ScannerDeviceType)
                {
                    string name = deviceInfo.Properties["Name"].get_Value();
                    Device? scanner = null;
                    string? error = "TIMEOUT";

                    Thread t = new Thread(() =>
                    {
                        try
                        {
                            scanner = deviceInfo.Connect();
                        }
                        catch (Exception e)
                        {
                            error = e.Message;
                        }
                    });
                    t.Start();
                    if (!t.Join(5 * 1000))
                    {
                        stringBuilder.AppendLine($"{name}:CONNECT TIMEOUT");
                        continue;
                    }
                    if (scanner == null)
                    {
                        stringBuilder.AppendLine($"{name}:CONNECT ERROR {error}");
                        continue;
                    }
                    stringBuilder.AppendLine($"{name}");
                    foreach (Property property in scanner.Properties)
                    {
                        var propName = property.Name;
                        var propValue = property.get_Value();
                        stringBuilder.AppendLine("  \\-" + propName + " : " + propValue);
                    }
                    foreach (Item item in scanner.Items)
                    {
                        stringBuilder.AppendLine("  \\-0");
                        foreach (Property property in item.Properties)
                        {
                            var propName = property.Name;
                            var propValue = property.get_Value();
                            stringBuilder.AppendLine("    \\-" + propName + " : " + propValue);
                        }
                    }

                }
            }

            Form f = new Form();
            f.Text = "Scanners";
            TextBox textBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Text = stringBuilder.ToString()
            };
            f.Controls.Add(textBox);
            f.Size = new System.Drawing.Size(800, 600);
            textBox.Text = stringBuilder.ToString();
            f.ShowDialog();
        }
    }
}

