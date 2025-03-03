
namespace WindowPrinterHelper
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length > 0)
                {
                    Dictionary<string, string> scannerProperties = new Dictionary<string, string>();
                    Dictionary<string, string> feederProperties = new Dictionary<string, string>();
                    Dictionary<string, string> settingProperties = new Dictionary<string, string>();
                    string output = "scan.jpg";
                    string format = "jpg";

                    for (int i = 0; i < args.Length; i++)
                    {
                        string arg = args[i];
                        if (arg.StartsWith("--scanner="))
                        {
                            addToProperties(scannerProperties, "--scanner=", arg);
                        }
                        else if (arg.StartsWith("--feeder="))
                        {
                            addToProperties(feederProperties, "--feeder=", arg);
                        }
                        else if (arg.StartsWith("--setting="))
                        {
                            addToProperties(settingProperties, "--setting=", arg);
                        }
                        else if (arg.StartsWith("--output="))
                        {
                            output = arg.Substring("--output=".Length);
                        }
                        else if (arg.StartsWith("--format="))
                        {
                            format = arg.Substring("--format=".Length);
                        }
                    }
                    new WmiScanner().scan(scannerProperties, feederProperties, settingProperties, format, output);
                }
                else
                {
                    new WmiScanner().list();
                }
                Environment.Exit(0);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }
        }

        private static void addToProperties(Dictionary<string, string> properties,string prefix, string key)
        {
            string value = key.Substring(prefix.Length);
            int idx = value.IndexOf(':');
            string name = value.Substring(0, idx).Trim();
            string val = value.Substring(idx + 1).Trim();
            properties.Add(name, val);
        }
    }
}