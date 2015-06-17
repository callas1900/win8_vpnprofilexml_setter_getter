using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.IO;

namespace ConsoleApplication3
{
    class VpnXmlManager
    {

        static string NamespacePath = "\\\\.\\ROOT\\Microsoft\\Windows\\RemoteAccess\\Client";
        static void Main(string[] args)
        {
            switch (args.Length)
            {
                case 0:
                    getVpn();
                    break;
                case 1:
                    setVpn(getVpnProfileFromFilePath(args[0]));
                    break;

            }
        }

        static string getVpnProfileFromFilePath(string filePath)
        {
            List<string> list = new List<string>();
            using (StreamReader reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            return string.Join(string.Empty, list);
        }

        static void getVpn()
        {
            Console.WriteLine("List ============> ");
            
            string ClassName = "PS_VpnConnection";
            LinkedList<ManagementObject> instanceList = new LinkedList<ManagementObject>();

            //Create ManagementClass
            ManagementClass oClass = new ManagementClass(NamespacePath + ":" + ClassName);
            ManagementObjectCollection instances = oClass.GetInstances();
            foreach (ManagementObject instance in instances) {
                instanceList.AddLast(instance);
                Console.WriteLine("{0} : {1}",instanceList.Count.ToString(), instance.Properties["Name"].Value);
            }
            if (instanceList.Count == 0)
            {
                Console.WriteLine("no vpn settings.");
                Console.ReadKey();
                return;
            }
            
            String inputString = "";
            int inputNumber = 0;
            while (!int.TryParse(inputString, out inputNumber))
            {
                Console.WriteLine("plz, type above index number.");
                inputString = Console.ReadLine();
            }
            int counter = 0;
            foreach (ManagementObject o in instanceList)
            {
                if (inputNumber == ++counter)
                {
                    Console.WriteLine(o.Properties["VpnConfigurationXml"].Value);
                }
            }

            Console.ReadKey();

        }

        static void setVpn(string vpnProfile)
        {

            try
            {
                string ClassName = "MSFT_VpnConnection";
                
                //Create ManagementClass
                ManagementClass oClass = new ManagementClass(NamespacePath + ":" + ClassName);
                ManagementBaseObject inParams = oClass.GetMethodParameters("Set");

                // Add the input parameters.
                inParams["Profile"] = vpnProfile;

                // Execute the method and obtain the return values.
                ManagementBaseObject outParams =
                    oClass.InvokeMethod("Set", inParams, null);

                // List outParams
                Console.WriteLine("Done ==>");
                Console.WriteLine(vpnProfile);
            }
            catch (ManagementException err)
            {
                Console.WriteLine("An error occurred while trying to execute the WMI method: " + err.Message);
            }
            Console.ReadKey();
        }
    }
}
