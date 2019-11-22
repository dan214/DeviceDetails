using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;

namespace SystemInformation
{
    public class SystemInformation
    {
        public static Dictionary<String,String> SystemInfo = new Dictionary<string, string>();

        public static string OutputFileDirectory = @"C:\\Temp";
        public SystemInformation()
        {
            GetDeviceDetails();
            SaveSystemInformationtoCsv();
            GetSystemInformation();

        }
        

        private void GetSystemInformation()
        {
            var organizationUrl = new Uri("https://dev.azure.com/TBL-Azure");   // Organization URL, for example: https://dev.azure.com/fabrikam               
            const string personalAccessToken = "5ca64zzub7p6sgxtbvyhogfck4qzlzrrfgdbi26vlvgnadjthgzq"; // See https://docs.microsoft.com/azure/devops/integrate/get-started/authentication/pats
            const int workItemId = 3; // ID of a work item, for example: 12

            // Create a connection
            VssConnection connection = new VssConnection(organizationUrl, new VssBasicCredential(string.Empty, personalAccessToken));

            // Show details a work item
            UpdateWorkItemDetails(connection, workItemId).Wait();

        }

        private static void GetDeviceDetails()
        {
            SystemInfo.Add("System","");
            var query = new SelectQuery(@"Select * from Win32_ComputerSystem");
            using (var searcher =
                new ManagementObjectSearcher(query))
            {
                //execute the query
                foreach (var o in searcher.Get())
                {
                    var process = (ManagementObject) o;
                    //print system info
                    process.Get();

                    SystemInfo.Add("Computer Name",process["Name"].ToString());
                    SystemInfo.Add("Computer Username", process["UserName"].ToString());

                }
            }
            SystemInfo.Add("", "");
            SystemInfo.Add("Operating System","");

            var searcher2 = new ManagementObjectSearcher(@"SELECT * FROM Win32_OperatingSystem");
            var colItems = searcher2.Get();

            foreach (var o in colItems)
            {
                var queryObj = (ManagementObject) o;

                try
                {
                    SystemInfo.Add("Name", queryObj["Name"].ToString());
                    SystemInfo.Add("BootDevice", queryObj["BootDevice"].ToString());
                    SystemInfo.Add("BuildNumber", queryObj["BuildNumber"].ToString());
                    SystemInfo.Add("Caption", queryObj["Caption"].ToString());

                }
                catch
                {
                    // ignored
                }
            }

            ManagementObjectSearcher searcher3 = new ManagementObjectSearcher(@"SELECT * FROM Win32_DriverForDevice");
            ManagementObjectCollection colItems1 = searcher3.Get();

            SystemInfo.Add("Drivers", "");

            Console.WriteLine("{0} instance{1}\n", colItems1.Count, (colItems1.Count == 1 ? String.Empty : "s"));

            foreach (var o in colItems1)
            {
                var queryObj = (ManagementObject) o;
                SystemInfo.Add(queryObj["Antecedent"].ToString(), $"{queryObj["Dependent"]}");
            }

        }

        private static void EnsureDirectoryExists(string v)
        {
            if (!Directory.Exists(v))
            {
                Directory.CreateDirectory(v);
            }
        }

        private static void SaveSystemInformationtoCsv()
        {
            try
            {
                //WriteObject("Writing to CSV..!");
                var csv = new StringBuilder();

                foreach (var newLine in SystemInfo)
                {
                    csv.AppendLine($"{newLine.Key},{newLine.Value}");
                }

                EnsureDirectoryExists(OutputFileDirectory);
                File.WriteAllText(OutputFileDirectory + "\\SystemInformation.csv", csv.ToString());


            }
            catch (Exception)
            {
                //WriteObject(ex.Message);
            }
        }

        private static async Task UpdateWorkItemDetails(VssConnection connection, int workItemId)
        {

            var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

            try
            {
                const string filePath = @"C:\\Temp\\SystemInformation.csv";


                //add attachment
                var attachment = witClient.CreateAttachmentAsync(filePath).GetAwaiter().GetResult();

                var patchDocument = new JsonPatchDocument
                {
                    new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/relations/-",
                        Value = new
                        {
                            rel = "AttachedFile",
                            url = attachment.Url,
                            attributes = new {comment = "Adding new attachment for Test Case"}
                        }
                    }
                };
                var result = witClient.UpdateWorkItemAsync(patchDocument, 3).Result;




            }
            catch (AggregateException aex)
            {
                if (aex.InnerException is VssServiceException vssex)
                {
                    Console.WriteLine(vssex.Message);
                }
            }
        }
    }
}
