using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace AutoActivationTool
{
    class OfficeActivator
    {
        public static System.Windows.Forms.Label status;
        static string sysDir = "C:/Windows/System32";

        public static void Activate()
        {
            Console.WriteLine("Detecting your all Office versions...");

            string osppFolderPath = GetLatestOsspFolderPath();
            if(osppFolderPath == ""){
                Console.WriteLine("No office version found!");
                return;
            }

            List<string> versionList = GetAllOfficeVersions(osppFolderPath);

            List<string> keyTypes = new List<string>();
            try
            {
                keyTypes = Common.getKeyTypesFromServer();
            }
            catch (Exception e)
            {
                if (e.Message == "")
                {
                    Console.WriteLine("Connect to server fail!");
                }
                else
                {
                    Console.WriteLine(e.Message);
                }

                return;
            }

            for (int i=0;i<versionList.Count;i++)
            {
                ActivateOfficeVersion(versionList[i], osppFolderPath, keyTypes);
            }
        }

        public static string GetLatestOsspFolderPath()
        {
            string[] systems = { "C:/Program Files (x86)/Microsoft Office/", "C:/Program Files/Microsoft Office/" };
            string[] versions = { "Office16", "Office15", "Office14" };

            for (int i = 0; i < versions.Length; i++)
            {
                for (int j = 0; j < systems.Length; j++)
                {
                    string osppFolder = systems[j] + versions[i];
                    string osppPath = osppFolder + "/ospp.vbs";
                    if (System.IO.File.Exists(osppPath))
                    {
                        return osppFolder;
                    }
                }
            }

            return "";
        }

        public static List<string> GetAllOfficeVersions(string osppFolderPath)
        {
            List<string> outList = new List<string>();

            string output = Common.RunCommand("cscript.exe", "ospp.vbs /dstatusall", osppFolderPath);
            List<string> versionList = Common.GetAllValueOfParamFromString(output, "LICENSE NAME", ':');
            List<string> statusList = Common.GetAllValueOfParamFromString(output, "LICENSE STATUS", ':');
            for (int k = 0; k < versionList.Count; k++)
            {
                if(statusList[k].ToUpper().Contains("UNLICENSED"))
                {
                    string version = versionList[k].Replace("MSDN", "");
                    version = version.Replace("edition", "");
                    if (version.Contains("Retail") || version.Contains("MAK"))
                    {
                        bool isExist = false;
                        for (int h = 0; h < outList.Count; h++)
                        {
                            if (outList[h] == version)
                            {
                                isExist = true;
                                break;
                            }
                        }

                        if (!isExist)
                        {
                            outList.Add(version);
                        }
                    }
                }
            }


            return outList;
        }

        public static void ActivateOfficeVersion(string version, string osppFolderPath, List<string> keyTypes)
        {
            Console.WriteLine("Activating " + version + "...");

            string versionNumber = version.Split(',')[0];
            string edition = "";
            string licenseType = ""; 
            if(versionNumber == "Office19")
            {
                edition = version.Split(',')[1].Replace("2019", "");
                edition = edition.Replace("Office19", "");
                edition = edition.Replace("VL_MAK_AE", "");
                edition = edition.Replace("R_Retail", "");

                licenseType = version.Split('_')[1];

                //                 if(licenseType.Contains("Retail"))
                //                 {
                //                     ConvertOffice2019RetailToVL(edition,osppFolderPath);
                //                     licenseType = "MAK";
                //                 }
            }
            else if(versionNumber == "Office16")
            {
                edition = version.Split(',')[1].Replace("Office16", "");
                edition = edition.Replace("VL_MAK", "");
                edition = edition.Replace("R_Retail", "");

                licenseType = version.Split('_')[1];
            }
            else if(versionNumber == "Office15")
            {
                edition = version.Split(',')[1].Replace("Office", "");
                edition = edition.Replace("VL_MAK", "");
                edition = edition.Replace("R_Retail", "");

                licenseType = version.Split('_')[1];
            }
            else if (versionNumber == "Office14")
            {
                versionNumber = "RTM_";

                edition = version.Split(',')[1].Replace("Office", "");
                edition = edition.Replace("-MAK", "");
                edition = edition.Replace("-Retail", "");

                licenseType = version.Split('-')[1];
            }
            else
            {
                return;
            }

            List<string> keyWords = new List<string>();
            keyWords.Add(versionNumber);
            keyWords.Add(edition);
            keyWords.Add(licenseType);

            for (int i = 0; i < keyTypes.Count; i++)
            {
                string keyType = keyTypes[i];
                if (Common.IsKeyTypeConsist(keyType, keyWords))
                {
                    int res = TryActiveByKeyType(keyType,osppFolderPath);
                    if (res > 0)
                    {
                        break;
                    }
                }
            }
        }

        static int TryActiveByKeyType(string keyType, string osppFolderPath)
        {
            List<string> activeKeys = new List<string>();
            try
            {
                activeKeys = Common.getActiveKeysFromServer(keyType);
            }
            catch (Exception e)
            {
                if (e.Message == "")
                {
                    Console.WriteLine("Connect to server fail!");
                }
                else
                {
                    Console.WriteLine(e.Message);
                }

                return 0;
            }

            if (activeKeys.Count == 0)
            {
                return 0;
            }

            for (int i = 0; i < activeKeys.Count; i++)
            {
                string productKey = activeKeys[i];
                string lastFive = Common.getLastFiveOfKey(productKey);
                Console.WriteLine("Applying key " + lastFive + "...");

                int res = TryActiveByUsingKey(productKey, lastFive, osppFolderPath);
                if (res > 0){
                    return 1;
                }
            }

            return 0;
        }

        static int TryActiveByUsingKey(string productKey, string lastFive, string osppFolderPath)
        {
            string install_res = Common.RunCommand("cscript.exe", "ospp.vbs /inpkey:" + productKey, osppFolderPath);
            if (install_res.ToUpper().Contains("SUCCESS"))
            {
                string act_res = Common.RunCommand("cscript.exe", "ospp.vbs /act", osppFolderPath);
                if (act_res.ToUpper().Contains("SUCCESS"))
                {
                    Console.WriteLine("Activation successfully!");
                    return 1;
                }
                else
                {
                    string errorCode = GetErrorCode(act_res, lastFive);
                    if (errorCode == "0xC004C008")
                    {
                        Console.WriteLine("Getting IID...");
                        string iid = GetIID(osppFolderPath);
                        Console.WriteLine(iid);
                        if (Common.IsIDValid(iid))
                        {
                            Console.WriteLine("Getting CID...");
                            string cid = Common.GetCID(iid);
                            Console.WriteLine(cid);

                            if (Common.IsIDValid(cid))
                            {
                                Common.RunCommand("cscript.exe", "ospp.vbs /actcid:" + cid, osppFolderPath);
                                act_res = Common.RunCommand("cscript.exe", "ospp.vbs /act", osppFolderPath);
                                if (act_res.ToUpper().Contains("SUCCESS"))
                                {
                                    Console.WriteLine("Activation successfully!");
                                    return 1;
                                }
                            }
                        }
                    }
                }

                //unpkey
                Common.RunCommand("cscript.exe", "ospp.vbs /unpkey:" + lastFive, osppFolderPath);
            }
            else
            {
                //unpkey
                if(lastFive != ""){
                    Common.RunCommand("cscript.exe", "ospp.vbs /unpkey:" + lastFive, osppFolderPath);
                }

                return -1;
            }

            return 0;
        }

        static string GetErrorCode(string act_res, string lastFive)
        {
            string res = "";
            string[] lines = act_res.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            for (int j = 0; j < lines.Length; j++)
            {
                string line = lines[j];
                if (line.ToUpper().Contains("ERROR CODE:"))
                {
                    string previousLine = lines[j - 1];
                    if (previousLine.Contains(lastFive))
                    {
                        string[] sLines = line.Split(new[] { ":" }, StringSplitOptions.None);
                        res = sLines[1].Replace(" ", "");
                        break;
                    }
                }
            }
            return res;
        }
        static string GetIID(string osppFolderPath)
        {
            string iid = "";

            string dti_res = Common.RunCommand("cscript.exe", "ospp.vbs /dinstid", osppFolderPath);
            string[] lines = dti_res.Split(new[] { "\r\n", "\r", "\n" },
            StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.ToUpper().Contains("ID"))
                {
                    var subLines = line.Split(':');
                    iid = subLines[subLines.Length - 1].Replace(" ", "");
                    break;
                }
            }

            return iid;
        }
        public static void ConvertOffice2019RetailToVL(string edition, string osppFolderPath)
        {
            try
            {
                Common.RunCommand("cscript", "slmgr.vbs /upk 52c4d79f-6e1a-45b7-b479-36b666e0a2f8", sysDir);
                Common.RunCommand("cscript", "slmgr.vbs /upk 5f472f1e-eb0a-4170-98e2-fb9e7f6ff535", sysDir);
                Common.RunCommand("cscript", "slmgr.vbs /upk a3072b8f-adcc-4e75-8d62-fdeb9bdfae57", sysDir);
                Common.RunCommand("cscript", "ospp.vbs /remhst", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /ckms-domain", osppFolderPath);

                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_KMS_Client_AE-ppd.xrm-ms\"", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_KMS_Client_AE-ul.xrm-ms\"", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_KMS_Client_AE-ul-oob.xrm-ms\"", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_MAK_AE-pl.xrm-ms\"", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_MAK_AE-ppd.xrm-ms\"", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_MAK_AE-ul-oob.xrm-ms\"", osppFolderPath);
                Common.RunCommand("cscript", "ospp.vbs /inslic:\"../root/Licenses16/ProPlus2019VL_MAK_AE-ul-phn.xrm-ms\"", osppFolderPath);
            }
            catch(Exception e) { };
        }
        public static string GetOsppFolderPath(string osppFolderName)
        {
            string xNativePath = "C:/Program Files/Microsoft Office/";
            string x86Path = "C:/Program Files (x86)/Microsoft Office/";

            string osppFolderPath = xNativePath + osppFolderName;
            if (!System.IO.File.Exists(osppFolderPath + "/ospp.vbs"))
            {
                osppFolderPath = x86Path + osppFolderName;
                if (!System.IO.File.Exists(osppFolderPath + "/ospp.vbs"))
                {
                    status.Text = "Installed folder not found!";
                }
                else
                {
                    return osppFolderPath;
                }
            }
            else
            {
                return osppFolderPath;
            }

            return "";
        }
    }
}