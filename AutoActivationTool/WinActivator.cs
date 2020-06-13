using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace AutoActivationTool
{
    class WinActivator
    {
        static string sysDir = "C:/Windows/System32";
        public static void Activate()
        {
            Console.WriteLine("Detecting your Windows version...");
            string winVer = GetWinVer();
            Console.WriteLine(winVer);

            Console.WriteLine("Detecting your license status...");
            string licenseInfo = GetLicenseTypeAndLicenseStatus();
            string licenseStatus = licenseInfo.Split(',')[0];
            string licenseType = licenseInfo.Split(',')[1];
            Console.WriteLine(licenseStatus);

            if(licenseStatus.Contains("Licensed"))
            {
                Console.WriteLine("Your Windows is activated. No action needed.");

                return;
            }

            List<string> keyWords = GetKeyWordsFromWinVerAndLicenseType(winVer,licenseType);

//             for(int i=0;i<keyWords.Count;i++)
//             {
//                 Console.Write(keyWords[i] + " ");
//             }
//             Console.WriteLine("");

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

            for(int i = 0; i < keyTypes.Count; i++)
            {
                string keyType = keyTypes[i];

                if (Common.IsKeyTypeConsist(keyType, keyWords))
                {
                    int res = TryActiveByKeyType(keyType);
                    if (res > 0)
                    {
                        break;
                    }
                }
            }
        }

        public static List<string> GetKeyWordsFromWinVerAndLicenseType(string winVer, string licenseType)
        {
            if(winVer.Contains(" for"))
            {
                winVer = winVer.Replace(" for", "");
            }
            else if (winVer.Contains("LTSC"))
            {
                winVer = winVer.Replace("LTSC", "S");
            }
            else if (winVer.Contains("LTSB"))
            {
                winVer = winVer.Replace("LTSB", "S");
            }
            else if (winVer.Contains("2012"))
            {
                winVer = winVer.Replace("2012", "12");
            }

            List<string> outWords = new List<string>();

            string[] words = winVer.Split(' ');

            int index = 1;
            if (words[index].Contains("Win"))
            {
                outWords.Add("Win");
                index++;
            }

            if (words[index].Contains("Server"))
            {
                outWords.Add("Server");
                index++;
            }

            string winverNum = words[index];
            outWords.Add(words[index]);
            index++;

            if (winverNum == "10")
            {
                if (words[index].Contains("Home"))
                {
                    words[index] = "Core";
                }
            }
            else if(winverNum == "7")
            {
                if(licenseType.Contains("VL"))
                {
                    outWords.Add("Volume");
                    return outWords;
                }
            }

            outWords.Add(words[index]);
            index++;

            if(index < words.Length)
            {
                outWords.Add(words[index]);
                index++;

                if (index < words.Length)
                {
                    outWords.Add(words[index]);
                    index++;

                }
            }

            return outWords;
        }

        public static string GetWinVer()
        {
            string output = Common.RunCommand("wmic", "os get * /value", sysDir);
            string winVer = Common.GetValueOfParamFromString(output, "Caption", '=');
            return winVer;
        }

        public static string GetProductID()
        {
            string output = Common.RunCommand("wmic", "os get SerialNumber /value", sysDir);
            string productId = Common.GetValueOfParamFromString(output, "SerialNumber", '=');

            return productId;
        }

        public static string GetLicenseTypeAndLicenseStatus()
        {
            string output = Common.RunCommand("cscript.exe", "slmgr.vbs /dli", sysDir);
            string status = Common.GetValueOfParamFromString(output, "License Status", ':');
            string type = "";
            string description = Common.GetValueOfParamFromString(output, "Description", ':').ToUpper();
            if (description.Contains("RETAIL"))
            {
                type = "Retail";
            }
            else if (description.Contains("VL") || description.Contains("VOLUME") || description.Contains("MAK"))
            {
                type = "VL";
            }
            else if (description.Contains("OEM"))
            {
                type = "OEM";
            }
            return status + "," + type;
        }
        static int TryActiveByKeyType(string keyType)
        {
            List<string> activeKeys = new List<string>();
            try
            {
                activeKeys = Common.getActiveKeysFromServer(keyType);
            }
            catch (Exception e)
            {
                if (e.Message == ""){
                    Console.WriteLine("Connect to server fail!");
                }
                else{
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

                Console.WriteLine("Applying key " + Common.getLastFiveOfKey(productKey) + "...");

                if(keyType.Contains("7") || keyType.Contains("2008")) //Activate Win 7
                {
                    int res = TryActiveWin7ByUsingKey(productKey);
                    if (res < 0)
                    {
                        break;
                    }
                    if (res > 0)
                    {
                        return 1;
                    }
                }
                else
                {
                    int res = TryActiveByUsingKey(productKey);
                    if (res < 0 && !keyType.Contains("EnterpriseS"))
                    {
                        break;
                    }
                    if (res > 0)
                    {
                        return 1;
                    }
                }
            }

            return 0;
        }

        static int TryActiveByUsingKey(string productKey)
        {
            string install_res = Common.RunCommand("cscript.exe", "slmgr.vbs /ipk " + productKey, sysDir);
            if (install_res.ToUpper().Contains("SUCCESS"))
            {
                string ato_res = Common.RunCommand("cscript.exe", "slmgr.vbs /ato", sysDir);
                if (ato_res.ToUpper().Contains("SUCCESS"))
                {
                    Console.WriteLine("Activation successfully!");
                    return 1;
                }
                else
                {
                    string errorCode = GetErrorCode(ato_res);
                    if (errorCode == "0xC004C008")
                    {
                        Console.WriteLine("Getting IID...");
                        string iid = GetIID();
                        Console.WriteLine(iid);
                        if (Common.IsIDValid(iid))
                        {
                            Console.WriteLine("Getting CID...");
                            string cid = Common.GetCID(iid);
                            Console.WriteLine(cid);
                            if (Common.IsIDValid(cid))
                            {
                                Common.RunCommand("cscript.exe", "slmgr.vbs /atp " + cid, sysDir);
                                ato_res = Common.RunCommand("cscript.exe", "slmgr.vbs /ato", sysDir);
                                if (ato_res.ToUpper().Contains("SUCCESS"))
                                {
                                    Console.WriteLine("Activation successfully!");
                                    return 1;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                return -1;
            }

            return 0;
        }

        static int TryActiveWin7ByUsingKey(string productKey)
        {
            string install_res = Common.RunCommand("cscript.exe", "slmgr.vbs /ipk " + productKey, sysDir);
            if (install_res.ToUpper().Contains("SUCCESS"))
            {
                string ato_res = Common.RunCommand("cscript.exe", "slmgr.vbs /ato", sysDir);
                if (ato_res.ToUpper().Contains("SUCCESS"))
                {
                    Console.WriteLine("Activation successfully!");
                    return 1;
                }
                else
                {
                    Console.WriteLine("Getting IID...");
                    string iid = GetIID();
                    Console.WriteLine(iid);
                    if (Common.IsIDValid(iid))
                    {
                        Console.WriteLine("Getting CID...");
                        string cid = Common.GetCID(iid);
                        Console.WriteLine(cid);
                        if (Common.IsIDValid(cid))
                        {
                            Common.RunCommand("cscript.exe", "slmgr.vbs /atp " + cid, sysDir);
                            Console.WriteLine("Activation successfully!");
                            return 1;
                        }
                    }
                }
            }
            else
            {
                return -1;
            }

            return 0;
        }

        static string GetIID()
        {
            string iid = "";

            string dti_res = Common.RunCommand("cscript.exe", "slmgr.vbs /dti", sysDir);
            string[] lines = dti_res.Split(new[] { "\r\n", "\r", "\n" },
            StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.ToUpper().Contains("ID:"))
                {
                    var subLines = line.Split(' ');
                    iid = subLines[2];
                    break;
                }
            }

            return iid;
        }
        static string GetErrorCode(string ato_res)
        {
            string res = "";
            string[] lines = ato_res.Split(new[] { "\r\n", "\r", "\n" },
            StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.ToUpper().Contains("ERROR:"))
                {
                    var subLines = line.Split(' ');
                    res = subLines[1];
                    break;
                }
            }
            return res;
        }
    }
}