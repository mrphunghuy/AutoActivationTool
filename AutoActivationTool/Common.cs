using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;
using System.Windows.Forms;
using System.Security.Cryptography;

namespace AutoActivationTool
{
    class Common
    {
        public static System.Windows.Forms.Label status;
        public static string m_username;
        public static string m_password;
        public static bool Login(string username, string password)
        {
            try
            {
                if (username == "" || password == "")
                {
                    status.Text = "Emtpy!";
                    status.ForeColor = System.Drawing.Color.Red;
                    Application.DoEvents();
                    return false;
                }

                status.Text = "Logging in...";
                status.ForeColor = System.Drawing.Color.Black;
                Application.DoEvents();

                var request = (HttpWebRequest)WebRequest.Create("https://winoffice.org/api/act-login");

                var postData = "username=" + Uri.EscapeDataString(username);
                postData += "&password=" + Uri.EscapeDataString(password);
                var data = Encoding.ASCII.GetBytes(postData);

                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                var response = (HttpWebResponse)request.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                responseString = responseString.Replace("\"", "");

                if (responseString == "1") //Login successfully
                {
                    m_username = username;
                    m_password = password;

                    return true;
                }
                else
                {
                    status.Text = "Invalid user name or password!";
                    status.ForeColor = System.Drawing.Color.Red;
                    Application.DoEvents();
                    return false;
                }
            }
            catch (Exception e)
            {
                status.Text = e.Message;
                status.ForeColor = System.Drawing.Color.Red;
                Application.DoEvents();
            }

            return false;
        }

        public static List<string> getKeyTypesFromServer()
        {
            List<string> output = new List<string>();

            Console.WriteLine("Getting key from server...");

            var request = (HttpWebRequest)WebRequest.Create("https://winoffice.org/api/act-get-key-types");

            var postData = "username=" + Uri.EscapeDataString(m_username);
            postData += "&password=" + Uri.EscapeDataString(m_password);
            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            responseString = responseString.Replace("\"", "");

            if (responseString == "-1")
            {
                throw new System.Exception("No Permission!");
            }

            string[] splitedStrings = responseString.Split(',');

            for (int i = 0; i < splitedStrings.Length; i++)
            {
                output.Add(splitedStrings[i]);
            }

            return output;
        }
        public static List<string> getActiveKeysFromServer(string keyType)
        {
            List<string> output = new List<string>();

            //Console.WriteLine("Getting key from server...");

            var request = (HttpWebRequest)WebRequest.Create("https://winoffice.org/api/act-get-active-keys");

            var postData = "username=" + Uri.EscapeDataString(m_username);
            postData += "&password=" + Uri.EscapeDataString(m_password);
            postData += "&table_name=" + Uri.EscapeDataString(keyType);
            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            responseString = responseString.Replace("\"", "");

            if (responseString == "-1")
            {
                throw new System.Exception("No Permission!");
            }

            string[] splitedStrings = responseString.Split(',');

            int keyNum = splitedStrings.Length / 3;
            for (int i = 0; i < keyNum; i++)
            {
                output.Add(splitedStrings[i * 3]);
            }

            return output;
        }

        public static bool isStartWith(string keyType, string startWith)
        {
            string keyTypeUpper = keyType.ToUpper();
            string startWithUpper = startWith.ToUpper();
            for (int i = 0; i < startWithUpper.Length; i++)
            {
                if (startWithUpper[i] != keyTypeUpper[i])
                {
                    return false;
                }
            }
            return true;
        }

        public static bool isContainsTags(string keyType, string tags)
        {
            string keyTypeUpper = keyType.ToUpper();

            string[] splitedStrings = tags.Split(',');
            for (int i = 0; i < splitedStrings.Length; i++)
            {
                string tagUpper = splitedStrings[i].ToUpper();
                if (!keyTypeUpper.Contains(tagUpper))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool isNotConntainsTags(string keyType, string notTags)
        {
            if (notTags == "")
            {
                return true;
            }

            string keyTypeUpper = keyType.ToUpper();

            string[] splitedStrings = notTags.Split(',');
            for (int i = 0; i < splitedStrings.Length; i++)
            {
                string tagUpper = splitedStrings[i].ToUpper();
                if (keyTypeUpper.Contains(tagUpper))
                {
                    return false;
                }
            }

            return true;
        }
        public static string GetCID(string iid)
        {
            string res = "";

            try
            {
                var request = (HttpWebRequest)WebRequest.Create("https://winoffice.org/api/act-get-cid");

                var postData = "username=" + Uri.EscapeDataString(m_username);
                postData += "&password=" + Uri.EscapeDataString(m_password);
                postData += "&iid=" + Uri.EscapeDataString(iid);
                var data = Encoding.ASCII.GetBytes(postData);

                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                var response = (HttpWebResponse)request.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                responseString = responseString.Replace("\"", "");

                bool success = true;


                if (success)
                {
                    res = responseString;
                }
            }
            catch (Exception e)
            {
                res = e.Message;
            }

            return res;
        }

        public static bool IsIDValid(string CID)
        {
            if(CID == ""){
                return false;
            }

            for(int i=0;i<CID.Length;i++)
            {
                if(CID[i] < '0' || CID[i] > '9')
                {
                    return false;
                }
            }

            return true;
        }

        public static string RunCommand(string command, string args, string workingDir = "")
        {
            //Console.WriteLine(command + " " + args);

            System.Diagnostics.Process pProcess = new System.Diagnostics.Process();

            //strCommand is path and file name of command to run
            pProcess.StartInfo.FileName = command;

            //strCommandParameters are parameters to pass to program
            pProcess.StartInfo.Arguments = args;

            pProcess.StartInfo.UseShellExecute = false;

            //Set output of program to be written to process output stream
            pProcess.StartInfo.RedirectStandardOutput = true;

            //Optional
            pProcess.StartInfo.WorkingDirectory = workingDir;

            pProcess.StartInfo.CreateNoWindow = true;

            //Start the process
            pProcess.Start();

            //Get program output
            string strOutput = pProcess.StandardOutput.ReadToEnd();

            //Wait for process to finish
            pProcess.WaitForExit();

            //Console.WriteLine(strOutput);
            pProcess.Close();

            return strOutput;
        }

        static readonly string PasswordHash = "P@@2242342340asdf";
        static readonly string SaltKey = "0jasf234234";
        static readonly string VIKey = "@asdf240234baf23";
        public static string Encrypt(string plainText)
        {
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
            var encryptor = symmetricKey.CreateEncryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));

            byte[] cipherTextBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                {
                    cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                    cryptoStream.FlushFinalBlock();
                    cipherTextBytes = memoryStream.ToArray();
                    cryptoStream.Close();
                }
                memoryStream.Close();
            }
            return Convert.ToBase64String(cipherTextBytes);
        }

        public static string Decrypt(string encryptedText)
        {
            byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
            byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

            var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));
            var memoryStream = new MemoryStream(cipherTextBytes);
            var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
        }

        static readonly string authenticationFileName = "logintoken.dat";
        public static void saveAccount(string username, string password)
        {
            string text = username + "\n" + password;
            string encryptedstring = Encrypt(text);
            File.WriteAllText(authenticationFileName, encryptedstring);
        }

        public static void loadAccount(TextBox userTextbox, TextBox passTextbox)
        {
            try
            {
                string text = File.ReadAllText(authenticationFileName);
                string decrytedstring = Decrypt(text);
                string[] splitedStrings = decrytedstring.Split('\n');
                userTextbox.Text = splitedStrings[0];
                passTextbox.Text = splitedStrings[1];
            }
            catch (Exception e) { }
        }

        public static string getLastFiveOfKey(string productKey)
        {
            try
            {
                string[] splitedString = productKey.Split('-');
                string lastFive = splitedString[4];

                return lastFive;
            }
            catch(Exception e) { };

            return "";
        }
        public static string GetValueOfParamFromString(string inString, string param, char splitChar)
        {
            string value = "";

            string[] lines = inString.Split('\n');
            foreach (string line in lines)
            {
                if (line.ToUpper().Contains(param.ToUpper()))
                {
                    value = line.Split(splitChar)[1];
                    value = value.Replace("\"", "");
                    value = value.Replace("\t", "");
                    value = value.Replace("\r", "");

                    value = RemoveAllStartEndSpaces(value);

                    break;
                }
            }

            return value;
        }

        public static string RemoveAllStartEndSpaces(string instr)
        {
            string outstr = "";

            int startIdx = 0;
            int endIdx = instr.Length - 1;
            while (instr[startIdx] == ' ') startIdx++;
            while (instr[endIdx] == ' ') endIdx--;

            outstr = instr.Substring(startIdx, endIdx - startIdx + 1);

            return outstr;
        }

        public static List<string> GetAllValueOfParamFromString(string inString, string param, char splitChar)
        {
            List<string> values = new List<string>();

            string[] lines = inString.Split('\n');
            foreach (string line in lines)
            {
                if (line.ToUpper().Contains(param.ToUpper()))
                {
                    string value = line.Split(splitChar)[1];
                    value = value.Replace(" ", "");
                    value = value.Replace("\"", "");
                    value = value.Replace("\t", "");
                    value = value.Replace("\r", "");

                    values.Add(value);
                }
            }

            return values;
        }
        public static bool IsKeyTypeConsist(string keyType, List<string> keyWords)
        {
            for (int i = 0; i < keyWords.Count; i++)
            {
                if (!keyType.Contains(keyWords[i]))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
