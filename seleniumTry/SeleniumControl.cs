using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;
using System.Xml.Linq;

namespace SharabiProject
{
    public class SeleniumControl
    {
        public void DownloadLatestVersionOfChromeDriver()
        {
            string path = DownloadLatestVersionOfChromeDriverGetVersionPath();
            var version = DownloadLatestVersionOfChromeDriverGetChromeVersion(path);
            var urlToDownload = DownloadLatestVersionOfChromeDriverGetURLToDownload(version);
            DownloadLatestVersionOfChromeDriverKillAllChromeDriverProcesses();
            DownloadLatestVersionOfChromeDriverDownloadNewVersionOfChrome(urlToDownload);
        }

        public string DownloadLatestVersionOfChromeDriverGetVersionPath()
        {
            //Path originates from here: https://chromedriver.chromium.org/downloads/version-selection            
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe"))
            {
                if (key != null)
                {
                    Object o = key.GetValue("");
                    if (!String.IsNullOrEmpty(o.ToString()))
                    {
                        return o.ToString();
                    }
                    else
                    {
                        throw new ArgumentException("Unable to get version because chrome registry value was null");
                    }
                }
                else
                {
                    throw new ArgumentException("Unable to get version because chrome registry path was null");
                }
            }
        }

        public string DownloadLatestVersionOfChromeDriverGetChromeVersion(string productVersionPath)
        {
            if (String.IsNullOrEmpty(productVersionPath))
            {
                throw new ArgumentException("Unable to get version because path is empty");
            }

            if (!File.Exists(productVersionPath))
            {
                throw new FileNotFoundException("Unable to get version because path specifies a file that does not exists");
            }

            var versionInfo = FileVersionInfo.GetVersionInfo(productVersionPath);
            if (versionInfo != null && !String.IsNullOrEmpty(versionInfo.FileVersion))
            {
                return versionInfo.FileVersion;
            }
            else
            {
                throw new ArgumentException("Unable to get version from path because the version is either null or empty: " + productVersionPath);
            }
        }

        public string DownloadLatestVersionOfChromeDriverGetURLToDownload(string version)
        {
            if (String.IsNullOrEmpty(version))
            {
                throw new ArgumentException("Unable to get url because version is empty");
            }

            //URL's originates from here: https://chromedriver.chromium.org/downloads/version-selection
            string html = string.Empty;
            string urlToPathLocation = @"https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_116" + String.Join(".", version.Split('.').Take(3));

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlToPathLocation);
            request.AutomaticDecompression = DecompressionMethods.GZip;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())            
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                html = reader.ReadToEnd();
            }

            if (String.IsNullOrEmpty(html))
            {
                throw new WebException("Unable to get version path from website");
            }

            return "https://chromedriver.storage.googleapis.com/" + html + "/chromedriver_win32.zip";
        }

        public void DownloadLatestVersionOfChromeDriverKillAllChromeDriverProcesses()
        {
            //It's important to kill all processes before attempting to replace the chrome driver, because if you do not you may still have file locks left over
            var processes = Process.GetProcessesByName("chromedriver");
            foreach (var process in processes)
            {
                try
                {
                    process.Kill();
                }
                catch
                {
                    //We do our best here but if another user account is running the chrome driver we may not be able to kill it unless we run from a elevated user account + various other reasons we don't care about
                }
            }
        }

        public void DownloadLatestVersionOfChromeDriverDownloadNewVersionOfChrome(string urlToDownload)
        {
            if (String.IsNullOrEmpty(urlToDownload))
            {
                throw new ArgumentException("Unable to get url because urlToDownload is empty");
            }

            //Downloaded files always come as a zip, we need to do a bit of switching around to get everything in the right place
            using (var client = new WebClient())
            {
                if (File.Exists(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.zip"))
                {
                    File.Delete(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.zip");
                }

                client.DownloadFile(urlToDownload, "chromedriver.zip");

                if (File.Exists(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.zip") && File.Exists(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.exe"))
                {
                    File.Delete(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.exe");
                }

                if (File.Exists(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.zip"))
                {
                    System.IO.Compression.ZipFile.ExtractToDirectory(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\chromedriver.zip", System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location));
                }
            }
        }


        private Tuple<bool, string> CorrectPhoneNumber(string phoneNumber)
        {
            bool isPhoneNumberCorrect = true;
            phoneNumber = phoneNumber.Trim();
            phoneNumber = phoneNumber.Replace(" ", "");
            phoneNumber = phoneNumber.Replace("-", "");
            if (phoneNumber.Length > 1)
            {
                if (phoneNumber[0] == '0')
                    phoneNumber = phoneNumber.Substring(1);
                if (phoneNumber[0] != '5')
                    isPhoneNumberCorrect = false;
                if (phoneNumber.Length != 9)
                    isPhoneNumberCorrect = false;
                //phoneNumber += $"+972{phoneNumber}";
                //phoneNumber = $"0{phoneNumber.Substring(0, 2)}-{phoneNumber.Substring(2)}";
            }
            else
            {
                isPhoneNumberCorrect = false;
            }

            return Tuple.Create(isPhoneNumberCorrect, phoneNumber);
        }

        private string msgToCustomer(Customer customer, string message)
        { // return the message to the specific customer        
            //message.Replace("\r\n", "\n");            
            int price = 0;
            string orderList = "";
            foreach (var item in cropItemList)
            {
                int cropQuantity = 0;
                if (customer.RowData[item.CropIndex] != "")
                {
                    cropQuantity = Int32.Parse(customer.RowData[item.CropIndex]);
                }
                if (item.IsIncludeInMessage && cropQuantity != 0)
                {
                    price += cropQuantity * item.CropPrice;
                    orderList += (cropQuantity + " " + item.CropName + "\r\n");
                }
            }
            message = message.Replace("נמען", customer.FirstName);
            message = message.Replace("$", price + " ₪");
            message = message.Replace("#", orderList);
            return message;
        }

        private List<CropItem> cropItemList;

        public void SeleniumSendAllBackground(ChromeDriver driver, List<Customer> customers, List<CropItem> myCropItems, string message)
        {
            cropItemList = myCropItems;
            //List<int> idxSendedCustomer;
            List<Customer> sendedCustomerList;
            List<Customer> notContactList;
            List<Customer> problemNumberCustomerList;
            string filePath = "sharabiJson.json";
            var jsonDic = new JObject();
            using (StreamReader sr = new StreamReader(filePath))
            {
                var json = sr.ReadToEnd();
                jsonDic = JObject.Parse(json);
            }
            //idxSendedCustomer = new List<int>();
            sendedCustomerList = new List<Customer>();
            notContactList = new List<Customer>();
            problemNumberCustomerList = new List<Customer>();

            int index = 0;
            IWebElement searchButton = driver.FindElement(By.XPath(                
                "//*[@id=\"side\"]/div[1]/div/div/button/div[2]/span"
                ));
            IWebElement labelSearch = driver.FindElement(By.XPath               
            ("//*[@id=\"side\"]/div[1]/div/div[2]/div/div/div/p"));

              //*[@id=\"side\"]/div[1]/div/div[2]/div/div/div/p

            //Console.WriteLine(  "finish 1 before loop");
            foreach (Customer customer in customers)
            {
                index++;
                searchButton.Click();
  
                labelSearch = driver.FindElement(By.XPath
            ("//*[@id=\"side\"]/div[1]/div/div[2]/div/div/div/p")); // not in the dom -> need find it again

                labelSearch.Click();

                string phoneNumber = customer.PhoneNumber;
                Tuple<bool, string> correctPhoneNumber = CorrectPhoneNumber(phoneNumber);
                //Console.WriteLine(correctPhoneNumber.Item1);
                if (correctPhoneNumber.Item1) // if phone number correct 
                {
                    phoneNumber = correctPhoneNumber.Item2;
                    phoneNumber = "+972" + phoneNumber.Replace("-", "");
                } // else put the original phone number insert by customer
                else
                {
                    problemNumberCustomerList.Add(customer);
                    continue;
                }

                Thread.Sleep(2000);
                labelSearch.Clear();
                labelSearch.SendKeys(phoneNumber);
                while(labelSearch.Text.ToString() != phoneNumber.ToString())
                { 
                    Console.WriteLine(labelSearch.Text.ToString());
                    Console.WriteLine(phoneNumber.ToString());
                    Thread.Sleep(2000);
                    labelSearch.Clear();
                    labelSearch.SendKeys(phoneNumber);
                }
                Thread.Sleep(3000);

                IWebElement accountToClick = searchButton; // initialize something not important. (just initialize. It change later)                             

                string xpathUserSearch = "//*[@id=\"pane-side\"]/div[1]/div/div/div[{0}]";
                int count = 0;
                int isExitLoop = 0;
                string headText = "";
                Thread.Sleep(4000);
                // search the customer and click:
                while (isExitLoop != 2)
                {
                    count++;
                    try
                    {
                        //Console.WriteLine(count);
                        string thisXPath = string.Format(xpathUserSearch, count.ToString());
                        var detailToPrint = driver.FindElement(By.XPath(thisXPath));
                        string styleOfElement = detailToPrint.GetAttribute("style").ToString();
                        //Console.WriteLine($"style of element: {styleOfElement}");
                        //int numStyleForAcoount = styleOfElement.Split("72px").Length - 1; // this function not work in .net 48
                        int numStyleForAcoount = 0;
                        string lastText = styleOfElement;
                        while (lastText.Contains("72px"))
                        {
                            lastText = lastText.Substring(lastText.IndexOf("72px") + 4);
                            //Console.WriteLine(lastText);
                            numStyleForAcoount++;
                        }
                        // if "72px is appear twice it mean that this is the right account to click


                        bool numStyleForHeader = styleOfElement.Contains("translateY(0px)");
                        if (numStyleForHeader)
                        {
                            headText = detailToPrint.Text;
                            isExitLoop++;
                        }
                        if (numStyleForAcoount == 2)
                        {
                            //Console.WriteLine("come in if");
                            isExitLoop++;
                            accountToClick = detailToPrint;
                        }
                        //Console.WriteLine("finish find customer");
                    }
                    catch (Exception)
                    {
                        //Console.WriteLine("not find customer");
                        isExitLoop = 2;
                        //notContactList.Add(customer);                                       
                        continue;
                        //throw;
                    }
                    //Console.WriteLine($"is exit loop = {isExitLoop}");
                }

                
                if (headText == "קבוצות" || headText == "הודעות" || headText == "")
                {
                    //Console.WriteLine("this number is not in contact" + phoneNumber);
                    notContactList.Add(customer);
                    continue;
                }
                try
                {
                    accountToClick.Click();
                } catch (Exception ex)
                {
                    Console.WriteLine($"show to Gilad the error: {ex}");
                    notContactList.Add(customer);
                    continue;
                }
                Thread.Sleep(3000);
                IWebElement editBox = driver.FindElement(By.XPath(                    
                    "//*[@id=\"main\"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p"
                  ));

                editBox.Clear();
                string[] listOfMessage = msgToCustomer(customer, message).Split('\n');
                foreach (var text in listOfMessage)
                {
                    editBox.SendKeys(text);
                    editBox.SendKeys(Keys.Shift + Keys.Enter);
                }
                Thread.Sleep(2000);

                //IWebElement sendButton = driver.FindElement(By.XPath(
               //     "*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span"
                    //"//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button"

             //       ));

                //sendButton.Click();
                editBox.SendKeys(Keys.Enter); // the send button not work
                

                // add customer to Json database 
                sendedCustomerList.Add(customer);
                List<Customer> sendedCustomers = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["sendedCustomers"].ToString());               
                sendedCustomers.Add(customer);
                jsonDic["sendedCustomers"] = JsonConvert.SerializeObject(sendedCustomers);
                File.WriteAllText(filePath, jsonDic.ToString(), Encoding.UTF8);


                //idxSendedCustomer.Add(index);
                Thread.Sleep(2000);
                //Console.WriteLine(idxSendedCustomer.Count);           
            }
            jsonDic["notInContactCustomers"] = JsonConvert.SerializeObject(notContactList);
            //jsonDic["idxSendedCustomer"] = JsonConvert.SerializeObject(idxSendedCustomer);
            jsonDic["problemCustomerList"] = JsonConvert.SerializeObject(problemNumberCustomerList);
            File.WriteAllText(filePath, jsonDic.ToString(), Encoding.UTF8);
            Thread.Sleep(5000);
            driver.Quit();
        }

        public void SeleniumUpdateUiAfterSend()
        {
            string filePath = "sharabiJson.json";
            var jsonDic = new JObject();
            using (StreamReader sr = new StreamReader(filePath))
            {
                var json = sr.ReadToEnd();
                jsonDic = JObject.Parse(json);
            }
            List<int> idxSendedCustomer = JsonConvert.DeserializeObject<List<int>>(jsonDic["idxSendedCustomer"].ToString());
            List<Customer> sendedCustomerList = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["sendedCustomers"].ToString());
            List<Customer> notContactList = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["notInContactCustomers"].ToString());
            List<Customer> problemNumberCustomerList = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["problemCustomerList"].ToString());


            string messageToStav = "";
            messageToStav += $"הודעות נשלחו ל{idxSendedCustomer.Count} לקוחות.";
            messageToStav += "\n\n";
            if (notContactList.Count != 0)
            {
                messageToStav += "לקוחות שלא באנשי הקשר: \n";
                foreach (var customer1 in notContactList)
                {
                    string customerName = customer1.FirstName;
                    string customerlastName = customer1.LastName;
                    string customerPhoneNumber = customer1.PhoneNumber;
                    messageToStav += $"{customerName} {customerlastName} {customerPhoneNumber} \n";
                }

            }
            if (problemNumberCustomerList.Count != 0)
            {
                messageToStav += "לקוחות עם מספרי טלפון לא נכונים:";
                foreach (var customer1 in problemNumberCustomerList)
                {
                    string customerName = customer1.FirstName;
                    string customerlastName = customer1.FirstName;
                    string customerPhoneNumber = customer1.PhoneNumber;
                    messageToStav += $"{customerName} {customerlastName} {customerPhoneNumber} \n";
                }
            }
            MainWindow mainWindow = (MainWindow)Application.Current.MainWindow;
            mainWindow.updateMessageToStav.Text = messageToStav;
            mainWindow.notInContactListBox.ItemsSource = null;
            mainWindow.notInContactListBox.ItemsSource = notContactList;
            //mainWindow.problemPhoneNumberListBox.ItemsSource = null;
            //mainWindow.problemPhoneNumberListBox.ItemsSource = problemNumberCustomerList;
            MainWindow.newCustomers.Clear();
            mainWindow.newOrdersListBox.ItemsSource = MainWindow.newCustomers;
            
        }

        public void SendTestMessage(ChromeDriver driver, string exampleMsg)
        {
            //var driver = SetupChromeDriver();
         
            IWebElement searchButton = driver.FindElement(By.XPath(
                
                "//*[@id=\"side\"]/div[1]/div/div/button/div[2]/span"
                ));
            IWebElement labelSearch = driver.FindElement(By.XPath
                ("//*[@id=\"side\"]/div[1]/div/div/div[2]/div/div[2]"));
                  
            System.Threading.Thread.Sleep(1000);
            searchButton.Click();
            labelSearch.Clear();
            //labelSearch.SendKeys("+972549727227");
            labelSearch.SendKeys("+972545678726");
            System.Threading.Thread.Sleep(3000);
            labelSearch.SendKeys(Keys.Enter);
            System.Threading.Thread.Sleep(2000);
            IWebElement editBox = driver.FindElement(By.XPath(
                    //"//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]"

                    //"//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p"
                    "//*[@id=\"main\"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p"
                ));
            editBox.Clear();

            string[] listOfMessage = exampleMsg.Split('\n');
            foreach (var text in listOfMessage)
            {
                editBox.SendKeys(text);
                editBox.SendKeys(Keys.Shift + Keys.Enter);
            }
            Thread.Sleep(2000);
            //IWebElement sendButton = driver.FindElement(By.XPath(
            //     "//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button"));
            //Thread.Sleep(500);
            //sendButton.Click();

            // //*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button
            editBox.SendKeys(Keys.Enter);
            editBox.SendKeys(Keys.Enter);
            Thread.Sleep(5000);
            driver.Quit();
        }
    }
}
