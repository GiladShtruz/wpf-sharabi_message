using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using System.Text;
using System.Windows;

namespace SharabiProject
{
    public partial class MainWindow
    {
        static string[] exampleRowData;
        public static List<Customer> allNewCsvCustomers = new List<Customer>();
        public static List<Customer> prevCustomers = new List<Customer>();
        public static List<Customer> newCustomers = new List<Customer>();
        public static List<Customer> problemNumberCustomers = new List<Customer>();
        public static List<Customer> sendedCustomers = new List<Customer>();
        public static ObservableCollection<CropItem> _cropItems = new ObservableCollection<CropItem>();

        public static Tuple<bool, string> CorrectPhoneNumber(string phoneNumber)
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
        public static void resetNewCustomers()
        {
            
            

        }

        public static bool readCsv(string url, bool isFromUrl)
        {
            // url to share
            //Console.WriteLine("get readCsv def");
            bool succeedReadCsv = false;
            // https://docs.google.com/spreadsheets/d/1tTqDOhfkkdrUFxs_OgEVMElw2zOotDVD3NutWRH2jH4/edit?usp=sharing
            //_cropItems.Clear();
            //mainWindow.cropListBox.ItemsSource = null;
            //WebClient webClient = new WebClient();
            //String csvFromWeb = webClient.DownloadString("https://docs.google.com/spreadsheets/d/1TkTqz1A3wBFigI5EsUcFw8DKpFuVxl4hh6PA421v56A/export?format=csv");

            /// from here its work:
            //HttpWebRequest req = (HttpWebRequest)WebRequest.Create("https://docs.google.com/spreadsheets/d/1TkTqz1A3wBFigI5EsUcFw8DKpFuVxl4hh6PA421v56A/export?format=csv");
            //HttpWebResponse resp = (HttpWebResponse)req.GetResponse();

            //StreamReader sr = new StreamReader(resp.GetResponseStream());
            StreamReader csvSr; // it can read from web and from local file
            if (isFromUrl)
            {
                try
                {
                    url = url.Trim();
                    url = url.Replace("edit?usp=sharing", "export?format=csv");
                    //WebClient webClient = new WebClient();
                    //String csvFromWeb = webClient.DownloadString(url);
                    ////Console.WriteLine(result); 
                    //string csvDir = System.IO.Directory.GetCurrentDirectory() + @"\orders.csv";
                    //Console.WriteLine(currentDir);        
                    //File.WriteAllText(csvDir, csvFromWeb, Encoding.GetEncoding("Windows-1255"));
                    //csvSr = new StreamReader("orders.csv");

                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url); // get to url
                    HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
                    csvSr = new StreamReader(resp.GetResponseStream()); // csv reader
                    succeedReadCsv = true;
                    //Console.WriteLine("succeedd read csv");
                }
                catch
                {
                    Console.WriteLine("Fail in read google sheet form");
                    succeedReadCsv = false;
                    csvSr = new StreamReader("orders.csv");
                }
            }
            else
            {
                csvSr = new StreamReader("orders.csv");
            }

            //StreamReader sr = new StreamReader(@"C:\Users\Lila-PC\VisualStudio\firstConsoleApp\firstConsoleApp\orders.csv");
            if (succeedReadCsv)
            {
                allNewCsvCustomers.Clear();
                prevCustomers.Clear();
                newCustomers.Clear();
                problemNumberCustomers.Clear();
                sendedCustomers.Clear();
                MainWindow mainWindow = (MainWindow)Application.Current.MainWindow;
                mainWindow.allCustomersListBox.ItemsSource = null;
                mainWindow.newOrdersListBox.ItemsSource = null;
                //mainWindow.notInContactListBox.ItemsSource = null;
                mainWindow.problemPhoneNumberListBox.ItemsSource = null;
                char csvSplitBy = '\u002C'; // represent comma (",")
                string row = csvSr.ReadLine(); // read the first line of csv file
                string[] rowData = row.Split(csvSplitBy); // split the first line to items (בזיל, פטרוזיליה..)
                exampleRowData = rowData; // for example rowData in show message (סתיו)
                int cropIndex = 0; // index of first row csv
                List<string> notCropsStrings = new List<string> { // list of category to not include int the list ui
                                    "חותמת זמן",
                                    "שם פרטי",
                                    "שם משפחה",
                                    "מספר טלפון",
                                    "הערות"
                                    };
                int firstNameColumnIdx = 0; // this initialize for not execpted
                int phoneNumColumnIdx = 0; // this initialize for not execpted
                int lastNameColumnIdx = 0; // this initialize for not execpted
                int timeColumnIdx = 0; // this initialize for not execpted
                _cropItems.Clear();


                // read json file:
                string currentDir = System.IO.Directory.GetCurrentDirectory();
                string filePath = "sharabiJson.json";
                Dictionary<String, String> cropPrices;
                var jsonDic = new JObject();
                using (StreamReader sr = new StreamReader(filePath))
                {
                    var json = sr.ReadToEnd();
                    jsonDic = JObject.Parse(json);
                    cropPrices = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonDic["cropPrices"].ToString());
                }

                foreach (string cropItem in rowData)
                {
                    if (cropItem.Trim() == "שם פרטי")
                        firstNameColumnIdx = cropIndex;
                    if (cropItem.Trim() == "מספר טלפון")
                        phoneNumColumnIdx = cropIndex;
                    if (cropItem.Trim() == "שם משפחה")
                        lastNameColumnIdx = cropIndex;
                    if (cropItem.Trim() == "חותמת זמן")
                        timeColumnIdx = cropIndex;

                    if (!notCropsStrings.Contains(cropItem.Trim()) && cropItem != "")
                    { // mean it in crop item and can counter quantity
                      // clean the cropName string:
                        char[] separators = new char[] { '[', ']', ',', '(', ')', '.' };
                        string currentCropItem = cropItem;
                        string[] temp = currentCropItem.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        currentCropItem = String.Join("\n", temp);
                        currentCropItem = currentCropItem.Trim();
                        // add item:
                        int cropPrice = 0;
                        try
                        {
                            cropPrice = Convert.ToInt32(cropPrices[currentCropItem]);
                        }
                        catch
                        {
                            cropPrices.Add(currentCropItem, "0");
                        }
                        _cropItems.Add(new CropItem
                        {
                            CropName = currentCropItem,
                            CropPrice = cropPrice,
                            CropIndex = cropIndex,
                            IsIncludeInMessage = true,
                        });
                        exampleRowData[cropIndex] = "1"; // for example msg to customer that have one item from each crop
                    }
                    cropIndex++;
                }

                // from here it read the continue csv. the customer orders:
                row = csvSr.ReadLine();
                if (row == null) // (it can two condition:  "" or null.
                    rowData[0] = ""; // don't get into the loop
                else
                    rowData = row.Split(csvSplitBy);

                int index = 0;
                //Console.WriteLine(rowData[0]);
                while (rowData[0] != "") // here it stop when the csv finish.
                { // (it can two condition:  "" or null.

                    string phoneNumber = rowData[phoneNumColumnIdx].ToString();
                    Tuple<bool, string> correctPhoneNumber = CorrectPhoneNumber(phoneNumber);
                    if (correctPhoneNumber.Item1) // if phone number correct 
                    {
                        phoneNumber = correctPhoneNumber.Item2;
                        phoneNumber = $"0{phoneNumber.Substring(0, 2)}-{phoneNumber.Substring(2)}";
                    } // else it put the original phone number insert by customer


                    Customer customer = new Customer
                    {
                        FirstName = rowData[firstNameColumnIdx].ToString(),
                        LastName = rowData[lastNameColumnIdx].ToString(),
                        PhoneNumber = phoneNumber,
                        Time = rowData[timeColumnIdx].ToString(),
                        RowData = rowData,
                        CsvRowIndex = index,
                    };

                    if (correctPhoneNumber.Item1)
                        allNewCsvCustomers.Add(customer);
                    else
                    {
                        problemNumberCustomers.Add(customer);
                        //allNewCsvCustomers.Add(customer);
                    }


                    row = csvSr.ReadLine();
                    if (row == null)
                        rowData[0] = ""; // break loop
                    else
                        rowData = row.Split(csvSplitBy);

                    index++;
                }

                // check the new orders:
                sendedCustomers = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["sendedCustomers"].ToString());
                List<Customer> jsonNotInContact = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["notInContactCustomers"].ToString());
                /*
                if (sendedCustomers.Count > allNewCsvCustomers.Count)
                {
                    sendedCustomers.Clear(); // make it empty list                
                    jsonDic["sendedCustomers"] = JsonConvert.SerializeObject(sendedCustomers);
                    jsonDic["notInContactCustomers"] = JsonConvert.SerializeObject(sendedCustomers); // because it's empty list
                }
                */
                foreach (var customer in allNewCsvCustomers)
                {
                    bool isExist = false;
                    foreach (var sendedCustomer in sendedCustomers)
                    {
                        if (customer.Time == sendedCustomer.Time && customer.LastName == sendedCustomer.LastName)
                        {
                            isExist = true;
                        }
                    }
                    if (!isExist)
                    {
                        newCustomers.Add(customer);
                    }
                }

                mainWindow.allCustomersListBox.ItemsSource = allNewCsvCustomers;
                mainWindow.problemPhoneNumberListBox.ItemsSource = problemNumberCustomers;
                mainWindow.newOrdersListBox.ItemsSource = newCustomers;
                jsonDic["prevAllCustomers"] = JsonConvert.SerializeObject(allNewCsvCustomers);
                // write to text all customers
                File.WriteAllText(filePath, jsonDic.ToString(), Encoding.UTF8);
            }            
            
            csvSr.Close();
            return succeedReadCsv;

        }
    }
}
