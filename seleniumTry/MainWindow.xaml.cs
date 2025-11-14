using FluentScheduler;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace SharabiProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool _isFinishInitializeComponent = false; // make sure can do change without crash in exampleMessageChange
        private bool isAutoMessageEnable = false;
        public MainWindow()
        {
            Console.WriteLine("loading...");
            InitializeComponent();

            //chromeOptionsText = "user-data-dir=C:\\Users\\שרה\\AppData\\Local\\Google\\Chrome\\User Data\\Default";
            chromeOptionsText = "user-data-dir=C:\\Users\\Sharabi\\AppData\\Local\\Google\\Chrome\\User Data\\Default";
            //chromeOptionsText = "user-data-dir=C:\\Users\\Lila-PC\\AppData\\Local\\Google\\Chrome\\User Data\\Default";
            //chromeOptionsText = "user-data-dir=C:\\Users\\yosis\\AppData\\Local\\Google\\Chrome\\User Data\\Default";
            //chromeOptionsText = "user-data-dir=C:\\Users\\gilad\\AppData\\Local\\Google\\Chrome\\User Data\\Default";
            grid1Property.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Star);
            grid2Property.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Star);
            settingButton.Content = "הגדרות";
            DataContext = _cropItems;
            //Console.WriteLine("main window");
            try
            {
                string filePath = "sharabiJson.json";
                string url = "";
                using (StreamReader sr = new StreamReader(filePath))
                {
                    var json = sr.ReadToEnd();
                    var jsonDic = JObject.Parse(json);
                    url = jsonDic["url"].ToString();

                }
                readCsv(url, true);
                urlTextBox.Text = url;
            }
            catch
            {
                Console.WriteLine("not read csv file at all");

            }
            // initialize hours to comboBox:
            for (int i = 1; i <= 24; i++)
            {
                ComboBoxItem item = new ComboBoxItem();
                ComboBoxItem item2 = new ComboBoxItem();
                item.Content = i.ToString() + ":00";
                item2.Content = i.ToString() + ":00";
                startTimeCB.Items.Add(item);
                finishTimeCB.Items.Add(item2);
            }
            exampleMsgToCustomer.Visibility = Visibility.Visible;
            ShowExampleMsgToCustomer();
            _isFinishInitializeComponent = true;


            // make sure can do change without crash in exampleMessageChange
            //int minuteToWait = 30;
            //try
            //{
            //    minuteToWait = Int32.Parse(timeToWaitTextBox.Text);
            //    errorTimeToWaitTextBox.Visibility = Visibility.Collapsed;
            //}
            //catch
            //{
            //    minuteToWait = 30;
            //    errorTimeToWaitTextBox.Visibility = Visibility.Visible;
            //}
            dateTime = DateTime.Now;
            timeToSendMsg = dateTime.AddMinutes(minuteToWait);
            //if ((int)dateTime.DayOfWeek == 2) // if tuesday

            JobManager.AddJob(
           this.DoScheduledWork,
        //schedule => schedule.ToRunEvery(3).Seconds());
        schedule => schedule.ToRunNow().AndEvery(minuteToWait).Minutes());

            JobManager.AddJob(
               this.ScheduledUpdateUrl,
            schedule => schedule.ToRunEvery(updateCsvEveryMinute).Minutes());

            Console.WriteLine("finish loading");
        }
        private DateTime dateTime;
        private DateTime timeToSendMsg;
        const int minuteToWait = 20;
        const int updateCsvEveryMinute = 6;

        private void ScheduledUpdateUrl()
        {
            dateTime = DateTime.Now;
            Dispatcher.Invoke(new Action(updateUrl));
        }

        private void updateUrl()
        {
            if (!seleniumProgresBar.IsVisible)
            {
                readCsv(urlTextBox.Text, true);
            }

        }
        private void DoScheduledWork()
        {
            if (isAutoMessageEnable)
            {
                DoPrimaryWorkOffUIThreed();
                Dispatcher.Invoke(new Action(ShowResultsOnUIThread));
            }
        }

        private void DoPrimaryWorkOffUIThreed()
        {
            dateTime = DateTime.Now;
        }

        private void ShowResultsOnUIThread() // send update msg to new custmer
        {
            dateTime = DateTime.Now;
            timeToSendMsg = dateTime.AddMinutes(minuteToWait);
            timeToSendTextBlock.Text = $"שליחה הבאה בשעה: {timeToSendMsg:HH:mm}";
            bool succeedLoadCsvFile;
            try
            {
                readCsv(urlTextBox.Text, true);
                succeedLoadCsvFile = true;
            }
            catch
            {
                updateMessageToStav.Text = $"לא הצליח לטעון קובץ בשעה {dateTime:HH:mm}";
                succeedLoadCsvFile = false;
            }
            if (succeedLoadCsvFile && newCustomers.Count > 0 && !seleniumProgresBar.IsVisible)
            {
                try
                {
                    seleniumSendMessage(newCustomers, false);
                }
                catch (Exception ex)
                {
                    // Console.WriteLine("the problem: " + ex);
                }
            }
        }
        private void enableWidgetsInThreed(bool isEnable)
        {
            if (!isEnable)
            {
                seleniumProgresBar.Visibility = Visibility.Visible;
            }
            else
            {
                seleniumProgresBar.Visibility = Visibility.Collapsed;
            }

            urlTextBox.IsEnabled = isEnable;
            updateCsvFile.IsEnabled = isEnable;
            msgToCustomers.IsEnabled = isEnable;
            sendNewOrdersNow.IsEnabled = isEnable;
            btnSendOrdersNow.IsEnabled = isEnable;
            btnSendMessageToAll.IsEnabled = isEnable;

            // close setting
            settingButton.IsEnabled = isEnable;
            grid1Property.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Star);
            grid2Property.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Star);
            settingButton.Content = "הגדרות";
        }

        private BackgroundWorker backgroundWorker;
        private string messageToStav = "";
        private void backgroundWorker_SeleniumWork(object sender, DoWorkEventArgs e)
        {
            try
            {

                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument(chromeOptionsText);
                //chromeOptions.AddAdditionalOption("detach", true);
                driver = new ChromeDriver(chromeOptions);
       
                
                driver.Navigate().GoToUrl("https://web.whatsapp.com");
            }
            catch
            {
                //new SeleniumControl().DownloadLatestVersionOfChromeDriver();
                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument(chromeOptionsText);

                //chromeOptions.AddAdditionalOption("detach", true);
                driver = new ChromeDriver(chromeOptions);
                driver.Navigate().GoToUrl("https://web.whatsapp.com");
            }

            Thread.Sleep(10000); // wait 10 sec 
            int waitIndex = 0;
            while (waitIndex < 600) // wait 5 minute to open whatsapp web. 
            { // if after 4 minute not active ask to restart app
                try
                {

                    driver.FindElement(By.XPath("//*[@id=\"side\"]/div[1]/div/div/button/div[2]/span")); // search for search buttom
                    waitIndex = 601; // mean it work properly
                }
                catch
                {
                    waitIndex++;
                    Thread.Sleep(5000);
                }
            }
            if (waitIndex == 600)
            {
                // don't get inside to the selenium
                messageToStav = "לא הצליח לפתוח את ווצאפ ווב. \n יש להפעיל מחדש את התוכנה";
                Console.WriteLine("please restart the app.");           
                System.Windows.Application.Current.Shutdown();            
            }
            else if (waitIndex == 601)
            {
                // get to selenium
                //Console.WriteLine("get to selenium");
                if (isTestMessage)
                {
                    new SeleniumControl().SendTestMessage(driver, seleniumExampleMessage);
                }
                else
                {
                    new SeleniumControl().SeleniumSendAllBackground(driver, listToSelenium, seleniumCropItems, seleniumMessage);
                }
                driver.Quit();
            }
        }
        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Console.WriteLine("in progress");
        }
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            { // אם זה התבטל באמצע
                Console.WriteLine("Operation Cancelled");

            }
            else if (e.Error != null)
            { // קרס
                Console.WriteLine("Error in Process :" + e.Error);
            }
            else
            { // אם זה באמת הצליח וסיים את הפעולה
                Console.WriteLine("Operation Completed " + e.Result);
                new SeleniumControl().SeleniumUpdateUiAfterSend();
                enableWidgetsInThreed(true);
            }
        }
        private List<Customer> listToSelenium;
        private bool isTestMessage;
        private string seleniumExampleMessage;
        private string seleniumMessage;
        private List<CropItem> seleniumCropItems;
        private void seleniumSendMessage(List<Customer> customersList, bool myIsTestMessage)
        {
            // initialazie element from UI
            seleniumExampleMessage = exampleMsgToCustomer.Text;
            TextRange textRange = new TextRange(msgToCustomers.Document.ContentStart, msgToCustomers.Document.ContentEnd);
            seleniumMessage = textRange.Text;
            listToSelenium = customersList;
            isTestMessage = myIsTestMessage;
            enableWidgetsInThreed(false);

            seleniumCropItems = new List<CropItem>(_cropItems);

            // enabled button
            backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            backgroundWorker.DoWork += backgroundWorker_SeleniumWork; //For the performing operation in the background.  
            backgroundWorker.ProgressChanged += backgroundWorker_ProgressChanged; //For the display of operation progress to UI.    
            backgroundWorker.RunWorkerCompleted += backgroundWorker_RunWorkerCompleted;  //After the completation of operation.  
            backgroundWorker.RunWorkerAsync("Press Enter in the next 5 seconds to Cancel operation:");
            updateMessageToStav.Text = messageToStav;
        }



        public string msgToCustomer(Customer customer)
        { // return the message to the specific customer
            TextRange textRange = new TextRange(msgToCustomers.Document.ContentStart, msgToCustomers.Document.ContentEnd);
            string message = textRange.Text;
            //message.Replace("\r\n", "\n");            
            int price = 0;
            string orderList = "";



            foreach (var item in _cropItems)
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
        private void ShowExampleMsgToCustomer()
        {
            Customer exampleCustomer;
            exampleCustomer = new Customer
            {
                PhoneNumber = "0545678726",
                FirstName = "סתיו",
                RowData = exampleRowData,
            };
            exampleMsgToCustomer.Text = msgToCustomer(exampleCustomer);
        }
        private void MsgToCustomers_TextBoxChanged(object sender, TextChangedEventArgs e)
        {            // every change in textBox make change in my example textBlock 
            if (_isFinishInitializeComponent) // else it crash..
            {
                ShowExampleMsgToCustomer();

            }
        }

        private void Btn_Update_Url(object sender, RoutedEventArgs e)
        {
            bool succeedReadCsv = readCsv(urlTextBox.Text, true);
            string filePath = "sharabiJson.json";
            string result = string.Empty;
            using (StreamReader sr = new StreamReader(filePath))
            {
                var json = sr.ReadToEnd();
                var jsonDic = JObject.Parse(json);
                //result = jsonDic.ToString();
                jsonDic["url"] = urlTextBox.Text;
                //Console.WriteLine(DictToString(cropPrices));                    
                result = jsonDic.ToString();
            }
            //JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(filePath, result, System.Text.Encoding.UTF8);
            ShowExampleMsgToCustomer();
            if (succeedReadCsv)
            {
                MessageBox.Show("הקובץ נטען בהצלחה");
            }
            else
            {
                MessageBox.Show("לא הצליח. יש לנסות שנית");
            }
        }

        private void BtnUpdateCropList_Click(object sender, RoutedEventArgs e)
        {


            // update local json file.
            string currentDir = System.IO.Directory.GetCurrentDirectory();
            //Console.WriteLine(currentDir);
            string filePath = "sharabiJson.json";
            string result = string.Empty;
            Dictionary<String, String> cropPrices;
            using (StreamReader sr = new StreamReader(filePath))
            {
                var json = sr.ReadToEnd();
                var jsonDic = JObject.Parse(json);
                //result = jsonDic.ToString();
                cropPrices = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonDic["cropPrices"].ToString());
                foreach (var item in _cropItems)
                {
                    try
                    {
                        cropPrices[item.CropName] = item.CropPrice.ToString();
                    }
                    catch (Exception)
                    {
                        cropPrices.Add(item.CropName, item.CropPrice.ToString());
                    }
                }
                //Console.WriteLine(DictToString(cropPrices));
                jsonDic["cropPrices"] = JsonConvert.SerializeObject(cropPrices);
                result = jsonDic.ToString();
            }
            //JsonConvert.SerializeObject(result, Formatting.Indented);
            File.WriteAllText(filePath, result, System.Text.Encoding.UTF8);
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            // update local json file
            if (listChooseHour.SelectedIndex == 0)
            {
                betweenHours.Visibility = Visibility.Visible;
            }
            else
            {
                betweenHours.Visibility = Visibility.Collapsed;
            }
        }

        private void AllCustomerList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                exampleMsgToCustomer.Text = msgToCustomer((Customer)allCustomersListBox.SelectedItem);
            }
            catch
            {
                ShowExampleMsgToCustomer();
            }
        }

        private void NewOrdersListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                exampleMsgToCustomer.Text = msgToCustomer((Customer)newOrdersListBox.SelectedItem);
            }
            catch
            {
                ShowExampleMsgToCustomer();
            }
        }
        

        private void NotInContactListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                exampleMsgToCustomer.Text = msgToCustomer((Customer)notInContactListBox.SelectedItem);
            }
            catch
            {
                ShowExampleMsgToCustomer();
            }
        }
        private void problemPhoneNumberListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                exampleMsgToCustomer.Text = msgToCustomer((Customer)problemPhoneNumberListBox.SelectedItem);
            }
            catch
            {
                ShowExampleMsgToCustomer();
            }
        }



        private void Btn_Send_Oreders_Now(object sender, RoutedEventArgs e)
        { // the final button to send the orders message
            try
            {
                readCsv(urlTextBox.Text, true);
                seleniumSendMessage(newCustomers, false);
            }
            catch
            {
                MessageBox.Show("לא הצליח לבצע שליחת הודעות");
            }
        }
        string chromeOptionsText;
        private void btnSendTestMessage(object sender, RoutedEventArgs e)
        {
            enableWidgetsInThreed(false);
            seleniumSendMessage(new List<Customer>(), true);

        }
        ChromeDriver driver;
        private void btnSendMessageToAll_Click(object sender, RoutedEventArgs e)
        {
            enableWidgetsInThreed(false);

            // asynce. after read csv open chrome.
            try
            {
                readCsv(urlTextBox.Text, true);
                seleniumSendMessage(allNewCsvCustomers, false);

            }
            catch
            {
                MessageBox.Show("לא הצליח לבצע שליחת הודעות");
            }

        }
        private void ClickSettingButton(object sender, RoutedEventArgs e)
        {         
            if (settingButton.Content == "הגדרות")
            {
                grid1Property.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Auto);
                grid1Property.UpdateLayout();
                grid2Property.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Auto);
                grid2Property.UpdateLayout();
                settingButton.Content = "סגור הגדרות";
            }
            else
            {
                grid1Property.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Star);
                grid2Property.ColumnDefinitions[0].Width = new GridLength(0, GridUnitType.Star);
                settingButton.Content = "הגדרות";
            }
        }

        private void btnCopyMsg_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(exampleMsgToCustomer.Text);
        }

        private void sliderTimeValueCahange(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            string text = $"השליחה תתבצע כל {sliderTime.Value.ToString()} דקות";
            timeToWaitTextBlock.Text = text;
        }
        private void btnAutoMsg_Click(object sender, RoutedEventArgs e)
        {
            if (isAutoMessageEnable)
            {
                isAutoMessageEnable = false;
                autoMsgBtn.Content = "כבוי";
                timeToSendTextBlock.Text = "שליחה אוטומטית לא פעילה";
            }
            else
            {
                isAutoMessageEnable = true;
                autoMsgBtn.Content = "פעיל";
                timeToSendTextBlock.Text = $"שליחה הבאה בשעה: {timeToSendMsg:HH:mm}";

            }
        }

        private void btnRemoveFromNewOrders_Click(object sender, RoutedEventArgs e)
        {
            int customerIdx = -1;
            customerIdx = newOrdersListBox.SelectedIndex;
            if (customerIdx > -1)
            {
                string filePath = "sharabiJson.json";
                var jsonDic = new JObject();
                using (StreamReader sr = new StreamReader(filePath))
                {
                    var json = sr.ReadToEnd();
                    jsonDic = JObject.Parse(json);
                }

                // update DataBase:
                List<Customer> sendedCustomerList = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["sendedCustomers"].ToString());
                sendedCustomerList.Add((Customer)newOrdersListBox.SelectedItem);
                jsonDic["sendedCustomers"] = JsonConvert.SerializeObject(sendedCustomerList);
                File.WriteAllText(filePath, jsonDic.ToString(), Encoding.UTF8);

                // update UI:
                newCustomers.RemoveAt(customerIdx);
                newOrdersListBox.ItemsSource = null;
                newOrdersListBox.ItemsSource = newCustomers;
                newOrdersListBox.SelectedIndex = customerIdx;
            }
            else
            {
                MessageBox.Show("לא סומן לקוח מרשימת הזמנות חדשות");
            }
        }
        private void btnRemoveAllNewOrders_Click(object sender, RoutedEventArgs e)
        {
            //resetNewCustomers();
            string filePath = "sharabiJson.json";
            var jsonDic = new JObject();
            var result = "";
            using (StreamReader sr = new StreamReader(filePath))
            {
                sendedCustomers.Clear(); // make it empty list                
                var json = sr.ReadToEnd();
                jsonDic = JObject.Parse(json);
                jsonDic["sendedCustomers"] = JsonConvert.SerializeObject(sendedCustomers);             
                jsonDic["notInContactCustomers"] = JsonConvert.SerializeObject(sendedCustomers); // because it's empty list
                newCustomers.Clear();

                //jsonDic["allCustomers"] = JsonConvert.SerializeObject(allNewCsvCustomers);
                newCustomers = JsonConvert.DeserializeObject<List<Customer>>(jsonDic["prevAllCustomers"].ToString()); ;
                result = jsonDic.ToString();

            }
            File.WriteAllText(filePath, result, Encoding.UTF8);
            newOrdersListBox.ItemsSource = null;
            newOrdersListBox.ItemsSource = newCustomers;
            MessageBox.Show("הלקוחות החדשים אופסו");
        }

}
}
