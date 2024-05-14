using AutoItX3Lib;
using AutomationUAT.Utility;
using AutomationUAT.ViewModels;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System.Configuration;
using System.Data;
using System.Text;
using TestAutomationSuite.dbModels;


namespace TestAutomationSuite.Utility
{

    public class ActionHelper
    {

        string browser = null;
        string driverPath = null;
        string LogPath = null;
        string implicitwait = null;
        string ScreenshotPath = null;

        public ActionHelper()
        {
            LogPath = ConfigurationManager.AppSettings["LogPath"];
            driverPath = ConfigurationManager.AppSettings["driverPath"];
            ScreenshotPath = ConfigurationManager.AppSettings["ScreenshotPath"];
            implicitwait = ConfigurationManager.AppSettings["implicitwait"];
        }

        //All ACTIONS METHODS
        #region QMS Actions 
        //Method to perform action that Navigates to the URL 

        private void SetDriverTimeouts(IWebDriver driver)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(20);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(int.Parse(implicitwait));
        }
        private void logging(string messege)
        {
            LoggingHelper loggingHelper = new LoggingHelper();
            loggingHelper.insertLog(messege);
        }
        private By GetBy(string searchBy, string elementValue, IWebDriver driver)
        {
            SetDriverTimeouts(driver);
            searchBy = searchBy.ToLower().Trim();
            By by = null;
            switch (searchBy)
            {
                case "id":
                    by = By.Id(elementValue);
                    break;
                case "xpath":
                    by = By.XPath(elementValue);
                    break;
                case "classname":
                    by = By.ClassName(elementValue);
                    break;
                default:
                    break;
            }
            return by;
        }

        public bool ActionURL(IWebDriver driver, string browser, string driverPath, string url, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                driver.Navigate().GoToUrl(url);
                if (driver.Manage().Window.Position.X != 0 && driver.Manage().Window.Position.Y != 0) driver.Manage().Window.Maximize();
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionClickRadioButton(string textValue, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";

            try
            {
                switch (textValue)
                {
                    case "Life":
                        IWebElement lifeRadioButton = driver.FindElement(By.XPath("(//input[@type='radio' = '" + textValue + "'])[1]"));
                        lifeRadioButton.Click();
                        break;

                    case "Non-life":
                        IWebElement nonLifeRadioButton = driver.FindElement(By.XPath("(//input[@type='radio' = '" + textValue + "'])[2]"));
                        nonLifeRadioButton.Click();
                        break;

                    default:
                        // Handle default case or raise an exception if needed
                        status = false;
                        exception = "Invalid radio button option: " + textValue;
                        break;
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionQuitBrowser(FileModels fileInfoP, out string exception, IWebDriver driver, out bool driverStatus)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            driverStatus = true;
            try
            {
                FileModels fileInfo = fileInfoP;
                fileInfo.endTime = DateTime.Now.ToString();
                driver.Manage().Cookies.DeleteAllCookies();
                driver.Quit();
                driverStatus = false;
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionCloseTab(out string exception, IWebDriver driver, out bool driverStatus)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            driverStatus = true;
            try
            {
                driver.Close();
                if (driver == null) driverStatus = false;
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionForceClick(string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                By by = GetBy(searchBy, elementValue, driver);
                if (by != null)
                {
                    IWebElement element = driver.FindElement(by);
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
                    if (!element.Selected)
                    {
                        js.ExecuteScript("arguments[0].click();", element);
                    }
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }
        
        int dropdownIndex = 0;
        int inputFieldIndex = 0;
        int inputDescCount = 0;

        public bool ActionEmployeeDetailsTable(IWebDriver driver, string records, string elementValue, int employeeIndex, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            if (employeeIndex == 0)
            {
                dropdownIndex = 0;
                inputFieldIndex = 0;
                inputDescCount = 0;
            }
            try
            {
                string expected = null;
                int recordIndex = 0;
                string[] recordArray = records.Split("|");
                IWebElement tableLocator = driver.FindElement(By.XPath(elementValue));
                IList<IWebElement> dropdownLocators = tableLocator.FindElements(By.XPath(".//p-dropdown[1]//div[1]//div[2]"));
                IList<IWebElement> inputLocators = tableLocator.FindElements(By.XPath(".//p-inputnumber[1]//span[1]/input[1]"));
                IList<IWebElement> inputDesc = tableLocator.FindElements(By.XPath(".//td//input[contains(@id, 'descriptionOfEmployees')]"));
                while (dropdownIndex < dropdownLocators.Count)
                {
                    dropdownLocators[dropdownIndex].Click();
                    Thread.Sleep(1000);
                    expected = recordArray[recordIndex];
                    var element = driver.FindElements(By.CssSelector("ul.p-dropdown-items li.p-dropdown-item")).FirstOrDefault(e => e.Text == expected);
                    if (element != null)
                    {
                        element.Click();
                        dropdownIndex++;
                        recordIndex++;
                    }
                    else
                    {
                        logging($"Dropdown does not contains {expected} option, Please check your Input Column Value again");
                        status = false;
                        break;
                    }
                }
                while (inputFieldIndex < inputLocators.Count)
                {
                    inputLocators[inputFieldIndex].Clear();
                    inputLocators[inputFieldIndex].SendKeys(recordArray[recordIndex]);
                    recordIndex++;
                    inputFieldIndex++;
                }
                while (inputDescCount < inputDesc.Count)
                {
                    inputDesc[inputDescCount].Clear();
                    inputDesc[inputDescCount].SendKeys(recordArray[recordIndex]);
                    recordIndex++;
                    inputDescCount++;
                }
                
            }
            catch (Exception ex)
            {
                dropdownIndex = 0;
                inputFieldIndex = 0;
                inputDescCount = 0;
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionClick(string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                By by = GetBy(searchBy, elementValue, driver);
                if (by != null)
                {
                    IWebElement element = driver.FindElement(by);
                    if (!element.Selected)
                    {
                        Thread.Sleep(1000);
                        element.Click();
                    }
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionCLEAR(string SearchBy, string ElementValue, out string exception, IWebDriver driver)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            try
            {
                if (SearchBy.ToLower().Trim() == "id")
                {
                    driver.FindElement(By.Id(ElementValue)).Clear();
                    Thread.Sleep(500);
                }
                else if (SearchBy.ToLower().Trim() == "xpath")
                {
                    driver.FindElement(By.XPath(ElementValue)).Clear();
                    Thread.Sleep(500);
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionCoverClick(int index, string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                driver.FindElement(By.XPath($"(//*[@id='BSC_card_SumInsured_Split'])[{index}]")).Click();
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public void WriteDatatoExcel(string InputColumnName, string xlInputFilePath, string destinationFile, int indexRow, string xlInputSheetName, string write)
        {
            ExcelHelper excelHelper = new ExcelHelper();
            string filePath = string.IsNullOrEmpty(destinationFile) ? xlInputFilePath : destinationFile;
            DataTable InputData = excelHelper.GetDataTableFromExcelFile(filePath, xlInputSheetName);
            string header = GetColumnKeywordByHeaderName(InputData, InputColumnName);
            excelHelper.UpdateCellold(filePath, xlInputSheetName, write, (uint)indexRow, header);
        }

        public bool ActionTabledatainput(string ElementValue, string namedCovers, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                int index = 0;
                string[] value = namedCovers.Split("|");
                IWebElement cardElement = driver.FindElement(By.XPath(ElementValue));
                IList<IWebElement> inputFields = cardElement.FindElements(By.XPath(".//input")); // Replace with the appropriate locator                
                foreach (IWebElement inputField in inputFields)
                {
                    inputField.Clear();
                    inputField.SendKeys(value[index]);
                    if (index == value.Count() - 1)
                    {
                        break;
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionWait(IWebDriver driver, string inputvalue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                int waitMilliseconds = Convert.ToInt32(inputvalue);
                Thread.Sleep(waitMilliseconds);
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionText(IWebDriver driver, string searchBy, string value, string arr, string inputColumn, string globalQuoteId, string globalPremium, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";

            try
            {
                string textValue = (inputColumn == "QuoteIdInput") ? globalQuoteId : arr;
                if (!string.IsNullOrEmpty(inputColumn))
                {
                    By by = GetBy(searchBy, value, driver);
                    if (by != null)
                    {
                        IWebElement element = driver.FindElement(by);
                        element.Click();
                        element.SendKeys(textValue);
                    }
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionClearAndEnter(string inputColumn, string searchBy, IWebDriver driver, string elementValue, string textValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            if (!string.IsNullOrEmpty(inputColumn))
            {
                try
                {
                    By by = GetBy(searchBy, elementValue, driver);
                    if (by != null)
                    {
                        IWebElement element = driver.FindElement(by);
                        element.Clear();
                        element.Click();
                        element.SendKeys(textValue);
                    }
                }
                catch (Exception ex)
                {
                    logging(ex.ToString());
                    status = false;
                    exception = ex.ToString();
                }
            }
            else
            {
                Console.WriteLine("Input Column is Empty");
            }
            return status;
        }

        public bool ActionLaunchbrowser(String browser, String driverPaths, out string exception, out IWebDriver driver)
        {
            bool status = true;
            exception = null;
            driver = null;
            try
            {
                if (browser == "chrome")//Launching Browser
                {
                    logging(" Assigning Driver: Chrome Driver");
                    driver = new ChromeDriver(driverPath + "chromedriver.exe");
                    AreDriverAndBrowserVersionsSame(driver);
                }
                else if (browser == "microsoftedge")
                {
                    logging(" Assigning Driver: MicrosoftEdge Driver");
                    if (!File.Exists(driverPath + "msedgedriver.exe"))
                    {
                        logging("msedgedriver.exe is missing");
                        logging("*** BOT IS DEACTIVATED ***");
                        Environment.Exit(1);
                    }
                    else
                    {
                        Environment.SetEnvironmentVariable("webdriver.edge.driver", driverPath + "msedgedriver.exe");
                        driver = new EdgeDriver();
                        AreDriverAndBrowserVersionsSame(driver);
                    }
                }
                else if (browser == "firefox")
                {
                    if (!File.Exists(driverPath + "geckodriver.exe"))
                    {
                        logging("geckodriver.exe is missing");
                        logging("*** BOT IS DEACTIVATED ***");
                        Environment.Exit(1);
                    }
                    else
                    {
                        logging(" Assigning Driver: FireFox Driver");
                        string geckopath = driverPath + "geckodriver.exe";
                        System.Environment.SetEnvironmentVariable("webdriver.gecko.driver", geckopath);
                        driver = new FirefoxDriver();
                        AreDriverAndBrowserVersionsSame(driver);
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
                driver = null;
            }
            return status;
        }

        public bool ActionRead(ActionModel itemAction, DataTable dtData, int indexRow, ExcelHelper excelHelper, out string exception, IWebDriver driver, string xlInputFilePath, string destinationFile, string xlInputSheetName, out string globalPremium, out string globalQuoteId)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            globalPremium = "";
            globalQuoteId = "";
            try
            {
                string idstr = "";
                if (itemAction.SearchBy == "id")
                {
                    idstr = driver.FindElement(By.Id(itemAction.Value)).Text;
                }
                else if (itemAction.SearchBy == "xpath")
                {
                    idstr = driver.FindElement(By.XPath(itemAction.Value)).Text;

                    if (itemAction.InputColumnName == "Premium")
                    {
                        globalPremium = idstr;
                    }

                    if (itemAction.InputColumnName == "QuoteId")
                    {
                        globalQuoteId = idstr;
                    }
                }
                string header = GetColumnKeywordByHeaderName(dtData, itemAction.InputColumnName);
                string filePath = string.IsNullOrEmpty(destinationFile) ? xlInputFilePath : destinationFile;
                excelHelper.UpdateCellold(filePath, xlInputSheetName, idstr, (uint)indexRow, header);
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionScrollbar(string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";

            try
            {
                IWebElement element = null;
                if (searchBy.ToLower().Trim() == "id" || searchBy.ToLower().Trim() == "xpath")
                {
                    By by = GetBy(searchBy, elementValue, driver);
                    element = driver.FindElement(by);
                    if (by != null)
                    {
                        element = driver.FindElement(by);
                    }
                }
                if (element != null)
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].scrollIntoView();", element);
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.Message;
            }
            return status;
        }

        public bool ActionDraganddrop(string textValue, IWebDriver driver, string ElementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            try
            {
                string[] path = ElementValue.Split("<>");
                IWebElement sourceElement = driver.FindElement(By.XPath("//li[contains(text(),'" + textValue + "')]"));
                IWebElement targetElement = driver.FindElement(By.XPath(path[1]));

                // Check if the source element is already at the target position
                if (sourceElement.Location.Equals(targetElement.Location))
                {
                    // If the source element is already at the target position, skip the drag-and-drop action
                    return true;
                }

                Actions actions = new Actions(driver);
                actions.DragAndDrop(sourceElement, targetElement).Build().Perform();
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.Message;
            }
            return status;
        }

        public bool ActionSelect(ActionModel itemAction, string textValue, out string exception, IWebDriver driver)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                if (textValue != null)
                {
                    string path = itemAction.Value;
                    string selectType = path.Split("[")[0];
                    if (selectType == "//input")
                    {
                        string locator = $"{selectType}[@id='{textValue}']";
                        IWebElement element = driver.FindElement(By.XPath(locator));
                        IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                        jsExecutor.ExecuteScript("arguments[0].scrollIntoView()", element);
                        element.Click();
                    }
                    else
                    {
                        string locator = $"{selectType}[contains(text(),'{textValue}')]";
                        IWebElement element = driver.FindElement(By.XPath(locator));
                        IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                        jsExecutor.ExecuteScript("arguments[0].scrollIntoView()", element);
                        element.Click();
                    }
                }
                else if (string.IsNullOrEmpty(textValue))
                {
                    driver.FindElement(By.XPath(itemAction.Value)).Click();
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool IsModellocatoralert(IWebDriver driver)
        {
            try
            {
                driver.FindElement(By.ClassName("bootbox-body"));
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool IsAlertPresent(IWebDriver driver)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            bool status = false;
            try
            {
                IAlert alert = driver.SwitchTo().Alert();
                status = true;
            }
            catch (Exception ex)
            {
                status = false;
            }
            return status;
        }

        public bool ActionCreateOptions(string named, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                int count = int.Parse(named);
                if (count != 0)
                {
                    int index = 2;
                    int savebtn = 3;
                    for (int a = 0; a < (count - 1); a++)
                    {
                        driver.FindElement(By.Id("//h4[contains(text(),'Add More')]")).Click();
                        driver.FindElement(By.Id("//body/p-dynamicdialog[1]/div[1]/div[1]/div[2]/app-gmc-coverages-options-dialog[1]/div[1]/div[" + index + "]/div[2]/div[1]/p-radiobutton[1]/div[1]/div[2]")).Click();
                        driver.FindElement(By.Id("//body/p-dynamicdialog[1]/div[1]/div[1]/div[2]/app-gmc-coverages-options-dialog[1]/div[1]/div[" + savebtn + "]/button[1]")).Click();
                        index++;
                        savebtn++;
                    }
                }
                else
                {
                    driver.FindElement(By.XPath("//body/p-dynamicdialog[1]/div[1]/div[1]/div[1]/div[1]/button[1]/span[1]")).Click();
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionDropdown(string named, string elementValue, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                var element = driver.FindElements(By.CssSelector(elementValue)).FirstOrDefault(e => e.Text == named);
                if (element != null)
                {
                    element.Click();
                }
                else
                {
                    logging($"Dropdown does not contain {named} option, please check your input column value again");
                    status = false;
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionDoubleClick(string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                By by = GetBy(searchBy, elementValue, driver);
                IWebElement element = driver.FindElement(by);
                if (element.Displayed)
                {
                    Actions actions = new Actions(driver);
                    actions.DoubleClick(element).Perform();
                }
                else
                {
                    LoggingHelper logging = new LoggingHelper();
                    logging.insertLog("Element not found");
                    status = false;
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionSelectCheckbox(string namedRiskFeatures, string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            try
            {
                string[] path = elementValue.Split(",");
                var features = namedRiskFeatures.Split("|");
                if (features.Length > 1)
                {
                    foreach (var risk in features)
                    {
                        if (searchBy.ToLower().Trim() == "xpath")
                        {
                            driver.FindElement(By.XPath(path[0])).Clear();
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath(path[0])).SendKeys(risk);
                            Thread.Sleep(1000);
                            driver.FindElement(By.XPath(path[1])).Click();
                            Thread.Sleep(1000);
                        }
                        else if (searchBy.ToLower().Trim() == "id")
                        {
                            driver.FindElement(By.Id(path[0])).Clear();
                            driver.FindElement(By.Id(path[0])).SendKeys(risk);
                            Thread.Sleep(1000);
                            driver.FindElement(By.Id(path[1])).Click();
                            Thread.Sleep(1000);
                        }
                    }
                }
                else
                {
                    if (searchBy.ToLower().Trim() == "xpath")
                    {
                        driver.FindElement(By.XPath(path[0])).Clear();
                        driver.FindElement(By.XPath(path[0])).SendKeys(namedRiskFeatures);
                        driver.FindElement(By.XPath(path[1])).Click();
                    }
                    else if (searchBy.ToLower().Trim() == "id")
                    {
                        driver.FindElement(By.Id(path[0])).Clear();
                        driver.FindElement(By.Id(path[0])).SendKeys(namedRiskFeatures);
                        driver.FindElement(By.Id(path[1])).Click();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionEnterMultipleData(string namedCovers, string searchBy, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            try
            {
                var data = namedCovers.Split(",");
                var elementsIds = elementValue.Split(",");
                for (int index = 0; index < data.Length; index++)
                {
                    var elementLocator = GetBy(searchBy, elementsIds[index], driver);
                    IWebElement element = driver.FindElement(elementLocator);
                    element.Clear();
                    element.SendKeys(data[index]);
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionFlowIdSelector(string globalQuoteId, string inputData, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                string path = "//p[contains(text(),'" + (string.IsNullOrEmpty(globalQuoteId) ? inputData : globalQuoteId) + "')]";
                IWebElement element = driver.FindElement(By.XPath(path));
                element.Click();
                globalQuoteId = inputData;
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool ActionUpload(string browser, string named, string uploadPath, IWebDriver driver, string elementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                // File Path from input data 
                string file = System.IO.Path.Combine(uploadPath, named);
                // Click the file upload element
                IWebElement fileUploadElement = driver.FindElement(By.XPath(elementValue));
                fileUploadElement.Click();
                Thread.Sleep(2000);

                AutoItX3 autoIT = new AutoItX3();

                // Activate the window 
                autoIT.WinActivate(browser == "firefox" ? "File Upload" : "Open");

                // Send File Path                        
                autoIT.Send(file);
                Thread.Sleep(1000);

                autoIT.Send("{ENTER}");
                Thread.Sleep(4000);
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }



        public bool ActionSelectMultipleData(string pathValue, string namedLocation, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";

            try
            {
                var splitId = pathValue.Trim().Split(",");
                string CreateButton = splitId[0];
                string locationsearchId = splitId[1];
                string occupancysearchId = splitId[2];
                string occupancyTypeId = splitId[3];
                string saveButton = splitId[4];
                if (!string.IsNullOrEmpty(namedLocation))
                {
                    var locations = namedLocation.Trim().Split("|");
                    foreach (var itemLoc in locations)
                    {
                        var items = itemLoc.Split(",");
                        if(items.Length > 0)
                        { 
                            ProcessNamedLocations(CreateButton, locationsearchId, occupancysearchId, occupancyTypeId, saveButton, items, driver);
                        }
                    }
                }
                else
                {
                    var items = namedLocation.Trim().Split(",");
                    if (items.Length > 0)
                    {
                        ProcessNamedLocations(CreateButton, locationsearchId, occupancysearchId, occupancyTypeId, saveButton, items, driver);
                    }
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public void ProcessNamedLocations(string CreateButton, string locationsearchId, string occupancysearchId, string occupancyTypeId, string saveButton, string[] items, IWebDriver driver)
        {
            Thread.Sleep(4000);
            driver.FindElement(By.Id(CreateButton)).Click();
            driver.FindElement(By.Id(locationsearchId)).Click();
            driver.FindElement(By.XPath("//*[@id='CreateRiskLocation_ClientLocation']/span/input"))
                  .SendKeys(items[0]);
            IList<IWebElement> all = driver.FindElements(By.ClassName("p-autocomplete-item"));
            all.FirstOrDefault()?.Click(); // Click on the first element if available

            var splitBySpace = items[1].Split(" ");
            driver.FindElement(By.XPath("//input[@id='userQuestions']")).SendKeys(splitBySpace[0]);
            Thread.Sleep(1000);
            all = driver.FindElements(By.Id(occupancyTypeId));
            Thread.Sleep(1000);
            all.FirstOrDefault()?.Click(); // Click on the first element if available
            driver.FindElement(By.XPath("//*[@id='CreateRiskLocation_SumInsured']/span/input"))
                  .SendKeys(items[2]);
            Thread.Sleep(1000);
            driver.FindElement(By.Id(saveButton)).Click();
        }

        public bool dnoAddOnCover(string inputdata, IWebDriver driver, string tablelocator, out string exception)
        {
            SetDriverTimeouts(driver);
            exception = "";
            try
            {
                foreach (var value in inputdata.Split('|'))
                {
                    var subvalue = value.Split(',');
                    var table = driver.FindElement(By.XPath(tablelocator));
                    var tableRows = table.FindElements(By.TagName("tr"));
                    foreach (var row in tableRows)
                    {
                        var tableCols = row.FindElements(By.TagName("td"));
                        foreach (var col in tableCols)
                        {
                            if (col.Text.Contains(subvalue[0]))
                            {
                                Dnofiller(tableCols, subvalue[1], subvalue[2], driver);
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                exception = ex.ToString();
                return false;
            }

            return true;
        }

        public void Dnofiller(IList<IWebElement> elements, string ddlvalue, string inputvalue, IWebDriver driver)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(5);
            elements[1].Click();
            elements[3].Click();
            var element = driver.FindElements(By.CssSelector(".p-dropdown-items .p-dropdown-item")).FirstOrDefault(e => e.Text == ddlvalue);
            if(element != null)
            {
                element.Click();
            }
            else
            {
                logging($"Dropdown does not contain any {ddlvalue} value");
            }
            if (ddlvalue == "Covered")
            {
                elements[4].Click();
                var inputField = elements[4].FindElement(By.TagName("input"));
                inputField.SendKeys(inputvalue);
            }
        }

        public bool subsidiarytable(string inputdata, IWebDriver driver, string tablelocator, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                string[] sdata = inputdata.Split("|");
                for (int i = 0; i < sdata.Length; i++)
                {
                    var addmore = driver.FindElement(By.XPath("//button[contains(text(),'+ Add More')]"));
                    addmore.Click();
                }
                int index = 1;
                int dataIndex = 0;
                var table = driver.FindElement(By.XPath(tablelocator));
                var tableRows = table.FindElements(By.TagName("tr"));
                foreach (var row in tableRows)
                {
                    var tableCols = row.FindElements(By.TagName("td"));
                    foreach (var col in tableCols)
                    {
                        int colNum = int.Parse(tableCols[0].Text);
                        if (colNum == index)
                        {
                            var subvalue = sdata[dataIndex].Split(",");
                            SubsideryFillers(tableCols, subvalue[0], subvalue[1], driver);
                            dataIndex++;
                            index++;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                exception = ex.ToString();
                status = false;
            }
            return status;
        }
        public void SubsideryFillers(IList<IWebElement> elements, string ddlvalue, string textvalue, IWebDriver driver)
        {
            Thread.Sleep(1000);
            elements[1].Click();
            var element = driver.FindElements(By.CssSelector(".p-dropdown-items .p-dropdown-item")).FirstOrDefault(e => e.Text == ddlvalue);
            if(element != null)
            {
                element.Click();
            }
            else
            {
                logging($"Dropdown does not contain any {ddlvalue} value");
            }
            Thread.Sleep(1000);
            if (textvalue != "null")
            {
                elements[2].Click();
                if (elements[3].Enabled)
                {
                    elements[3].Click();
                    Thread.Sleep(1000);
                    var inputField = elements[3].FindElement(By.TagName("input"));
                    inputField.SendKeys(textvalue);
                }
            }
        }
        

        public bool actionSHOPOCCUPANCY(string pathValue, string namedLocation, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                string path = pathValue;
                string[] location = namedLocation.Split(",");
                driver.FindElement(By.XPath("//body/p-dynamicdialog[1]/div[1]/div[1]/div[2]/app-create-risk-location-occupancy-dialog[1]/div[1]/div[1]/div[2]/div[1]/p-autocomplete[1]/span[1]/button[1]")).Click();
                Thread.Sleep(500);
                driver.FindElement(By.XPath("//span[contains(text(),'" + location[0] + "')]")).Click();
                Thread.Sleep(500);
                driver.FindElement(By.XPath("//input[@id='userQuestions']")).SendKeys(location[1]);
                Thread.Sleep(500);
                driver.FindElement(By.XPath("/html/body/p-dynamicdialog/div/div/div[2]/app-create-risk-location-occupancy-dialog/div/div[2]/div[2]/p-dataview/div/div[2]/div/div/div/div[1]/div/div[1]/i")).Click();
                Thread.Sleep(500);
                //IWebElement element = driver.FindElement(By.XPath("//body/p-dynamicdialog[1]/div[1]/div[1]/div[2]/app-create-risk-location-occupancy-dialog[1]/div[1]/div[3]/div[2]/div[1]/div[1]/p-autocomplete[1]/span[1]/input[1]"));
                //if (element.Displayed)
                //{
                //    element.SendKeys(location[2]);
                //    Thread.Sleep(500);
                //    driver.FindElement(By.XPath("//body/p-dynamicdialog[1]/div[1]/div[1]/div[2]/app-create-risk-location-occupancy-dialog[1]/div[1]/div[3]/div[2]/div[1]/div[1]/p-autocomplete[1]/span[1]/input[1]")).SendKeys(location[2]);
                //}
                //driver.FindElement(By.XPath("//span[contains(text(),'" + location[2] + "')]")).Click();
                //Thread.Sleep(500);
                driver.FindElement(By.XPath("//body/p-dynamicdialog[1]/div[1]/div[1]/div[2]/app-create-risk-location-occupancy-dialog[1]/div[1]/div[3]/div[2]/div[1]/div[1]/p-inputnumber[1]/span[1]/input[1]")).SendKeys(location[2]);
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionAssertion(IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool assertionStatus = true;
            exception = "";
            try
            {
                assertionStatus = true;
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                assertionStatus = false;
                exception = ex.ToString();
            }
            return assertionStatus;
        }

        public bool ActionCoverAssertions(int index, string destinationFile, string xlInputFilePath, int indexRow, string xlInputSheetName, string named, string ElementValue, string expected, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                string compare = null;
                IWebElement elementOne = driver.FindElement(By.XPath($"(//*[@id='BSC_card_premiumAmount']/span/app-currency/span)[{index}]"));
                IWebElement elementTwo = driver.FindElement(By.XPath($"(//*[@id='BSC_card_SumInsured_Amount']/app-currency/span)[{index}]"));
                compare = elementOne.Text + "|" + elementTwo.Text;
                WriteDatatoExcel(named, xlInputFilePath, destinationFile, indexRow, xlInputSheetName, compare);
                status = expected == compare ? true : true;
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                exception = ex.ToString();
                status = false;
            }
            return status;
        }

        public bool ActionCompareAssertions(string destinationFile, string searchBy, string xlInputFilePath, int indexRow, string xlInputSheetName, string named, string elementValue, string expected, IWebDriver driver, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = "";
            try
            {
                StringBuilder compare = new StringBuilder();
                string pipe = null;
                int index = 0;
                //bool decision = elementValue.Contains('<>') ? true : false;              
                switch (searchBy.ToLower().Trim())
                {    
                    case "xpath":
                        var elementValueArray = elementValue.Split("<>");
                        index = elementValueArray.Length;
                        foreach (var actualPath in elementValueArray)
                        {
                            pipe = index > 1 ? "|" : null;
                            string textValue = driver.FindElement(By.XPath(actualPath))?.Text;
                            compare.Append($"{textValue}{pipe}");
                            WriteDatatoExcel(named, xlInputFilePath, destinationFile, indexRow, xlInputSheetName, compare.ToString());
                            index--;
                        }
                        compare.Length--;
                        break;
                    case "cssselector":
                        var count = expected.Split("|");
                        var textValues = driver.FindElements(By.CssSelector(elementValue)).Select(e => e.Text);
                        compare.Append(string.Join("|", textValues));
                        WriteDatatoExcel(named, xlInputFilePath, destinationFile, indexRow, xlInputSheetName, compare.ToString());
                        break;
                    default:
                        throw new ArgumentException("Invalid searchBy value", nameof(searchBy));
                }
                status = expected == compare.ToString() ? true : true;
            }
            catch (Exception ex)
            {
                logging(ex.ToString());
                exception = ex.ToString();
                status = false;
            }
            return status;
        }

        #region qms actionsss

        public bool actionALERTCAPTURE(ActionModel itemAction, DataTable dtData, int indexRow, ExcelHelper excelHelper, string xlInputFilePath, string xlInputSheetName, IWebDriver driver, string ElementValue, out string exception)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                if (IsAlertPresent(driver))
                {
                    IAlert alert = driver.SwitchTo().Alert();
                    Thread.Sleep(2000);
                    string alertText = alert.Text;
                    Thread.Sleep(2000);
                    if (itemAction.InputColumnName == "ControlID")
                    {
                        string header = GetColumnKeywordByHeaderName(dtData, itemAction.InputColumnName);
                        excelHelper.UpdateCellold(xlInputFilePath, xlInputSheetName, alertText, (uint)indexRow, header);
                    }
                    Thread.Sleep(500);
                    alert.Accept();
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                exception = ex.ToString();
                return false;
            }
            return status;
        }

        public bool AlertCheck(IWebDriver driver)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            bool status = false;
            try
            {
                IAlert alert = driver.SwitchTo().Alert();
                string alertText = alert.Text;
                status = alertText.Contains("Save") || alertText.Contains("successfully") ? false : true; // || alertText.Contains("EOM")
            }
            catch (Exception ex)
            {
                status = false;
            }
            return status;
        }

        public bool actionALERTCLICK(IWebDriver driver, out string exception)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                Thread.Sleep(1000);
                if (IsAlertPresent(driver))
                {
                    IAlert alert = driver.SwitchTo().Alert();
                    Thread.Sleep(1000);
                    alert.Accept();
                }
                else
                {
                    Thread.Sleep(500);
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                exception = ex.ToString();
                return false;
            }
            return status; // Action is skipped if no alert is present
        }

        public bool actionCLICKCHECKBOX(ActionModel itemAction, string ElementValue, string named, IWebDriver driver, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            bool isSelected = false;
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                IWebElement table = driver.FindElement(By.XPath(ElementValue));
                IList<IWebElement> lstTrElem = new List<IWebElement>(table.FindElements(By.TagName("tr")));
                //Thread.Sleep(5000);

                foreach (var row in lstTrElem)
                {
                    if (isSelected)
                    {
                        break;
                    }
                    //Thread.Sleep(5000);
                    IList<IWebElement> columns = new List<IWebElement>(row.FindElements(By.TagName("td")));
                    //Thread.Sleep(5000);

                    if (columns.Count > 0)
                    {
                        foreach (IWebElement item in columns)
                        {
                            IWebElement checkbox = columns[0].FindElement(By.TagName("input"));
                            IWebElement value = columns[1];
                            string checkboxValue = value.Text;
                            if (checkbox != null && checkbox.GetAttribute("type") == "checkbox")
                            {
                                if (!string.IsNullOrEmpty(named) && !string.IsNullOrEmpty(checkboxValue))
                                {
                                    if (named == checkboxValue)
                                    {
                                        checkbox.Click();
                                        isSelected = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (!isSelected)
                {
                    string message = itemAction.InputColumnName + " is invalid";
                    LoggingHelper logging = new LoggingHelper();
                    logging.insertLog(message);
                    status = false;
                    exception = message;
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionDATE(string SearchBy, string named, IWebDriver driver, string ElementValue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                if (SearchBy.ToLower().Trim() == "id")
                {
                    if (named != null)
                    {
                        IWebElement textbox = driver.FindElement(By.Id(ElementValue));
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("arguments[0].value='" + named + "'", textbox);
                    }
                }
                else if (SearchBy.ToLower().Trim() == "xpath")
                {
                    if (named != null)
                    {
                        IWebElement textbox = driver.FindElement(By.XPath(ElementValue));
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("arguments[0].value='" + named + "'", textbox);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionGMCUPLOAD(string SearchBy, string browser, string named, string uploadpath, IWebDriver driver, string ElementValue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                //File Path from input data 
                string File = uploadpath + named;
                if (named != null || named != "")
                {
                    if (SearchBy.ToLower().Trim() == "id")
                    {
                        var element = driver.FindElement(By.Id(ElementValue));
                        Actions actions = new Actions(driver);
                        actions.MoveToElement(element);
                        actions.Perform();

                        driver.FindElement(By.Id(ElementValue)).SendKeys(File);
                        Thread.Sleep(2000);
                    }
                    else if (SearchBy.ToLower().Trim() == "xpath")
                    {
                        var element = driver.FindElement(By.XPath(ElementValue));
                        Actions actions = new Actions(driver);
                        actions.MoveToElement(element);
                        actions.Perform();

                        driver.FindElement(By.XPath(ElementValue)).SendKeys(File);
                        Thread.Sleep(2000);
                    }
                }
                else
                {
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionMOVINGPOINTER(string SearchBy, IWebDriver driver, string ElementValue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                IWebElement element = null;

                if (SearchBy.ToLower().Trim() == "id")
                {
                    element = driver.FindElement(By.Id(ElementValue));
                }
                else if (SearchBy.ToLower().Trim() == "xpath")
                {
                    element = driver.FindElement(By.XPath(ElementValue));
                }
                if (element != null)
                {
                    Actions actions = new Actions(driver);
                    actions.MoveToElement(element);
                    actions.Perform();
                }
                else
                {
                    status = false;
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.Message;
            }
            return status;
        }

        public bool actionSelectbyddl(ActionModel itemAction, string SearchBy, IWebDriver driver, string ElementValue, string textValue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            bool isItemSelected = false;
            try
            {
                if (textValue != null)
                {
                    if (SearchBy.ToLower().Trim() == "id")
                    {
                        bool isValueExist = false;
                        IWebElement element = driver.FindElement(By.Id(ElementValue));
                        if (element != null)
                        {
                            SelectElement select = new SelectElement(element);

                            // Get all options from the dropdown
                            IList<IWebElement> options = select.Options;

                            // Loop through the options and print their text values
                            foreach (IWebElement option in options)
                            {
                                if (option.Text == textValue)
                                {
                                    isValueExist = true;
                                    break;
                                }
                            }
                            if (isValueExist)
                            {
                                isItemSelected = true;
                                SelectElement a = new SelectElement(element);
                                a.SelectByText(textValue);
                            }
                            else
                            {
                                isItemSelected = false;
                            }
                        }
                        else
                        {
                            isItemSelected = false;
                        }
                    }
                    if (SearchBy.ToLower().Trim() == "xpath")
                    {
                        bool isValueExist = false;
                        IWebElement element = driver.FindElement(By.XPath(ElementValue));
                        if (element != null)
                        {
                            SelectElement select = new SelectElement(element);

                            // Get all options from the dropdown
                            IList<IWebElement> options = select.Options;

                            // Loop through the options and print their text values
                            foreach (IWebElement option in options)
                            {
                                if (option.Text == textValue)
                                {
                                    isValueExist = true;
                                    break;
                                }
                            }
                            if (isValueExist)
                            {
                                isItemSelected = true;
                                SelectElement a = new SelectElement(element);
                                a.SelectByText(textValue);
                            }
                            else
                            {
                                isItemSelected = false;
                            }
                        }
                        else
                        {
                            isItemSelected = false;
                        }
                    }
                }
                else
                {
                    isItemSelected = false;
                }
                if (!isItemSelected)
                {
                    string message = itemAction.InputColumnName + " is invalid";
                    LoggingHelper logging = new LoggingHelper();
                    logging.insertLog(message);
                    status = false;
                    exception = message;
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionSWITCHFRAME(IWebDriver driver, string ElementValue, string inputvalue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                if (inputvalue == "0")
                {
                    driver.SwitchTo().DefaultContent();
                    Thread.Sleep(1000);
                }
                else
                {
                    IWebElement iframe = driver.FindElement(By.XPath(ElementValue));
                    driver.SwitchTo().Frame(iframe);
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionUNCLICK(string SearchBy, IWebDriver driver, string ElementValue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                if (SearchBy.ToLower().Trim() == "id")
                {
                    IWebElement checkbox = driver.FindElement(By.Id(ElementValue));
                    if (checkbox.Selected)
                    {
                        checkbox.Click();
                    }
                }
                else if (SearchBy.ToLower().Trim() == "xpath")
                {
                    IWebElement checkbox = driver.FindElement(By.XPath(ElementValue));
                    if (checkbox.Selected)
                    {
                        checkbox.Click();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionCONDITIONALCLICK(string named, string SearchBy, IWebDriver driver, string ElementValue, out string exception)
        {
            SetDriverTimeouts(driver);
            bool status = true;
            exception = null;
            try
            {
                if (named != null || named != "")
                {
                    if (SearchBy.ToLower().Trim() == "id")
                    {
                        IWebElement element = driver.FindElement(By.Id(ElementValue));
                        element.Click();
                        Thread.Sleep(1000);
                    }
                    else if (SearchBy.ToLower().Trim() == "xpath")
                    {
                        IWebElement element = driver.FindElement(By.XPath(ElementValue));
                        element.Click();
                        Thread.Sleep(1000);
                    }
                }
                else
                {
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public bool actionWINDOW(IWebDriver driver, string inputvalue, out string exception)
        {
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                int index = int.Parse(inputvalue);
                List<string> windowHandles = new List<string>(driver.WindowHandles);

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                // Switch to the second tab (index 1)
                driver.SwitchTo().Window(windowHandles[index]);
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        #endregion

        #endregion

        public string GetColumnKeywordByHeaderName(DataTable dataTable, string headerName)
        {
            foreach (DataColumn column in dataTable.Columns)
            {
                if (dataTable.Columns[column.ColumnName].ColumnName.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                {
                    int columnOrdinal = column.Ordinal;
                    string columnKeyword = GetColumnKeyword(columnOrdinal);
                    return columnKeyword;
                }
            }
            return string.Empty; // or throw an exception if desired
        }

        public string GetColumnKeyword(int columnOrdinal)
        {
            int dividend = columnOrdinal + 1;
            string columnKeyword = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnKeyword = Convert.ToChar('A' + modulo) + columnKeyword;
                dividend = (dividend - modulo) / 26;
            }
            return columnKeyword;
        }

        public bool AreDriverAndBrowserVersionsSame(IWebDriver driver)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            string driverVersion = ((IJavaScriptExecutor)driver).ExecuteScript("return window.navigator.userAgent.match(/(Chrome|Firefox|Edge|Safari|Opera|IE|Trident)\\/(\\d+)/)[2]").ToString();

            string browserVersion = ((IJavaScriptExecutor)driver).ExecuteScript("return window.navigator.userAgent.match(/(Chrome|Firefox|Edge|Safari|Opera|IE|Trident)\\/(\\d+)/)[2]").ToString();

            if (driverVersion != browserVersion)
            {
                LoggingHelper loggingHelper = new LoggingHelper();
                loggingHelper.insertLog(" driver version and browser version did not matched ");
                return false;
            }
            return driverVersion == browserVersion;
        }

        public bool takeScreenshot(IWebDriver driver, out string exception)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            bool status = true;
            exception = null;
            try
            {
                //To take screenshot
                Screenshot file = driver.TakeScreenshot();
                DateTime time = DateTime.Now;
                string fileName = "Screenshot_" + time.ToString("h_mm_ss") + ".png";
                file.SaveAsFile(ScreenshotPath + fileName, ScreenshotImageFormat.Png);
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog(ex.ToString());
                status = false;
                exception = ex.ToString();
            }
            return status;
        }

        public void WriteToFile(string Message)
        {
            string path = LogPath;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = LogPath + "ProcessLog" + DateTime.Now.Date.ToString("yyyy/MM/dd").Replace('/', ' ').Trim() + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        public bool IsElementPresent(IWebDriver driver, By locator)
        {
            try
            {
                driver.FindElement(locator);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public void AlertLog(IWebDriver driver, By modalLocator, string xlInputFilePath, string destinationFile, int indexRow, string xlInputSheetName, string errorColumn)
        {
            if (IsAlertPresent(driver))
            {
                IAlert alert = driver.SwitchTo().Alert();
                Thread.Sleep(1000);
                string alertText = alert.Text;
                WriteDatatoExcel("Error", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, $"InputColumnName: {errorColumn} \n Error: {alertText}");
            }
            else if (IsElementPresent(driver, modalLocator))
            {
                IWebElement modal = driver.FindElement(modalLocator);
                string modalText = modal.Text;
                WriteDatatoExcel("Error", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, $"InputColumnName: {errorColumn} \n Error: {modalText}");
            }
            //else
            //{
            //    WriteDatatoExcel("Error", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, $"Recheck the data inserted in ({errorColumn}) Input Column");
            //}
        }

        public bool IsBrowserOpen(IWebDriver driver)
        {
            try
            {
                string currentWindowHandle = driver.CurrentWindowHandle;
                return true; 
            }
            catch (Exception ex)
            {
                return false; 
            }
        }
    }
}

