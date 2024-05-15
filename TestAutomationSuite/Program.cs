// See https://aka.ms/new-console-template for more information

using AutomationUAT.Utility;
using AutomationUAT.ViewModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Timers;
using TestAutomationSuite.dbModels;
using TestAutomationSuite.Utility;

//Global Declareation
#region Global Declaration
IWebDriver driver = null;
string
    keepbrowseropen,
    implicitwait,
    outputFolder,
    ActionSheet,
    DataSheet,
    xlInputFilePath,
    xlActionFilePath,
    xlInputSheetName,
    xlActionSheetName,
    OutputFilePath,
    ExecutionTimeColumn,
    ScreenshotPath,
    StageColumn,
    InputTimeTakenColumn,
    InputErrorColumn,
    driverPath,
    Result,
    QuoteID,
    Premium,
    LogPath,
    LogDirectory,
    browser,
    uploadpath,
    Timer,
    globalQuoteId,
    globalPremium,
    connectionstring,
    logmessege,
    destinationFile,
    sourceFile,
    filename,
    BSCcovers,
    Excpected = null,
    errorColumn = null,
    proStatus = null,
    OutputfileName = null;
bool driverStatus = true;
System.Timers.Timer timer;
bool isRunning;
int indexRow = 2;
int indexCover = 1;
int indexAssert = 1;
int employeeIndex = 0;
bool excelReadStatus = true;
bool assertionStatus = false;
ActionHelper actionHelper = new ActionHelper();
LoggingHelper LoggingHelper = new LoggingHelper();
DateTime Start;
DAL dal = new DAL();
ProcessModel processDataModel = new ProcessModel(); // Add Process data
int processInsertresult = 0;
#endregion


string appSettings(string appSettings)
{
    return ConfigurationManager.AppSettings[appSettings];
}

void InitializeConfigurationSettings()
{
    Timer = appSettings("Timer");
    browser = appSettings("browser");
    driverPath = appSettings("driverPath");
    implicitwait = appSettings("implicitwait");
    xlInputFilePath = appSettings("xlInputFilePath");
    xlActionFilePath = appSettings("xlActionFilePath");
    xlInputSheetName = appSettings("xlInputSheetName");
    xlActionSheetName = appSettings("xlActionSheetName");
    uploadpath = appSettings("uploadpath");
    OutputFilePath = appSettings("xOutputFilePath");
    ScreenshotPath = appSettings("ScreenshotPath");
    LogPath = appSettings("LogPath");
    LogDirectory = appSettings("LogDirectory");
    keepbrowseropen = appSettings("keepbrowseropen");
    connectionstring = ConfigurationManager.ConnectionStrings["TestAuto"].ConnectionString;
    string a = null;
}

try
{
    LoggingHelper.insertLog("*** BOT IS ACTIVATED ***");
    InitializeConfigurationSettings();
    if (int.TryParse(Timer, out int timelap))
    {
        List<string> appPaths = new List<string>
        {
        driverPath, xlInputFilePath, xlActionFilePath, uploadpath, OutputFilePath, ScreenshotPath, LogPath, LogDirectory
        };
        foreach (var path in appPaths)
        {
            if (!Directory.Exists(path))
            {
                LoggingHelper.insertLog($"The path {path} provided in App.Config file does not exist.\n ***BOT IS DEACTIVATED***");
                Environment.Exit(1);
            }
        }
        TestModule();
        timer = new System.Timers.Timer(timelap);
        timer.Elapsed += TimerElapsed;
        timer.AutoReset = true;
        timer.Start();
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
        timer.Stop();
        timer.Dispose();
    }
    else
    {
        LoggingHelper.insertLog("Please provide a valid Timer in App.config file");
        LoggingHelper.insertLog("*** BOT IS DEACTIVATED ***");
        Environment.Exit(1);
    }
}
catch (Exception ex)
{
    LoggingHelper.insertLog($"Error - {ex}");
}

void TimerElapsed(object sender, ElapsedEventArgs e)
{
    if (!isRunning)
    {
        InitializeConfigurationSettings();
        TestModule();
    }
}
void TestModule()
{
    string exceptionFileExecution = null;
    bool isFileExecutionSuccessful = FileExecution(out exceptionFileExecution);
}

#region File Execution

//This is File Execution function
bool FileExecution(out string exceptionFileExecution)
{
    isRunning = true;
    bool status = true;
    globalQuoteId = null;
    globalPremium = null;
    string exception = null;
    exceptionFileExecution = null;
    indexRow = 2;
    ProjectModel projectMasterData = new ProjectModel();
    //Insert Process data

    List<ProjectModel> projectMastersList = new List<ProjectModel>();

    try
    {
        ExcelHelper excelHelper = new ExcelHelper();

        // Get Action file from folder : Convert to JSON and insert project into database
        string actionFilesToConvert = string.Empty;
        actionFilesToConvert = excelHelper.getActionFileFromPath(xlActionFilePath, out exception);

        #region insert and get project information
        /*  if (!string.IsNullOrEmpty(actionFilesToConvert))
          {
              string actionFileName = actionFilesToConvert;
              string filepath = xlActionFilePath + actionFileName;
              DataTable dtActions = excelHelper.GetDataTableFromExcelFile(filepath, xlActionSheetName);
              List<ActionModel> actionsLsts = actionHelper.ConvertDataTable<ActionModel>(dtActions);
              actionsLsts = actionHelper.ConvertDataTable<ActionModel>(dtActions);
              string actionjson = JsonConvert.SerializeObject(actionsLsts);
              string projectName = System.IO.Path.GetFileNameWithoutExtension(filepath);

              ProjectModel dataOfProjects = new ProjectModel();  // Project Information to save into database
              dataOfProjects.projectName = projectName;     // Project Name
              dataOfProjects.actionFileJson = actionjson;  // Action JSON
              string[] splitname = projectName.Split(" ");
              dataOfProjects.inputFilePath = splitname[0] + "\\"; //Input file path       
              int projectInsertresult = dal.InsertProject(dataOfProjects);// Save Project to db and get projectId

              // Delete Action file once Action JSON saved in database
              File.Delete(filepath);
              LoggingHelper.insertLog("Action File " + projectName + " updated to Project Manager");
          }*/
        #endregion

        #region Get Projects and Run 
        projectMastersList = dal.GetAllProjects();
        int count = 0;
        if (projectMastersList.Count > 0)
        {
            foreach (ProjectModel project in projectMastersList)
            {
                if (project.projectStatus == "Active")
                {
                    projectMasterData = new ProjectModel();
                    projectMasterData = project;
                    #region pick up JSON action from the project Manager
                    List<ActionModel> actionsLst = new List<ActionModel>();
                    //convet ActionJSON  to Action List
                    actionsLst = JsonConvert.DeserializeObject<List<ActionModel>>(projectMasterData.actionFileJson);
                    #endregion

                    #region Pick up input file
                    string inputfile = string.Empty;
                    xlInputFilePath = ConfigurationManager.AppSettings["xlInputFilePath"];
                    if (!Directory.Exists(xlInputFilePath + projectMasterData.inputFilePath))
                    {
                        LoggingHelper.insertLog($"Create a sub-folder in path {xlInputFilePath} with name {projectMasterData.inputFilePath.Replace("\\", "")} and then insert Input data file in it to start processing");
                        LoggingHelper.insertLog("***BOT IS DEACTIVATED***");
                        Environment.Exit(1);
                    }
                    #endregion
                    else
                    {
                        // Combine Input Folder and Input File Path                    
                        inputfile = excelHelper.getActionFileFromPath(xlInputFilePath + projectMasterData.inputFilePath, out exception);
                        if (!string.IsNullOrEmpty(inputfile))
                        {
                            if (processInsertresult == 0)
                            {
                                #region insert and get process data                       
                                processDataModel = new ProcessModel();//Add Process data
                                int winProcessID = Process.GetCurrentProcess().Id;// Windows Process No
                                processDataModel.winProcID = winProcessID;
                                processDataModel.startTime = DateTime.Now.ToString();
                                processDataModel.endTime = Convert.ToString(default(DateTime));
                                processDataModel.processStatus = "Running";
                                processInsertresult = dal.InsertProcess(processDataModel);
                                #endregion
                            }

                            LoggingHelper.insertLog(" Started Processing... ");

                            excelHelper = new ExcelHelper();

                            LoggingHelper.insertLog(" picked up data file: " + inputfile);

                            ProjectModel dataOfProjects = new ProjectModel();

                            LoggingHelper.insertLog(" Picked up Action JSON of Project: " + projectMasterData.projectName);

                            #region Read and convet Input file to JSON Array

                            outputFolder = projectMasterData.inputFilePath + "_Output";
                            xlInputFilePath = xlInputFilePath + projectMasterData.inputFilePath + "\\" + inputfile;
                            DataTable dtData = excelHelper.GetDataTableFromExcelFile(xlInputFilePath, xlInputSheetName);
                            if (dtData != null && dtData.Rows.Count > 0)
                            {
                                var jsonDatainput = excelHelper.DataTableToJSONWithJSONNet(dtData);
                                JArray jsonArray = JArray.Parse(jsonDatainput);
                                #endregion

                                #region Insert and Get file info into database                       
                                FileModels fileInfo = new FileModels();//Insert File Data //Add File Data
                                FileModels fileModels = new FileModels();
                                fileModels.fileName = System.IO.Path.GetFileName(xlInputFilePath);
                                fileModels.startTime = DateTime.Now.ToString();
                                fileModels.endTime = Convert.ToString(default(DateTime));
                                fileModels.processId = processInsertresult;
                                fileModels.projectId = projectMasterData.projectId;
                                fileModels.status = "work in progress";
                                int fileInsertresult = dal.InsertFile(fileModels);

                                //Get newly added File Data
                                fileInfo = dal.GetfileDataById(fileInsertresult);
                                #endregion
                                LoggingHelper.insertLog(" Sheet name: " + "TestData");

                                #region Insert record --start--
                                //Insert Record data
                                RecordModel recordInfo = new RecordModel();
                                RecordModel recordModel = new RecordModel();
                                recordModel.fileId = fileInfo.fileId;
                                recordModel.inputJsonRow = jsonDatainput;
                                recordModel.startTime = DateTime.Now.ToString();
                                recordModel.fileId = fileInsertresult;
                                recordModel.processId = processInsertresult;
                                recordModel.endTime = Convert.ToString(default(DateTime));
                                //recordModel.OutputfileName = OutputfileName;
                                #endregion
                                //JSON Execution Function
                                bool isJSONInputTestSuccessful = JSONInputTest(excelHelper, jsonArray, actionsLst, dtData, recordModel, recordInfo, fileInfo, out exception);
                            }
                            //Input File delete
                            File.Delete(xlInputFilePath);
                        }
                        else
                        {
                            count = count + 1;
                        }
                    }
                    if (count == projectMastersList.Count)
                    {
                        if (processInsertresult != 0)
                        {
                            bool isProcessPassed = dal.getRecordstatus(processInsertresult);
                            string processStatus = "Passed";
                            if (proStatus == "Failed")
                            {
                                processStatus = "Failed";
                                proStatus = "";
                            }
                            ProcessModel processInfo = dal.GetprocessDataById(processInsertresult);
                            processInfo.endTime = DateTime.Now.ToString();
                            processInfo.processStatus = processStatus;
                            var processResult = dal.updateProcessEndTimeAndStatus(processInfo);
                            processInsertresult = 0;
                        }
                        //TODO: add log for no project data found
                        LoggingHelper.insertLog(" No .xlsx format input file is present in any of the input folders associated with projects present in Database");
                    }
                }
                else
                {
                    LoggingHelper.insertLog(" No .xlsx data files were found in active project folders \n");
                }
            }
        }
        else
        {
            //TODO: No projects found           
            LoggingHelper.insertLog(" No Project was found in Project Manager Table in Database or the File you were trying to Update is not in .xlsx format");
            LoggingHelper.insertLog(" ***BOT IS DEACTIVATED***");
            Environment.Exit(1);
        }
        #endregion
    }
    catch (Exception ex)
    {
        status = false;
        exceptionFileExecution = ex.ToString();
    }
    isRunning = false;
    return status;
}
#endregion

#region JSON Input Test
//This is JSON Execution function
bool JSONInputTest(ExcelHelper excelHelper, JArray jsonArray, List<ActionModel> actionsLst, DataTable dtData, RecordModel recordModel, RecordModel recordInfo, FileModels fileInfo, out string exception)
{
    bool status = true;
    exception = null;
    destinationFile = null;
    indexRow = 2;
    string exceptionJSONInputTest = null;
    try
    {
        int i = 0;
        foreach (var item in jsonArray)//INPUT TEST CASES
        {
            indexCover = 1;
            indexAssert = 1;
            LoggingHelper.insertLog(" picked up row: " + (indexRow - 1).ToString());

            bool isActionExecutionSuccessful = true;
            Start = DateTime.Now;

            recordModel.startTime = DateTime.Now.ToString();
            var rRecord = dal.updateRecordStartTimeAndStatus(recordInfo);
            dynamic inputData = JObject.Parse(item.ToString());
            string startstage = inputData["Stage"];
            if (string.IsNullOrEmpty(startstage) && startstage != "Last Stage")
            {
                startstage = "1";
            }
            if (startstage == "Last Stage")
            {
                LoggingHelper.insertLog(" Test is Already Passed ");
            }
            if (startstage == "1")
            {
                if (!actionHelper.IsBrowserOpen(driver))
                {
                    actionHelper.ActionLaunchbrowser(browser, driverPath, out exception, out driver);
                }
            }

            //Action Execution function
            DataTable dt1 = new DataTable();
            dt1 = dtData.Clone();
            int j = 0;
            foreach (DataRow myRow in dtData.Rows)
            {
                if (j == i)
                {
                    dt1.ImportRow(myRow);
                    break;
                }
                j++;
            }
            isActionExecutionSuccessful = ActionExecution(excelHelper, item, actionsLst, dtData, inputData, startstage, indexRow, fileInfo, recordModel, recordInfo, out exceptionJSONInputTest);
            //IF TEST PASSED
            if (isActionExecutionSuccessful)
            {
                DateTime end1 = DateTime.Now;
                TimeSpan timeDifference1 = end1 - Start;

                if (string.IsNullOrEmpty(destinationFile))
                {
                    //Copy Input to Output file 
                    string sourceFile = xlInputFilePath;
                    string outputDirectory = OutputFilePath + outputFolder + "\\";
                    if (!Directory.Exists(outputDirectory))
                    {
                        Directory.CreateDirectory(outputDirectory);
                    }
                    DateTime time = DateTime.Now;
                    string filename = System.IO.Path.GetFileName(xlInputFilePath).Replace(".xlsx", "");
                    destinationFile = outputDirectory + filename + "_Output" + time.ToString("h_mm_ss_mm") + ".xlsx";
                    OutputfileName = filename + "_Output" + time.ToString("h_mm_ss_mm") + ".xlsx";
                    File.Copy(sourceFile, destinationFile, true);
                }

                //Log the Input
                DataTable dtDataOutPut = excelHelper.GetDataTableFromExcelFile(destinationFile, xlInputSheetName);
                string header1 = actionHelper.GetColumnKeywordByHeaderName(dtDataOutPut, "Stage");
                excelHelper.UpdateCellold(destinationFile, xlInputSheetName, "Last Stage", (uint)indexRow, header1);

                string header2 = actionHelper.GetColumnKeywordByHeaderName(dtDataOutPut, "Error");
                excelHelper.UpdateCellold(destinationFile, xlInputSheetName, "No Error", (uint)indexRow, header2);

                string header3 = actionHelper.GetColumnKeywordByHeaderName(dtDataOutPut, "Execution Time");
                excelHelper.UpdateCellold(destinationFile, xlInputSheetName, timeDifference1.ToString(), (uint)indexRow, header3);
                
                string header4 = actionHelper.GetColumnKeywordByHeaderName(dtDataOutPut, "Result");
                excelHelper.UpdateCellold(destinationFile, xlInputSheetName, "PASS", (uint)indexRow, header4);

                string header5 = actionHelper.GetColumnKeywordByHeaderName(dtDataOutPut, "AssertionStatus");
                string Astatus = !assertionStatus ? "FAIL" : "PASS";
                excelHelper.UpdateCellold(destinationFile, xlInputSheetName, Astatus, (uint)indexRow, header5);

                ExcelHelper excelHelper1 = new ExcelHelper();
                DataTable dtDataOutPutlatest = excelHelper1.GetDataTableFromExcelFile(destinationFile, xlInputSheetName);
                var jsonDataoutput = excelHelper1.DataTableToJSONWithJSONNet(dtDataOutPutlatest);

                DataTable dtDataInputPutlatest = excelHelper.GetDataTableFromExcelFile(xlInputFilePath, xlInputSheetName);
                var jsonDatainput = excelHelper.DataTableToJSONWithJSONNet(dtDataInputPutlatest);

                recordModel.inputJsonRow = jsonDatainput;
                recordModel.outputJsonRow = jsonDataoutput;
                recordModel.status = "Passed";
                recordModel.lastStageProcessed = "Last Stage";
                recordModel.lastError = "No Error";
                recordModel.OutputfileName = OutputfileName;
                int recordInsertresult = dal.listofRecordModel(recordModel);
                recordInfo = dal.GetrecordDataById(recordInsertresult);

                //Add Log data
                LogModel lModel = new LogModel();
                lModel.logTime = timeDifference1.ToString();
                lModel.description = "No Error";
                lModel.recordId = recordInsertresult;
                string logs = dal.listofLogModel(lModel);
                string update = dal.updateLogs(lModel);

                //Log Status
                LoggingHelper.insertLog("*************** TEST PASSED ***************" + "\n");

                //Update File endtime
                fileInfo.endTime = DateTime.Now.ToString();
                var fresult = dal.updateFileEndTime(fileInfo);
                //Update File Status
                fileInfo.status = "pass";
                var fstatus = dal.updatefilestatus(fileInfo);
                //Update Outputfile name
                fileInfo.outputfilename = OutputfileName;
                var foutputfile = dal.updateoutputfilename(fileInfo);
                //Update Record endtime
                recordInfo.endTime = DateTime.Now.ToString();
                var rRecords = dal.updateRecordEndTimeAndStatus(recordInfo);
                if (keepbrowseropen == "No") driver.Quit();                
                //driver.Quit();
            }
            indexRow++;
            i++;
        }
    }
    catch (Exception ex)
    {
        exception = ex.ToString();
        LoggingHelper.insertLog(ex.ToString());
        status = false;
    }
    return status;
}
#endregion

#region Action Execution
//This is Action Execution Function
bool ActionExecution(ExcelHelper excelHelper, JToken? item, List<ActionModel> actionsLst, DataTable dtData, dynamic inputData, string startstage, int indexRow, FileModels fileInfo, RecordModel recordModel, RecordModel recordInfo, out string exception)
{
    bool actionStatus = true;
    bool headerStatus = true;
    bool status = true;
    exception = null;
    List<string> columns = excelHelper.IsHeaderPresent(xlInputFilePath);
    try
    {
        //Taking actions from action list
        foreach (var itemAction in actionsLst)
        {
            actionStatus = true;
            headerStatus = true;
            if (!string.IsNullOrEmpty(startstage))
            {
                int srno = Convert.ToInt32(itemAction.SrNo);
                int stageNo = Convert.ToInt32(startstage);
                if (Convert.ToInt32(itemAction.Stage) >= stageNo)
                {
                    LoggingHelper.insertLog($"Executing Action: {itemAction.SrNo} {itemAction.ActionType}: {itemAction.Description}\n");
                    string namedInput = null;
                    if (!string.IsNullOrEmpty(itemAction.InputColumnName))
                    {
                        namedInput = inputData[itemAction.InputColumnName];
                        if (itemAction.InputColumnName == "QuoteId" || itemAction.InputColumnName == "QuoteIdInput"
                            || itemAction.InputColumnName == "Premium" || itemAction.InputColumnName == "ControlID" || itemAction.InputColumnName.Contains("Actual"))
                        {

                        }
                        else if (!columns.Contains(itemAction.InputColumnName))
                        {
                            actionStatus = false;
                            headerStatus = false;
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(namedInput))
                            {
                                if (itemAction.ActionType.ToLower().Trim() == "employeetabledata")
                                {
                                    employeeIndex++;
                                }
                                LoggingHelper.insertLog($"Executing Action: {itemAction.SrNo} {itemAction.ActionType}: The Action {itemAction.ActionType} was skipped because data was not present in {itemAction.InputColumnName} Column\n");
                                continue;
                            }
                        }
                    }
                    if (itemAction.ImageFlag != "")
                    {
                        if (itemAction.ImageFlag == "Yes")
                        {
                            actionHelper.takeScreenshot(driver, out exception);
                        }
                    }

                    //All actions switch cases
                    string named = null;
                    if (actionStatus == true)
                    {
                        switch (itemAction.ActionType.ToLower().Trim())
                        {
                            case null:
                                actionStatus = false;
                                break;

                            case "url":
                                if (!actionHelper.IsBrowserOpen(driver))
                                {
                                    actionHelper.ActionLaunchbrowser(browser, driverPath, out exception, out driver);
                                }                               
                                actionStatus = actionHelper.ActionURL(driver, browser, driverPath, itemAction.InputValue, out exception);
                                break;

                            case "text":
                                string inputText = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                if (itemAction.InputColumnName == "QuoteIdInput")
                                {
                                    errorColumn = itemAction.InputColumnName;
                                    string iText = inputData["QuoteId"];
                                    if (string.IsNullOrEmpty(globalQuoteId))
                                    {
                                        globalQuoteId = iText;
                                    }
                                }
                                actionStatus = actionHelper.ActionText(driver, itemAction.SearchBy, itemAction.Value, inputText, itemAction.InputColumnName, globalQuoteId, globalPremium, out exception);
                                break;

                            case "quitbrowser":
                                actionStatus = actionHelper.ActionQuitBrowser(fileInfo, out exception, driver, out driverStatus);
                                var fresult = dal.updateFileEndTime(fileInfo);
                                break;

                            case "closetab":
                                actionStatus = actionHelper.ActionCloseTab(out exception, driver, out driverStatus);
                                break;

                            case "clickradiobutton":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionClickRadioButton(named, driver, out exception);
                                break;

                            case "forceclick":
                                actionStatus = actionHelper.ActionForceClick(itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "coverclick":
                                actionStatus = actionHelper.ActionCoverClick(indexCover, itemAction.SearchBy, driver, itemAction.Value, out exception);
                                indexCover++;
                                break;

                            case "click":
                                actionStatus = actionHelper.ActionClick(itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "clear":
                                actionStatus = actionHelper.actionCLEAR(itemAction.SearchBy, itemAction.Value, out exception, driver);
                                break;

                            case "unclick":
                                actionStatus = actionHelper.actionUNCLICK(itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "window":
                                actionStatus = actionHelper.actionWINDOW(driver, itemAction.InputValue, out exception);
                                break;

                            case "switchframe":
                                actionStatus = actionHelper.actionSWITCHFRAME(driver, itemAction.Value, itemAction.InputValue, out exception);
                                break;

                            case "doubleclick":
                                actionStatus = actionHelper.ActionDoubleClick(itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "alertclick":
                                actionStatus = actionHelper.actionALERTCLICK(driver, out exception);
                                break;

                            case "alertcapture":
                                actionStatus = actionHelper.actionALERTCAPTURE(itemAction, dtData, indexRow, excelHelper, xlInputFilePath, xlInputSheetName, driver, itemAction.Value, out exception);
                                break;

                            case "read":
                                actionStatus = actionHelper.ActionRead(itemAction, dtData, indexRow, excelHelper, out exception, driver, xlInputFilePath, destinationFile, xlInputSheetName, out globalPremium, out globalQuoteId);
                                break;

                            case "conditionalclick":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionCONDITIONALCLICK(named, itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "coverassertions":
                                if (!string.IsNullOrEmpty(Excpected))
                                {
                                    errorColumn = itemAction.InputColumnName;
                                    actionStatus = actionHelper.ActionCoverAssertions(indexAssert, destinationFile, xlInputFilePath, indexRow, xlInputSheetName, itemAction.InputColumnName, itemAction.Value, Excpected, driver, out exception);
                                    assertionStatus = actionStatus;
                                    indexAssert++;
                                    Excpected = null;
                                }
                                else
                                {
                                    actionStatus = true;
                                    assertionStatus = true;
                                }
                                break;
                                    
                            case "tabledatainput":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionTabledatainput(itemAction.Value, named, driver, out exception);
                                break;            

                            case "createoptions":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionCreateOptions(named, driver, out exception);
                                break;

                            case "entermultipledata":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionEnterMultipleData(named, itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "date":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionDATE(itemAction.SearchBy, named, driver, itemAction.Value, out exception);
                                break;

                            case "dropdown":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionDropdown(named, itemAction.Value, driver, out exception);
                                break;

                            case "selectcheckbox":
                                string namedriskfeatures = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionSelectCheckbox(namedriskfeatures, itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;
                                
                            case "flowidselector":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionFlowIdSelector(globalQuoteId, named, driver, out exception);
                                break;

                            case "upload":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionUpload(browser, named, uploadpath, driver, itemAction.Value, out exception);
                                break;

                            case "gmcupload":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionGMCUPLOAD(itemAction.SearchBy, browser, named, uploadpath, driver, itemAction.Value, out exception);
                                break;

                            case "dnoaddoncovers":
                                named = inputData[itemAction.InputColumnName];
                                actionStatus = actionHelper.dnoAddOnCover(named, driver, itemAction.Value, out exception);
                                break;

                            case "subsidiarytable":
                                named = inputData[itemAction.InputColumnName];
                                actionStatus = actionHelper.subsidiarytable(named, driver, itemAction.Value, out exception);
                                break;

                            case "selectmultipledata":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionSelectMultipleData(itemAction.Value, named, driver, out exception);
                                break;

                            case "selectbyddl":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionSelectbyddl(itemAction, itemAction.SearchBy, driver, itemAction.Value, named, out exception);
                                break;

                            case "clearandenter":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionClearAndEnter(itemAction.InputColumnName, itemAction.SearchBy, driver, itemAction.Value, named, out exception);
                                break;

                            case "wait":
                                actionStatus = actionHelper.ActionWait(driver, itemAction.InputValue, out exception);
                                break;

                            case "launchbrowser":
                                actionStatus = actionHelper.ActionLaunchbrowser(browser, driverPath, out exception, out driver);
                                break;

                            case "select":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionSelect(itemAction, named, out exception, driver);
                                break;

                            case "scrollbar":
                                actionStatus = actionHelper.ActionScrollbar(itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "movingpointer":
                                actionStatus = actionHelper.actionMOVINGPOINTER(itemAction.SearchBy, driver, itemAction.Value, out exception);
                                break;

                            case "draganddrop":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionDraganddrop(named, driver, itemAction.Value, out exception);
                                break;

                            case "clickcheckbox":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionCLICKCHECKBOX(itemAction, itemAction.Value, named, driver, out exception);
                                break;

                            case "shopriskoccupancy":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionSHOPOCCUPANCY(itemAction.Value, named, driver, out exception);
                                break;

                            case "employeetabledata":
                                named = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.ActionEmployeeDetailsTable(driver, named, itemAction.Value, employeeIndex, out exception);
                                employeeIndex++;
                                break;

                            case "assertion":
                                Excpected = inputData[itemAction.InputColumnName];
                                errorColumn = itemAction.InputColumnName;
                                actionStatus = actionHelper.actionAssertion(driver, out exception);
                                break;

                            case "compareassertions":
                                if (!string.IsNullOrEmpty(Excpected))
                                {
                                    errorColumn = itemAction.InputColumnName;
                                    actionStatus = actionHelper.ActionCompareAssertions(destinationFile, itemAction.SearchBy, xlInputFilePath, indexRow, xlInputSheetName, itemAction.InputColumnName, itemAction.Value, Excpected, driver, out exception);
                                    assertionStatus = actionStatus;
                                    Excpected = null;
                                }
                                else
                                {
                                    actionStatus = true;
                                    assertionStatus = true;
                                }
                                break;

                            default:
                                actionStatus = true;
                                break;
                        }
                    }
                    //bool driverActive = !driverStatus ? false : actionHelper.IsAlertPresent(driver, out exception);
                    bool driverActive = true;                    
                    driverActive = !driverStatus ? false : actionHelper.AlertCheck(driver);
                    if (!actionStatus || driverActive)
                    {                       
                        if (itemAction.IfErrorOccurs == "Continue Next" || itemAction.IfErrorOccurs == "")
                        {
                            LoggingHelper.insertLog("Action Skipped");
                        }
                        else if (itemAction.IfErrorOccurs == "Stop")
                        {
                            DateTime end = DateTime.Now;
                            TimeSpan timeDifference = end - Start;
                            DataTable dtDataInputPutlatest = excelHelper.GetDataTableFromExcelFile(xlInputFilePath, xlInputSheetName);
                            var jsonDatainput = excelHelper.DataTableToJSONWithJSONNet(dtDataInputPutlatest);
                            if (string.IsNullOrEmpty(destinationFile))
                            {
                                sourceFile = xlInputFilePath;
                                string outputDirectory = OutputFilePath + outputFolder + "\\";
                                if (!Directory.Exists(outputDirectory))
                                {
                                    Directory.CreateDirectory(outputDirectory);
                                }
                                DateTime time = DateTime.Now;
                                filename = System.IO.Path.GetFileName(xlInputFilePath).Replace(".xlsx", "");
                                destinationFile = outputDirectory + filename + "_Output" + time.ToString("h_mm_ss_mm") + ".xlsx";
                                OutputfileName = filename + "_Output" + time.ToString("h_mm_ss_mm") + ".xlsx";
                                File.Copy(sourceFile, destinationFile, true);
                            }

                            if (actionHelper.IsAlertPresent(driver)==true || actionHelper.IsModellocatoralert(driver)==true)
                            {
                                By modalLocator = By.ClassName("bootbox-body");
                                actionHelper.AlertLog(driver, modalLocator, xlInputFilePath, destinationFile, indexRow, xlInputSheetName, errorColumn);
                                actionHelper.WriteDatatoExcel("Result", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, headerStatus == false ? "PASS" : "FAIL");
                                actionHelper.WriteDatatoExcel("Stage", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, headerStatus == false ? "Last Stage" : itemAction.Stage);
                                actionHelper.WriteDatatoExcel("Execution Time", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, timeDifference.ToString());
                                actionHelper.WriteDatatoExcel("AssertionStatus", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, !assertionStatus ? "FAIL" : "PASS");
                            }
                            else
                            {
                                actionHelper.WriteDatatoExcel("Error", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, headerStatus == false?  "No Error" : $"Recheck the data inserted in ({errorColumn}) Input Column");//itemAction.InputColumnName,errorColumn
                                actionHelper.WriteDatatoExcel("Result", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, headerStatus == false ? "PASS" : "FAIL");
                                actionHelper.WriteDatatoExcel("Stage", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, headerStatus == false ? "Last Stage" : itemAction.Stage);
                                actionHelper.WriteDatatoExcel("Execution Time", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, timeDifference.ToString());
                                actionHelper.WriteDatatoExcel("AssertionStatus", xlInputFilePath, destinationFile, indexRow, xlInputSheetName, !assertionStatus ? "FAIL" : "PASS");
                            }
                            DataTable dtDataOutPut = excelHelper.GetDataTableFromExcelFile(destinationFile, xlInputSheetName);                      

                            ExcelHelper excelHelper1 = new ExcelHelper();
                            DataTable dtDataOutPutlatest = excelHelper1.GetDataTableFromExcelFile(destinationFile, xlInputSheetName);
                            var jsonDataoutput = excelHelper1.DataTableToJSONWithJSONNet(dtDataOutPutlatest);
                            recordModel.inputJsonRow = jsonDatainput;
                            recordModel.outputJsonRow = jsonDataoutput;
                            recordModel.status = "Failed";
                            proStatus = "Failed";
                            recordModel.lastStageProcessed = itemAction.Stage;
                            recordModel.lastError = exception==null ? "" : exception;
                            recordModel.OutputfileName = OutputfileName;
                            int recordInsertresult = dal.listofRecordModel(recordModel);

                            //Get newly added Record Data
                            recordInfo = dal.GetrecordDataById(recordInsertresult);

                            //Add Log data
                            LogModel lModel = new LogModel();
                            lModel.logTime = timeDifference.ToString();
                            lModel.description = exception.ToString();
                            lModel.recordId = recordInsertresult;
                            string logs = dal.listofLogModel(lModel);
                            string update = dal.updateLogs(lModel);

                            //LOG THE ERROR                           
                            LoggingHelper.insertLog(" Error in Action: " + exception);
                            LoggingHelper.insertLog("*************** TEST FAILED ***************" + "\n");
                            employeeIndex = 0;

                            //Update File endTime
                            fileInfo.endTime = DateTime.Now.ToString();
                            var fresult = dal.updateFileEndTime(fileInfo);
                            fileInfo.status = "failed";
                            var fstatus = dal.updatefilestatus(fileInfo);
                            fileInfo.outputfilename = OutputfileName;
                            var foutputfile = dal.updateoutputfilename(fileInfo);
                            recordInfo.endTime = DateTime.Now.ToString();
                            var rRecords = dal.updateRecordEndTimeAndStatus(recordInfo);
                            if (keepbrowseropen == "No") driver.Quit();
                            //driver.Quit();
                            if (!string.IsNullOrEmpty(itemAction.WaitTime))
                            {
                                Thread.Sleep(Convert.ToInt32(itemAction.WaitTime));
                            }
                            status = false;
                            break;
                        }
                    }
                    Thread.Sleep(1000);
                    if (!string.IsNullOrEmpty(itemAction.WaitTime))
                    {
                        Thread.Sleep(Convert.ToInt32(itemAction.WaitTime));
                    }
                }
            }
        }
    }
    catch (Exception ex)
    {
        LoggingHelper.insertLog(ex.ToString());
        exception = ex.ToString();
        status = false;
    }
    return status;
}
#endregion