
using System.Data.SqlClient;
using TestAutomationSuite.dbModels;
using System.Configuration;


namespace TestAutomationSuite.Utility
{
    public class DAL
    {
        string connectionstring = null;
        public DAL()
        {
            ErrorLog errorLog = new ErrorLog();
            ActionHelper actionHelper = new ActionHelper();
            connectionstring = ConfigurationManager.ConnectionStrings["TestAuto"].ConnectionString;
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionstring))
                {
                    connection.Open();
                    LoggingHelper logging = new LoggingHelper();
                    logging.insertLog("DataBase Connection Established successfully");
                    // You can perform additional validation or queries here if needed
                }
            }
            catch (Exception ex)
            {
                LoggingHelper logging = new LoggingHelper();
                logging.insertLog("Failed to Establish Data Base connection, please check the Connection String or Database name that you have provided in App.config file");
                Environment.Exit(1);
            }
        }
        //Method to insert data in file table of db
        public FileModels GetfileDataById(int id)
        {
            FileModels fileModels = new FileModels();
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "SELECT * FROM DataFiles WHERE fileId = @id";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            fileModels = new FileModels
                            {
                                fileId = Convert.ToInt32(reader["fileId"]),
                                processId = Convert.ToInt32(reader["processId"]),
                                projectId = Convert.ToInt32(reader["projectId"]),
                                fileName = reader["fileName"].ToString(),
                                startTime = reader["startTime"].ToString(),
                                endTime = reader["endTime"].ToString(),
                                status = reader["status"].ToString()
                            };
                        }
                    }
                }
            }
            return fileModels;
        }


        //Method to insert data in process table of db
        public ProcessModel GetprocessDataById(int id)
        {
            ProcessModel processModel = new ProcessModel();
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "SELECT * FROM Process WHERE processId = @id";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            processModel = new ProcessModel
                            {
                                processId = Convert.ToInt32(reader["processId"]),
                                winProcID = Convert.ToInt32(reader["winProcID"]),
                                startTime = reader["startTime"].ToString(),
                                endTime = reader["endTime"].ToString(),
                                processStatus = reader["processStatus"].ToString()
                            };
                        }
                    }
                }
            }
            return processModel;
        }


        //Method to insert data in record table of db
        public RecordModel GetrecordDataById(int id)
        {
            RecordModel recordModel = new RecordModel();
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "SELECT * FROM Records WHERE recordId = @id";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            recordModel = new RecordModel
                            {
                                recordId = Convert.ToInt32(reader["recordId"]),
                                processId = Convert.ToInt32(reader["processId"]),
                                fileId = Convert.ToInt32(reader["fileId"]),
                                inputJsonRow = reader["inputJsonRow"].ToString(),
                                startTime = reader["startTime"].ToString(),
                                endTime = reader["endTime"].ToString(),
                                outputJsonRow = reader["outputJsonRow"].ToString(),
                                status = reader["status"].ToString(),
                                lastStageProcessed = reader["lastStageProcessed"].ToString(),
                                lastError = reader["lastError"].ToString(),
                                OutputfileName = reader["OutputfileName"].ToString()
                            };
                        }
                    }
                }
            }
            return recordModel;
        }

        //Method to update data in process table of db
        public string updateProcessEndTimeAndStatus(ProcessModel process)
        {
            string result = "success";
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "UPDATE Process SET endTime = @endTime, processStatus = @processStatus WHERE processId = @processId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@endTime", process.endTime);
                    cmd.Parameters.AddWithValue("@processStatus", process.processStatus);
                    cmd.Parameters.AddWithValue("@processId", process.processId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }


        //Method to update data in file table of db
        public string updateFileEndTime(FileModels file)
        {
            string result = "success";
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "UPDATE DataFiles SET endTime = @endTime WHERE fileId = @fileId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@endTime", file.endTime);
                    cmd.Parameters.AddWithValue("@fileId", file.fileId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }


        public string updatefilestatus(FileModels file)
        {
            string result = "success";
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "UPDATE DataFiles SET status = @status WHERE fileId = @fileId";

                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@status", file.status);
                    cmd.Parameters.AddWithValue("@fileId", file.fileId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }

        //method to update outputfilename
        public string updateoutputfilename(FileModels file)
        {
            string result = "success";
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "UPDATE DataFiles SET OutputFileName = @outputFileName WHERE fileId = @fileId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@outputFileName", file.outputfilename);
                    cmd.Parameters.AddWithValue("@fileId", file.fileId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }


        //Method to update data in log table of db
        public string updateLogs(LogModel logs)
        {
            string result = "success";
            using (SqlConnection cnn = new SqlConnection(connectionstring))
            {
                cnn.Open();

                string query = "UPDATE Logs SET description = @Description WHERE logId = @LogId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@Description", logs.description);
                    cmd.Parameters.AddWithValue("@LogId", logs.logId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }

        //Method to update data in record table of db
        public string updateRecordEndTimeAndStatus(RecordModel record)
        {
            string result = "success";
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "UPDATE Records SET endTime = @endTime WHERE recordId = @recordId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@endTime", record.endTime);
                    cmd.Parameters.AddWithValue("@recordId", record.recordId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }


        //Method to update data in record table of db
        public string updateRecordStartTimeAndStatus(RecordModel record)
        {
            string result = "success";
            if (record.startTime == null)
            {
                return "Error: startTime cannot be null.";
            }
            using (SqlConnection cnn = new SqlConnection(connectionstring))
            {
                cnn.Open();
                string query = "UPDATE Records SET startTime = @startTime WHERE recordId = @recordId";

                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@startTime", record.startTime);
                    cmd.Parameters.AddWithValue("@recordId", record.recordId);
                    cmd.ExecuteNonQuery();
                }
            }
            return result;
        }



        //Method to insert data in project manager table of db
        public int InsertProject(ProjectModel project)
        {
            int result = 0;
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "INSERT INTO ProjectManager (projectName, actionFileJson, inputFilePath, projectStatus) OUTPUT INSERTED.projectId VALUES (@projectName, @actionFileJson, @inputFilePath, @projectStatus)";
                using (SqlCommand command = new SqlCommand(query, cnn))
                {
                    command.Parameters.AddWithValue("@projectName", project.projectName);
                    command.Parameters.AddWithValue("@actionFileJson", project.actionFileJson);
                    command.Parameters.AddWithValue("@inputFilePath", project.inputFilePath);
                    command.Parameters.AddWithValue("@projectStatus", project.projectStatus);
                    result = (int)command.ExecuteScalar();
                }
            }
            return result;
        }

        //Method to insert data in file table of db
        public int InsertFile(FileModels project)
        {
            int result = 0;
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "INSERT INTO DataFiles (processId, projectId, fileName, startTime, endTime, status) OUTPUT INSERTED.fileId VALUES (@processId, @projectId, @fileName, @startTime, @endTime, @status)";
                using (SqlCommand command = new SqlCommand(query, cnn))
                {
                    command.Parameters.AddWithValue("@processId", project.processId);
                    command.Parameters.AddWithValue("@projectId", project.projectId);
                    command.Parameters.AddWithValue("@fileName", project.fileName);
                    command.Parameters.AddWithValue("@startTime", project.startTime);
                    command.Parameters.AddWithValue("@endTime", project.endTime);
                    command.Parameters.AddWithValue("@status", project.status);

                    result = (int)command.ExecuteScalar();
                }
            }
            return result;
        }

        //Method to insert data in log table of db
        public string listofLogModel(LogModel log)
        {
            string result = "success";
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();

                string query = "INSERT INTO Logs (recordId, logTime, description) VALUES (@recordId, @logTime, @description)";

                using (SqlCommand command = new SqlCommand(query, cnn))
                {
                    command.Parameters.AddWithValue("@recordId", log.recordId);
                    command.Parameters.AddWithValue("@logTime", log.logTime);
                    command.Parameters.AddWithValue("@description", log.description);
                    command.ExecuteNonQuery();
                }
            }
            return result;
        }


        //Method to insert data in Process table of db
        public int InsertProcess(ProcessModel project)
        {
            int result = 0;
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "INSERT INTO Process (winProcID, startTime, endTime, processStatus) OUTPUT INSERTED.processId VALUES (@winProcID, @startTime, @endTime, @processStatus)";
                using (SqlCommand command = new SqlCommand(query, cnn))
                {
                    command.Parameters.AddWithValue("@winProcID", project.winProcID);
                    command.Parameters.AddWithValue("@startTime", project.startTime);
                    command.Parameters.AddWithValue("@endTime", project.endTime);
                    command.Parameters.AddWithValue("@processStatus", project.processStatus);
                    result = (int)command.ExecuteScalar();
                }
            }
            return result;
        }


        //Method to insert data in record table of db
        public int listofRecordModel(RecordModel project)
        {
            int result = 0;
            string connect = connectionstring;
            using (SqlConnection cnn = new SqlConnection(connect))
            {
                cnn.Open();
                string query = "INSERT INTO Records (processId, fileId, inputJsonRow, startTime, endTime, outputJsonRow, status, lastStageProcessed, lastError, OutputfileName) OUTPUT INSERTED.recordId VALUES (@processId, @fileId, @inputJsonRow, @startTime, @endTime, @outputJsonRow, @status, @lastStageProcessed, @lastError, @OutputfileName)";
                using (SqlCommand command = new SqlCommand(query, cnn))
                {
                    command.Parameters.AddWithValue("@processId", project.processId);
                    command.Parameters.AddWithValue("@fileId", project.fileId);
                    command.Parameters.AddWithValue("@inputJsonRow", project.inputJsonRow);
                    command.Parameters.AddWithValue("@startTime", project.startTime);
                    command.Parameters.AddWithValue("@endTime", project.endTime);
                    command.Parameters.AddWithValue("@outputJsonRow", project.outputJsonRow);
                    command.Parameters.AddWithValue("@status", project.status);
                    command.Parameters.AddWithValue("@lastStageProcessed", project.lastStageProcessed);
                    command.Parameters.AddWithValue("@lastError", project.lastError);
                    command.Parameters.AddWithValue("@OutputfileName", project.OutputfileName);
                    result = (int)command.ExecuteScalar();
                }
            }
            return result;
        }

        public bool getRecordstatus(int processId)
        {
            bool result = false;
            using (SqlConnection cnn = new SqlConnection(connectionstring))
            {
                cnn.Open();
                string query = "SELECT status FROM Records WHERE processId = @processId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@processId", processId);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        List<string> statusList = new List<string>();
                        while (reader.Read())
                        {
                            string status = reader["status"].ToString();
                            statusList.Add(status);
                        }
                        result = statusList.Any(x => x.Equals("Passed"));
                    }
                }
            }

            return result;
        }


        //Get Data from Sql db
        public ProjectModel GetDataProjectDataById(int projectId)
        {
            ProjectModel projectModel = null;
            using (SqlConnection cnn = new SqlConnection(connectionstring))
            {
                cnn.Open();
                string query = "SELECT * FROM ProjectManager WHERE projectName IS NOT NULL AND actionFileJson IS NOT NULL AND projectId = @projectId";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    cmd.Parameters.AddWithValue("@projectId", projectId);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            projectModel = new ProjectModel
                            {
                                projectId = Convert.ToInt32(reader["projectId"]),
                                projectName = reader["projectName"].ToString(),
                                actionFileJson = reader["actionFileJson"].ToString(),
                                inputFilePath = reader["inputFilePath"].ToString(),
                                projectStatus = reader["projectStatus"].ToString()
                            };
                        }
                    }
                }
            }
            return projectModel;
        }

        public List<ProjectModel> GetAllProjects()
        {
            List<ProjectModel> projectModelList = new List<ProjectModel>();
            using (SqlConnection cnn = new SqlConnection(connectionstring))
            {
                cnn.Open();
                string query = "SELECT * FROM ProjectManager WHERE projectName IS NOT NULL AND actionFileJson IS NOT NULL";
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ProjectModel project = new ProjectModel
                        {
                            projectId = Convert.ToInt32(reader["projectId"]),
                            projectName = reader["projectName"].ToString(),
                            actionFileJson = reader["actionFileJson"].ToString(),
                            inputFilePath = reader["inputFilePath"].ToString(),
                            projectStatus = reader["projectStatus"].ToString()
                        };
                        projectModelList.Add(project);
                    }
                }
            }
            return projectModelList;
        }

    }
}
