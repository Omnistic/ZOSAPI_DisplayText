using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZOSAPI;
using ZOSAPI.Analysis;
using ZOSAPI.Common;
using System.Text.RegularExpressions;

namespace CSharpUserAnalysisApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Find the installed version of OpticStudio
            bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            // Note -- uncomment the following line to use a custom initialization path
            //bool isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(@"C:\Program Files\OpticStudio\");
            if (isInitialized)
            {
                LogInfo("Found OpticStudio at: " + ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory());
            }
            else
            {
                HandleError("Failed to locate OpticStudio!");
                return;
            }

            BeginUserAnalysis();
        }

        static void BeginUserAnalysis()
        {
            // Create the initial connection class
            ZOSAPI_Connection TheConnection = new ZOSAPI_Connection();

            // Attempt to connect to the existing OpticStudio instance
            IZOSAPI_Application TheApplication = null;
            try
            {
                TheApplication = TheConnection.ConnectToApplication(); // this will throw an exception if not launched from OpticStudio
            }
            catch (Exception ex)
            {
                HandleError(ex.Message);
                return;
            }
            if (TheApplication == null)
            {
                HandleError("An unknown connection error occurred!");
                return;
            }

            // Check the connection status
            if (!TheApplication.IsValidLicenseForAPI)
            {
                HandleError("Failed to connect to OpticStudio: " + TheApplication.LicenseStatus);
                return;
            }

            switch (TheApplication.Mode)
            {
                case ZOSAPI_Mode.UserAnalysis:
                    RunUserAnalysis(TheApplication);
                    break;
                case ZOSAPI_Mode.UserAnalysisSettings:
                    ShowUserAnalysisSettings(TheApplication);
                    break;
                default:
                    HandleError("User plugin was started in the wrong mode: expected UserAnalysis, found " + TheApplication.Mode.ToString());
                    return;
            }

            // Clean up
            FinishUserAnalysis(TheApplication);
        }

        static void RunUserAnalysis(IZOSAPI_Application TheApplication)
        {
            IOpticalSystem TheSystem = TheApplication.PrimarySystem;

            IUserAnalysisData TheAnalysisData = TheApplication.UserAnalysisData;
            ISettingsData TheSettings = TheAnalysisData.UserSettings;

            // Add your custom code here...
            TheAnalysisData.WindowTitle = "RSLM Results";

            string fileFullPath = TheSystem.SystemFile;
            string fileDir = Path.GetDirectoryName(fileFullPath);
            string fileName = Path.GetFileNameWithoutExtension(fileFullPath);
            string logPath = Path.Combine(fileDir, fileName + "_RSLM.TXT");

            string LSMResults = "";
            try
            {
                LSMResults = File.ReadAllText(logPath);
            }
            catch
            {
                MessageBox.Show("RSLM result file not found. Please run the RSLM user extension before using this analysis.", "File not found",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            
            LSMResults = Regex.Replace(LSMResults, @"(\s*\r\n){2,}", "\n\n");
            LSMResults = Regex.Replace(LSMResults, @"(\s*\r\n){1,}", "\n");

            IUserTextData TheTextData = TheApplication.UserAnalysisData.MakeText();
            TheTextData.Data = LSMResults;

            // Use TheAnalysisData to create a specific plot type and populate the data
            //IUser2DLineData linePlot = TheAnalysisData.Make2DLinePlot("New 2D Line Plot", 1, new double[] { 1, 2, 3 });
            //linePlot.AddSeries(...);
        }

        static void ShowUserAnalysisSettings(IZOSAPI_Application TheApplication)
        {
            IOpticalSystem TheSystem = TheApplication.PrimarySystem;

            IUserAnalysisData TheAnalysisData = TheApplication.UserAnalysisData;
            ISettingsData TheSettings = TheAnalysisData.UserSettings;

            // TODO - retrieve the settings specific to your analysis here

            // This will show a form to modify your settings (currently blank)...
            AnalysisSettingsForm SettingsForm = new AnalysisSettingsForm();
            // Add your custom code here, and to the SettingsForm...
            System.Windows.Forms.Application.Run(SettingsForm);            

            // TODO - write settings back to TheSettings
        }

        static void FinishUserAnalysis(IZOSAPI_Application TheApplication)
        {
            // Note - OpticStudio will wait for the operand to complete until this application exits 
        }

        static void LogInfo(string message)
        {
            // TODO - add custom logging
            Console.WriteLine(message);
        }

        static void HandleError(string errorMessage)
        {
            // TODO - add custom error handling
            throw new Exception(errorMessage);
        }

    }
}
