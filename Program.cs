using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.Xml.Linq;

using databaseAPI;
using GNAgeneraltools;
using GNAsurveytools;
using GNAspreadsheettools;
using OfficeOpenXml;
using Twilio.Rest.Api.V2010.Account;
using System.Diagnostics.Metrics;
using EASendMail;
using System.Data.SqlTypes;

namespace settopOutOfToleranceAlarm
{
    // 20240522 Software created 
    //

    internal class Program
    {
        static void Main(string[] args)
        {

#pragma warning disable CS0164
#pragma warning disable CS8061
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable IDE0059


            gnaTools gnaT = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();
            GNAsurveycalcs gnaSurvey = new();

            string strProgramStart = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            //==== Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;

            //==== Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            //==== System config variables

            string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;
            string strProjectTitle = ConfigurationManager.AppSettings["ProjectTitle"];
            string strContractTitle = ConfigurationManager.AppSettings["ContractTitle"];
            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];
            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strSurveyWorksheet = ConfigurationManager.AppSettings["SurveyWorksheet"];
            string strSMSTitle = ConfigurationManager.AppSettings["SMSTitle"];

            // allocate the sms mobile numbers
            string[] settopDataFolder = new string[10];
            settopDataFolder[1] = ConfigurationManager.AppSettings["SettopDataFolder1"];
            settopDataFolder[2] = ConfigurationManager.AppSettings["SettopDataFolder2"];
            settopDataFolder[3] = ConfigurationManager.AppSettings["SettopDataFolder3"];
            settopDataFolder[4] = ConfigurationManager.AppSettings["SettopDataFolder4"];
            settopDataFolder[5] = ConfigurationManager.AppSettings["SettopDataFolder5"];
            settopDataFolder[6] = ConfigurationManager.AppSettings["SettopDataFolder6"];
            settopDataFolder[7] = ConfigurationManager.AppSettings["SettopDataFolder7"];
            settopDataFolder[8] = ConfigurationManager.AppSettings["SettopDataFolder8"];
            settopDataFolder[9] = ConfigurationManager.AppSettings["SettopDataFolder9"];

            // allocate the sms mobile numbers
            string[] smsMobile = new string[10];
            smsMobile[1] = ConfigurationManager.AppSettings["RecipientPhone1"];
            smsMobile[2] = ConfigurationManager.AppSettings["RecipientPhone2"];
            smsMobile[3] = ConfigurationManager.AppSettings["RecipientPhone3"];
            smsMobile[4] = ConfigurationManager.AppSettings["RecipientPhone4"];
            smsMobile[5] = ConfigurationManager.AppSettings["RecipientPhone5"];
            smsMobile[6] = ConfigurationManager.AppSettings["RecipientPhone6"];
            smsMobile[7] = ConfigurationManager.AppSettings["RecipientPhone7"];
            smsMobile[8] = ConfigurationManager.AppSettings["RecipientPhone8"];
            smsMobile[9] = ConfigurationManager.AppSettings["RecipientPhone9"];



            string strActivityFolder = ConfigurationManager.AppSettings["SystemStatusFolder"];
            string strAlarmsFolder = ConfigurationManager.AppSettings["SystemAlarmsFolder"];

            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];
            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string strEmailRecipients = ConfigurationManager.AppSettings["EmailRecipients"];

            //==== Main program
            gnaT.WelcomeMessage("settopOutOfToleranceAlarm 20240522");
            string strSoftwareLicenseTag = "SETTOP";
            _ = gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);

            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;

            string strNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("     Project: " + strProjectTitle);
            Console.WriteLine("     Master workbook: " + strMasterWorkbookFullPath);

            string strAlarmFile = "settopOutOfToleranceAlarm.txt";
            string strSystemActivityFile = "SystemActivityLog.txt";
            string strAlarmLog = strAlarmsFolder + strAlarmFile;
            string strActivityLog = strActivityFolder + strSystemActivityFile;
            string strLatestFile = "latestFile";
            string strAnswer = "";

            if (strFreezeScreen == "Yes")
            {
                gnaDBAPI.testDBconnection(strDBconnection);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                gnaT.checkFileExistance(strAlarmsFolder, strAlarmFile);
                Console.WriteLine("     " + strAlarmLog + " ok");
                gnaT.checkFileExistance(strActivityFolder, strSystemActivityFile);
                Console.WriteLine("     " + strActivityLog + " ok");
                string strAction = "Continue";
                for (int j = 1; j <= 9; j++)
                {
                    if (settopDataFolder[j] != "None")
                    {
                        strAnswer = gnaT.doesFileExist(settopDataFolder[j], strLatestFile);

                        if (strAnswer == "Yes")
                        {
                            Console.WriteLine("     " + settopDataFolder[j] + strLatestFile + " ok");
                        }
                        else
                        {
                            Console.WriteLine("     " + settopDataFolder[j] + strLatestFile + " MISSING");
                            strAction = "Exit";
                        }
                    }
                }

                if (strAction == "Exit")
                {
                    Console.WriteLine("\n     Fix the missing folders");
                    Console.WriteLine("     press key to exit");
                    Console.ReadLine();
                    goto ThatsAllFolks;
                }
            }
            else
            {
                Console.WriteLine("     System environment not checked");
            }










ThatsAllFolks:

            Console.WriteLine("\nsettopOutOfToleranceAlarm checking completed...");
            gnaT.freezeScreen(strFreezeScreen);

            Environment.Exit(0);


        }
    }
}