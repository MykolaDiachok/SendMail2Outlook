using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using NDesk.Options;
using NLog;
using NLog.Config;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;
using System.Timers;
using System.Diagnostics;
using System.Management;


namespace SendMailOutlook
{
    class Program
    {
        private static string jsonfile;
        private static bool killOutlook=false;
        private static Timer _Timer;
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static int returncode = (int)ExitCode.Success;
        //private static Office.CommandBarButton EncButton = null;
        //private static Office.CommandBarButton DigButton = null;

        static int Main(string[] args)
        {
                     
            new OptionSet()
              .Add("f=|file=","set json file", f => jsonfile = f)
              .Add("k","if error kill outlook proc",k=> killOutlook=true)
              .Add("?|h|help","Show this help", h => DisplayHelp())
              .Parse(args);

            if (jsonfile == null)
            {
                DisplayHelp();
            #if DEBUG
                Console.WriteLine("Press any key!");
                Console.ReadKey();
                #endif
                return (int)ExitCode.InvalidFilename;
            }

            

            ForOutLook tJSOn = JsonConvert.DeserializeObject<ForOutLook>(File.ReadAllText(jsonfile));
            try
            {
                Microsoft.Office.Interop.Outlook.Application olkApp1 =
                    new Microsoft.Office.Interop.Outlook.Application();
                Outlook.Accounts accounts = olkApp1.Session.Accounts;
                foreach (Outlook.Account account in accounts)
                {
                    logger.Info("Send mail from:{0}", account.SmtpAddress);
                    SendEmailFromAccount(olkApp1, tJSOn, account.SmtpAddress);
                    break; //отправляем только с 1 аккаунта
                           //Console.WriteLine("smtp:{0}", account.SmtpAddress);
                }
            }
            catch (COMException ex)
            {
                logger.Error(ex);
                returncode = (int)ExitCode.OutlookRunProblem;
            }

            #if DEBUG
            logger.Info("return:{0}", returncode);
            Console.WriteLine("Press any key!");
            Console.ReadKey();
            
            #endif
            return returncode;
        }


        private static void DisplayHelp()
        {
            Console.WriteLine("====================HELP==========================");
            Console.WriteLine("   -f' or '-file' Reference json file");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Sample struct of file:");
            Console.WriteLine("Struct of file:");
            Console.WriteLine(" {");
            Console.WriteLine("     \"subject\": \"Mail subject\",");
            Console.WriteLine("     \"body\": \"Text bode or HTML body mail\",");
            Console.WriteLine("     \"To\": [");
            Console.WriteLine("         \"mail1@site\",");
            Console.WriteLine("         \"mailN@site\"");
            Console.WriteLine("           ],");
            Console.WriteLine("     \"Attachments\": [");
            Console.WriteLine("     {");
            Console.WriteLine("         \"filename\": \"NameOfFile.Type\",");
            Console.WriteLine("         \"Base64\": \"Data file in Base64\",");
            Console.WriteLine("     }");
            Console.WriteLine("     ]");
            Console.WriteLine(" }");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("return info:");
            Console.WriteLine("     0 - Success");
            Console.WriteLine("     2 - InvalidFilename");
            Console.WriteLine("     3 - EncryptionProblems");
            Console.WriteLine("     4 - OutlookRunProblem");
            Console.WriteLine("     10 - UnknownError");
            
        }

        private static void addOutlookEncryption(ref Outlook.MailItem mItem)
        {
            CommandBarButton encryptBtn;
            mItem.Display(false);
            encryptBtn = mItem.GetInspector.CommandBars.FindControl(MsoControlType.msoControlButton, 718, Type.Missing, Type.Missing) as CommandBarButton;
            if (encryptBtn == null)
            {
                //if it's null, then add the encryption button
                encryptBtn = (CommandBarButton)mItem.GetInspector.CommandBars["Standard"].Controls.Add(Type.Missing, 718, Type.Missing, Type.Missing, true);
            }
            if (encryptBtn.Enabled)
            {
                if (encryptBtn.State == MsoButtonState.msoButtonUp)
                {
                    encryptBtn.Execute();
                }
            }
            mItem.Close(Outlook.OlInspectorClose.olDiscard);
        }


        private static void SendEmailFromAccount(Outlook.Application application, ForOutLook inJSON, string smtpAddress)
        {

            // Create a new MailItem and set the To, Subject, and Body properties. 
            Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem);
            //newMail.t
            Microsoft.Office.Interop.Outlook.Recipients oRecips = (Microsoft.Office.Interop.Outlook.Recipients)newMail.Recipients;

            //EncButton = (Office.CommandBarButton)newMail.GetInspector.CommandBars.FindControl(Office.MsoControlType.msoControlButton, 718, Type.Missing, true);
            //if (EncButton.State == Office.MsoButtonState.msoButtonDown)
            //{
            //    logger.Info("Mail will be Encrypted");
            //}
            //DigButton = (Office.CommandBarButton)newMail.GetInspector.CommandBars.FindControl(Office.MsoControlType.msoControlButton, 719, Type.Missing, true);
            //if (DigButton.State == Office.MsoButtonState.msoButtonDown)
            //{
            //    logger.Info("Mail will be Digitally signed");
            //}

            logger.Info("Subject:{0}", inJSON.subject);
            newMail.Subject = inJSON.subject;
            logger.Info("Body:{0}", inJSON.body);
            newMail.Body = inJSON.body;
            foreach (var recipient in inJSON.TO)
            {
                logger.Info("recipient:{0}", recipient);
                Microsoft.Office.Interop.Outlook.Recipient oRecip = (Microsoft.Office.Interop.Outlook.Recipient)oRecips.Add(recipient);
                oRecip.Resolve();
            }
            

            foreach (ForOutLookAttachments attachment in inJSON.Attachments)
            {
                logger.Info("attachment:{0}", attachment.filename);
                Byte[] bytes = Convert.FromBase64String(attachment.Base64);
                string fileName = System.IO.Path.GetTempPath() + attachment.filename;
                File.WriteAllBytes(fileName, bytes);
                //newMail.Attachments.Add(fileName, "file.xlsx");
                var tO = newMail.Attachments.Add(fileName, Outlook.OlAttachmentType.olByValue, 1, "file.xlsx");
                
            }

            //addOutlookEncryption(ref newMail);
            // Retrieve the account that has the specific SMTP address. 
            Outlook.Account account = GetAccountForEmailAddress(application, smtpAddress);
            // Use this account to send the e-mail. 
            newMail.SendUsingAccount = account;
            _Timer = new Timer();
            _Timer.Elapsed += new ElapsedEventHandler(DisplayTimeEvent2);
            _Timer.Interval = 1000;
            _Timer.Start();
            try
            {
                newMail.Send();
            }
            catch(COMException ex)
            {
                logger.Error(ex);
            }
            _Timer.Stop();



            logger.Info("Job end");

        }


        const UInt32 WM_KEYDOWN = 0x0100;
        const int VK_F5 = 0x74;
        const int VK_ESCAPE = 0x1B;

        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, UInt32 Msg, int wParam, int lParam);

        private static void DisplayTimeEvent(object source, ElapsedEventArgs e)
        {
            _Timer.Stop();
            Process[] processes = Process.GetProcessesByName("OUTLOOK");

            foreach (Process proc in processes)
            {
                if ((proc.MainWindowTitle == "Encryption Problems")
                    ||(proc.MainWindowTitle == "Неполадки шифрования")
                    ||(proc.MainWindowTitle == "Неполадки шифрування")
                    )
                {
                    if (killOutlook)
                    {
                        proc.Kill();
                        logger.Trace("kill process \"Outlook\" id:{0}", proc.Id);
                    }
                    else
                    {
                        
                        PostMessage(proc.MainWindowHandle, WM_KEYDOWN, VK_ESCAPE, 0); //it's work
                        logger.Trace("Send key \"ESC\"");
                    }


                    //PostMessage(proc.MainWindowHandle, WM_KEYDOWN, KEY_MENU, 0); //it's not work
                    //PostMessage(proc.MainWindowHandle, WM_KEYDOWN, KEY_N, 0);
                    returncode = (int)ExitCode.EncryptionProblems;
                }
            }
        }


        private static void DisplayTimeEvent2(object source, ElapsedEventArgs e)
        {
            _Timer.Stop();
            
            ManagementScope scope = new ManagementScope(@"\\.\root\cimv2");
            scope.Connect();

            string query = string.Format("SELECT * FROM Win32_Process WHERE Name='{0}'", "OUTLOOK.EXE");
            ManagementObjectSearcher searcher =
                new ManagementObjectSearcher(query);
            foreach (ManagementObject obj in searcher.Get())
            {
                uint processId = (uint)obj["ProcessId"];
                Process process = null;
                try
                {
                    process = Process.GetProcessById((int)processId);
                }
                catch (ArgumentException ex)
                {
                    logger.Error(ex);
                }
                try
                {
                    if (process != null)
                    {
                        process.Kill();
                    }
                }
               
                catch (InvalidOperationException ex)
                {
                    logger.Error(ex);
                }
            }
            returncode = (int)ExitCode.EncryptionProblems;
        }




        public static Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress)
        {

            // Loop over the Accounts collection of the current Outlook session. 
            Outlook.Accounts accounts = application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                // When the e-mail address matches, return the account. 
                if (account.SmtpAddress == smtpAddress)
                {
                    return account;
                }
            }
            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", smtpAddress));
        }

        enum ExitCode : int
        {
            Success = 0,            
            InvalidFilename = 2,
            EncryptionProblems = 3,
            OutlookRunProblem = 4,
            UnknownError = 10
        }

    }
}
