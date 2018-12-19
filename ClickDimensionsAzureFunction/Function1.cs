using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;


namespace ClickDimensionsUptimeMonitoring
{
    public static class Function2
    {
        /// SMTP Email strings
        /// Ensure that this configuration has been checked prior to implementing into production
        private static string FROM = "";
        private static string FROMNAME = "CRM Online User";
        //  private static string TO = "";
        private static string TO = "";
        private static string SMTP_USERNAME = "";
        private static string SMTP_PASSWORD = "";
        private static string CONFIGSET = "ConfigSet";
        private static string HOST = "smtp.office365.com";
        private static int PORT = 587;

        /// Connection strings for CRM && for Exchange Service
        /// For testing purposes, this code may reflect connection strings for Mural Dynamics 365 CRM
        /// Ensure that this configuration has been checked prior to implementing into production
        private static string username = "";
        private static string password = "";
        // private static string url = "";
        private static string url = "";


        /// These are the Ids for the mailbox that the email will be routed to
        /// For testing purposes, this code may reflect connection strings for testing mailboxes
        /// Ensure that this configuration has been checked prior to implementing into production
        /// Production Mailbox
        private static string ClickDBox = "";
        /// Local Testing Mailbox
        // private static string ClickDBox2 = "";

        /// For reporting purposes
        private static string str;
        private static int isClickDUp = 0;

        /// _workflowId is the Guid for the workflow that sends the ClickD email
        /// For testing purposes, this code may reflect workflow Guid in Mural Dynamics 365 CRM Mural Sandbox Instance
        /// Ensure that this configuration has been checked prior to implementing into production
        private static Guid _workflowId2 = new Guid(""); // Production
        // private static Guid _workflowId = new Guid(""); // Sandbox


        /// _entityId is the Guid for the contact that the ClickD email is sent to
        /// For testing purposes, this code may reflect contact Guid in Mural Dynamics 365 CRM Mural Sandbox Instance
        /// Ensure that this configuration has been checked prior to implementing into production
        private static Guid _entityId2 = new Guid(""); // Production
        // private static Guid _entityId = new Guid(""); // Sandbox


        [FunctionName("Function2")]

        /// CRON interval is set to run every 30 minutes
        /// The Azure function takes 7 minutes to run
        /// ClickDimensions allows us to send up to 200K emails per year
        /// With this interval; approximately 17,500 +/- emails will be sent per year
        public static void Run([TimerTrigger("0 */30 * * * *")]TimerInfo myTimer, TraceWriter log)
        {

            log.Info($"C# Timer trigger function executed at: {DateTime.Now}");


            /// Connecting to Exchange Server for CRM Online Service User
            /// This function will mark the last message in the ClickDBox as READ
            /// This will prevent any conflict with previously ran functions
            ExchangeService _service1;
            {
                try
                {
                    log.Info("Registering Initial Exchange Connection");
                    _service1 = new ExchangeService
                    {
                        Credentials = new WebCredentials(username, password)
                    };
                }

                catch
                {
                    log.Info("New ExchangeService failed to connect.");
                    return;
                }

                _service1.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");


                SearchFilter sf = new SearchFilter.SearchFilterCollection(
                        LogicalOperator.And, new SearchFilter.IsEqualTo(ItemSchema.Subject, "ClickDimensions Uptime Monitoring"));

                foreach (EmailMessage email in _service1.FindItems(ClickDBox, sf, new ItemView(1)))
                {
                    if (!email.IsRead)
                    {
                        log.Info("Old Unread Messages Present");

                        email.IsRead = true;
                        email.Update(ConflictResolutionMode.AutoResolve);
                        log.Info("MarkedAsRead");

                    }

                    else
                    {
                        if (email.IsRead)
                        {
                            log.Info("No New Emails");

                        }


                    }


                }

            }



            /// Connection to CRM
            /// For testing purposes, this code may reflect setup in Mural Dynamics 365 CRM Mural Sandbox Instance
            /// Ensure that this configuration has been checked prior to implementing into production
            IServiceManagement<IOrganizationService> orgServiceManagement = ServiceConfigurationFactory.
                CreateManagement<IOrganizationService>(new Uri(url));
            AuthenticationCredentials authCredentials = new AuthenticationCredentials();
            authCredentials.ClientCredentials.UserName.UserName = username;
            authCredentials.ClientCredentials.UserName.Password = password;
            AuthenticationCredentials tokenCredentials = orgServiceManagement.Authenticate(authCredentials);
            OrganizationServiceProxy organizationProxy = new OrganizationServiceProxy
                (orgServiceManagement, tokenCredentials.SecurityTokenResponse);


            log.Info("Connected To CRM");


            /// Run pre-built workflow in CRM
            /// This workflow will send a ClickDimensions Email to a contact
            /// Contact will be the CRM Online User
            /// For testing purposes, this code may reflect setup in Mural Dynamics 365 CRM Mural Sandbox Instance
            /// Ensure that this configuration has been checked prior to implementing into production
            ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
            {
                /// For testing purposes only
                // WorkflowId = _workflowId,
                // EntityId = _entityId,

                WorkflowId = _workflowId2,
                EntityId = _entityId2,

            };

            ExecuteWorkflowResponse response = (ExecuteWorkflowResponse)organizationProxy.Execute(request);


            if (request == null)
            {
                log.Info("Workflow Failed");
            }

            else
            {
                log.Info("Workflow Triggered");
            }



            /// Put this process to sleep for 7 minutes
            /// This will give ClickDimensions enough time for the email to be delivered
            /// After 7 minutes, the process will check to see if email was delivered
            /// Ensure that this configuration has been checked prior to implementing into production
            {
                for (int i = 0; i < 1; i++)
                {
                    log.Info("Sleep for 7 minutes.");
                    Thread.Sleep(420000);

                }
                log.Info("Wake");
            }


            /// Connect to CRM Online Service User Email
            ExchangeService _service2;
            {
                try
                {
                    log.Info("Registering Exchange Connection");
                    _service2 = new ExchangeService
                    {

                        /// Connection credentials for CRM Online User to access email 
                        /// For testing purposes, this code may reflect process owners Mural Office 365 credentials
                        /// Ensure that this configuration has been checked prior to implementing into production
                        Credentials = new WebCredentials(username, password)

                    };
                }

                catch
                {

                    log.Info("New ExchangeService failed to connect.");
                    return;

                }

                _service2.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

                try
                {


                    /*
                      
                    /// This will search the Mailbox for the top level folders and display their Ids
                    /// Use this function to obtain the Id necessary for this code                  

                    /// Get all the folders in the message's root folder.
                    Folder rootfolder = Folder.Bind(_service2, WellKnownFolderName.MsgFolderRoot);

                    log.Info("The " + rootfolder.DisplayName + " has " + rootfolder.ChildFolderCount + " child folders.");

                    /// A GetFolder operation has been performed.
                    /// Now display each folder's name and ID.
                    rootfolder.Load();

                    foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))
                    {
                        log.Info("\nName: " + folder.DisplayName + "\n  Id: " + folder.Id);
                    }

                    */


                    log.Info("Reading mail");


                    /// HTML email string for ClickDIsDown Email
                    string HTML = @"<font size=6><font color=red> ClickDimensions Uptime Monitoring <font size=3><font color=black><br /><br /><br />
                                                The ClickDimensions emailing system may not be working at this time or has become slow. Please check the integration and 
                                                    alert the appropriate teams if necessary.<br /> <br /> Regards,<br /><br /> Business & Product Development Team <br /> 
                                            Email: productplatform@mural365.com <br /> Mural Consulting <br /> https://mural365.com <br /><br />";



                    /// This will view the last email in the mailbox
                    /// Search for "ClickD Uptime Monitoring System" in Subject line
                    SearchFilter sf = new SearchFilter.SearchFilterCollection(
                        LogicalOperator.And, new SearchFilter.IsEqualTo(ItemSchema.Subject, "ClickDimensions Uptime Monitoring"));

                    foreach (EmailMessage email in _service2.FindItems(ClickDBox, sf, new ItemView(1)))
                    {

                        /// If the email was sent (unread email = ClickD is working)
                        /// Mark the email as read
                        /// This folder SHOULD NOT contain any UNREAD emails
                        if (!email.IsRead)
                        {
                            log.Info("Email: " + email.Subject);
                            log.Info("Email Sent: " + email.DateTimeSent + " Email Received: " +
                                email.DateTimeReceived + " Email Created: " + email.DateTimeCreated);


                            email.IsRead = true;
                            email.Update(ConflictResolutionMode.AutoResolve);
                            log.Info("MarkedAsRead");

                            isClickDUp = 1;

                        }

                        else
                        {
                            /// If the email was not sent (no unread email = no email)
                            /// Send the alert email via SMTP client from CRM Online Service User
                            if (email.IsRead)
                            {
                                string SUBJECT = "ClickDimensions May be Failing";

                                string BODY = HTML;
                                // string BODY = "Please check ClickDimensions Integration";

                                MailMessage message = new MailMessage
                                {
                                    IsBodyHtml = true,
                                    From = new MailAddress(FROM, FROMNAME)
                                };
                                message.To.Add(new MailAddress(TO));
                                message.Subject = SUBJECT;
                                message.Body = BODY;

                                message.Headers.Add("X-SES-CONFIGURATION-SET", CONFIGSET);

                                isClickDUp = 0;


                                using (var client = new SmtpClient(HOST, PORT))
                                {
                                    /// Pass SMTP credentials
                                    client.Credentials =
                                        new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);

                                    /// Enable SSL encryption
                                    client.EnableSsl = true;

                                    try
                                    {
                                        log.Info("Attempting to send email...");
                                        client.Send(message);
                                        log.Info("Email sent!");

                                    }
                                    catch (Exception ex)
                                    {
                                        log.Info("The email was not sent.");
                                        log.Info("Error message: " + ex.Message);

                                    }

                                }


                            }


                        }

                        /// This will mark the email as read as a backup
                        email.IsRead = true;
                        email.Update(ConflictResolutionMode.AutoResolve);
                        log.Info("BackupMarkedAsRead");

                    }

                }

                catch (Exception e)
                {
                    log.Info("An error has occured. \n:" + e.Message);
                }

            }


            if (DateTime.Compare(DateTime.Today, new DateTime(2018, 12, 10, 0, 0, 0)) >= 0)
            {


                string monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday).ToString("yyyy-MM-dd");
                string lastMonday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday).AddDays(-7).ToString("yyyy-MM-dd");


                string path = @"D:\home\site\wwwroot\Function2\ClickDUptime Weekly Report " + monday + ".csv";
                DateTime localDate2 = DateTime.Now;
                str = localDate2.ToString() + "," + isClickDUp;

                if (!File.Exists(path))
                {
                    log.Info("Creating & Updating Excel");

                    /// Create a file to write to.
                    string createText = "Timestamp,ClickDimensions Status" + Environment.NewLine + str;
                    File.WriteAllText(path, createText);


                    /// path is reset to old file here
                    path = @"D:\home\site\wwwroot\Function2\ClickDUptime Weekly Report " + lastMonday + ".csv";


                    /// Send email for last weeks excel sheet
                    if (File.Exists(path))
                    {
                        MailMessage mail = new MailMessage();
                        SmtpClient SmtpServer = new SmtpClient(HOST);
                        mail.From = new MailAddress(FROM);

                        // mail.To.Add("");
                        mail.To.Add("");

                        mail.Subject = "ClickDUptime Monitoring Report " + monday;
                        mail.Body = "";

                        log.Info("Sending Excel");

                        try
                        {
                            System.Net.Mail.Attachment attachment;
                            attachment = new System.Net.Mail.Attachment(path);
                            mail.Attachments.Add(attachment);

                            SmtpServer.Port = 587;
                            SmtpServer.Credentials = new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);
                            SmtpServer.EnableSsl = true;
                            SmtpServer.Send(mail);

                            log.Info("Excel Sent");

                        }
                        catch (Exception)
                        {
                            /// mail send failure
                            /// figure out what to do here
                            log.Info("Failing");

                        }

                    }

                }

                else
                {
                    /// This text is always added, making the file longer over time
                    /// if it is not deleted.
                    string appendText = Environment.NewLine + str;
                    File.AppendAllText(path, appendText);

                    /// Open the file to read from.
                    string readText = File.ReadAllText(path);

                }
            }

        }

    }

}
