using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ABBYYEmailBot
{
    public class EmailBot
    {

        private Application application = null;
        private string emailSubject = "";
        private string emailBody = "";
        private string debugText;
        
        public EmailBot(EmailSender emailSender)
        {
            GetApplicaitonObject();

            switch (emailSender.status)
            {
                case SenderStatus.NONE:
                    break;
                case SenderStatus.SUCCESS:
                    EmailMessage(", SUCCESSFUL!", ", was a Success and is now ready for use! ");
                    break;
                case SenderStatus.FAIL:
                    EmailMessage(", FAILED! Open for more details", ", ABBYY was unable to process your submission, It is undetermined at this time what caused the issue.Please email Tech Services or try again! ");
                    break;
                case SenderStatus.PROCESSING:
                    EmailMessage(", IN-PROGRESS.", string.Format(" was recieved and is currently being processed, a notification email will follow once ABBYY has completed processing this submission! Feel Free to email Tech Services if it has been more than 3 minutes since receiving this email and you have not recieved a status update!" + Environment.NewLine +
                        "Some fun facts about ABBYY while you wait:" + Environment.NewLine +
                        "\u2022 ABBYY is able to read most 140 documents to some degree of accuracy, currently ABBYY struggles with anything after the year 2010!" + Environment.NewLine +
                        "\u2022 ABBYY Will try to auto select the Construction Type, Burglar Alarms, and Fire Alarms for you if it finds them on the 140 you submit!" + Environment.NewLine +
                        "\u2022 ABBYY won't be able to recognize any emails you send with an Excel file attachment =("+Environment.NewLine+
                        "\u2022 ABBYY works best for you when you have a workstation with a lot of locations, anything less than 5 submissions you may be better off processing manually!" + Environment.NewLine +
                    "Please Continue to Send us your feedback and errors so we can make ABBYY great together!"));
                    break;
                default:
                    break;
            }
            SendEmail(emailSender);

        }
        private Application GetApplicaitonObject()
        {
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {
                application = Marshal.GetActiveObject("Outlook.Application") as Application;
            }
            else
            {
                application = new Application();
                NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Missing.Value, Missing.Value);
                nameSpace = null;
            }
            return application;
        }
        private void EmailMessage(string subject, string body)
        {
            emailSubject = subject;
            emailBody = body;
        }
        private void SendEmail(EmailSender sender)
        {
            try
            {
                if (sender.name != null) //this protects me from sending emails by accident to other active users besides me
                {
                    MailItem mailItem = (MailItem)application.Session.OpenSharedItem(@"C:\Users\along\Documents\ConsoleWEBAPI\ConsoleWEBAPI\ABBYYSuccessEmail.msg");
                    mailItem.Subject = "CN#:"+sender.ctrlNumber + emailSubject;
                    mailItem.Body = "CN#:" + sender.ctrlNumber + emailBody;
                    mailItem.To = sender.name;
                    mailItem.Send();
                    //debugText = "SUCCESS!";
                }

                else
                    debugText = "Email was:" + sender;
            }
            catch (System.Exception exception)
            {
                debugText = exception.Message;
                //need to insert error log file here
            }

        }
    }
}
