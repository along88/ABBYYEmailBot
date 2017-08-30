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
                    EmailMessage(", was SUCCESSFUL!", ", your submission was a Success and is now ready for use! ");
                    break;
                case SenderStatus.FAIL:
                    EmailMessage(", FAILED! Open for more details", ", I was unable to process your submission, It is undetermined at this time what caused the issue.Please email Tech Services or try again! ");
                    break;
                case SenderStatus.PROCESSING:
                    EmailMessage(", is being Processed. Please Wait.", string.Format(" was recieved  and is currently being processed, I will notify you when I have completed processing this submission! Feel Free to email Tech Services if it has been more than 3 minutes since receiving this email and you have not recieved a status update from me!" + Environment.NewLine +
                        "Some fun facts about me(ABBY) while you wait:" + Environment.NewLine +
                        "\u2022 I am able to read most 140 documents to some degree of accuracy, currently I struggle with anything after the year 2010!" + Environment.NewLine +
                        "\u2022 I Will try to auto select the Construction Type, Burglar, and Fire Alarms for you if I find it on the 140 you submit!" + Environment.NewLine +
                        "\u2022 I won't be able to recognize any emails you send to me with an Excel file in it =("+Environment.NewLine+
                        "\u2022 I work best for you when you have a workstation with a lot of locations, anything less than 5 submissions you may be better processing manually!"));
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
                if (sender.name.ToLower().Equals("along@wkfc.com")) //this protects me from sending emails by accident to other active users besides me
                {
                    MailItem mailItem = (MailItem)application.Session.OpenSharedItem(@"C:\Users\along\Documents\ConsoleWEBAPI\ConsoleWEBAPI\ABBYYSuccessEmail.msg");
                    mailItem.Subject = "CN#:"+sender.ctrlNumber + emailSubject;
                    mailItem.Body = sender.ctrlNumber + emailBody;
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
