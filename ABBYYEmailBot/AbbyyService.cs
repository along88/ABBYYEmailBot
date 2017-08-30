using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ABBYYEmailBot
{ 
     public class AbbyyService
    {
        private Flexicapture11.FlexiCaptureWebServiceApiVersion3 service = new Flexicapture11.FlexiCaptureWebServiceApiVersion3();
        private EmailSender emailSender = new EmailSender();
        private int roleType;
        private int stationType;
        private int sessionID;
        private Flexicapture11.Project[] projects;
        private Flexicapture11.Project project;
        private int projectID;

        public AbbyyService()
        {
            
            
            service.Url = "http://10.10.11.110/FlexiCapture11/Server/WebServicesExternal.dll?Handler=Version3";
            service.Credentials = new NetworkCredential("along", "20Wkf@16!");
            roleType = 6;
            stationType = 2;
            sessionID = service.OpenSession(roleType, stationType);
            projectID = OpenTestProject();
            while (true)
            {
                MonitorImports();
                MonitorExports();
            }
        }
        private int OpenTestProject()
        {
            try
            {
                projects = service.GetProjects();
                if (projects != null)
                {
                    foreach (Flexicapture11.Project _project in projects)
                    {
                        if(_project.Name.Equals("ABBYY_Test"))
                        {
                            project = _project;
                            break;
                        }
                    }
                }
                return service.OpenProject(sessionID, project.Guid); ;
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void MonitorImports()
        {
            bool isMonitoring = true;
            while (isMonitoring)
            {
                Flexicapture11.Batch[] Batches = service.GetBatches(projectID, sessionID, false);

                foreach (Flexicapture11.Batch batch in Batches)
                {
                    if (batch.StageExternalId == 200)
                    {
                        //Batch is being processed
                        //System.Windows.MessageBox.Show(batch.Name + "ABBYY is Processing now!");
                        GetEmailSender(batch, SenderStatus.PROCESSING); // get the email of the thing being processed in case we want to inform that its being procssed
                        Model.EmailBot importEmail = new Model.EmailBot(emailSender);//send email
                                                                                     //keeps looping need to fix this
                        isMonitoring = false;
                        break;
                    }
                }
            }
        }
        private void MonitorExports()
        {
            bool isMonitoring = true;
            while (isMonitoring)
            {
                Flexicapture11.Batch[] Batches = service.GetBatches(projectID, sessionID, false);
                foreach (Flexicapture11.Batch batch in Batches)
                {
                    if (batch.StageExternalId == 800)
                    {
                        GetEmailSender(batch, SenderStatus.SUCCESS);
                        ABBYYEmailBot.Model.EmailBot email = new Model.EmailBot(emailSender);
                        isMonitoring = false;
                        break;
                    }
                }
            }
        }
        private string GetControlNumber(string subjectLine)
        {
            // We told the office to only use square brackets but that's got a snowball's chance in hell of happening
            char[] leftEnclosures = { '(', '{', '[' };
            char[] rightEnclosures = { ')', '}', ']' };
            string controlNumber = "";

            int openEnclosure = subjectLine.IndexOfAny(leftEnclosures);
            int closeEnclosure = subjectLine.IndexOfAny(rightEnclosures);

            try
            {
                controlNumber = subjectLine.Substring(openEnclosure + 1, (closeEnclosure - openEnclosure) - 1);
                return controlNumber;
            }
            catch (ArgumentOutOfRangeException rangeExc)
            {
                rangeExc.ToString();
            }
            return null;
        }
        private void GetEmailSender(Flexicapture11.Batch batch, SenderStatus status)
        {
            emailSender.ctrlNumber = GetControlNumber(batch.Properties[0].Value.ToString());
            var batchEmailSender = batch.Properties[1].Value.ToString();
            var firstInitial = batchEmailSender[0].ToString();
            var emptySpace = batchEmailSender.IndexOf(" ");
            var lastName = batchEmailSender.Substring(emptySpace + 1);
            emailSender.name = string.Format("{0}{1}@WKFC.com", firstInitial, lastName);
            emailSender.status = status;   
        }
        
    }

}

