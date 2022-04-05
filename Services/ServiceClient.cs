using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebservicesSage.Cotnroller;
using WebservicesSage.Utils;

namespace WebservicesSage.Services
{
    class ServiceClient : ServiceAbstract
    {
        public ServiceClient()
        {
            setAlive(true);
        }
        public void ToDoOnLaunch()
        {
            try
            {
                if (isAlive())
                {
                    ControllerClient.LaunchService();
                }
            }
            catch (Exception e)
            {
                UtilsMail.SendErrorMail(DateTime.Now + e.Message + Environment.NewLine + e.StackTrace + Environment.NewLine, "SERVICES CLIENT : ToDoOnLaunch");
            }
        }
        public override void ToDoOnFirstCommit()
        {
            try
            {
                if (isAlive())
                {
                    
                    Task taskA = new Task(() => ControllerClient.SendAllClients());
                    taskA.Start();
                }
            }
            catch(Exception e)
            {
                UtilsMail.SendErrorMail(DateTime.Now + e.Message + Environment.NewLine + e.StackTrace + Environment.NewLine, "SERVICES CLIENT : ToDoOnFirstCommit");
            }
        }
        public void sendProspect()
        {
            if (isAlive())
            {
                Task taskA = new Task(() => ControllerClient.SendProspect());
                taskA.Start(); 
            }
        }
        public void SendClient(string ct_num)
        {
            try
            {
                if (isAlive())
                {
                    Task taskA = new Task(() => ControllerClient.SendClient(ct_num));
                    taskA.Start();
                }
            }
            catch (Exception e)
            {
                UtilsMail.SendErrorMail(DateTime.Now + e.Message + Environment.NewLine + e.StackTrace + Environment.NewLine, "SERVICES CLIENT : ToDoOnFirstCommit");
            }
        }
        public void SendSaleDocument(string ct_num)
        {
            try
            {
                if (isAlive())
                {
                    Task taskA = new Task(() => ZohoController.SendSalesDocument(ct_num));
                    taskA.Start();
                }
            }
            catch (Exception e)
            {
                UtilsMail.SendErrorMail(DateTime.Now + e.Message + Environment.NewLine + e.StackTrace + Environment.NewLine, "SERVICES CLIENT : ToDoOnFirstCommit");
            }
        }
        public void SendSaleDocument()
        {
            try
            {
                if (isAlive())
                {
                    Task taskA = new Task(() => ZohoController.SendSalesDocument());
                    taskA.Start();
                }
            }
            catch (Exception e)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + e.Message + Environment.NewLine);
                sb.Append(DateTime.Now + e.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\ErrorDoc.txt", sb.ToString());
                sb.Clear();
                UtilsMail.SendErrorMail(DateTime.Now + e.Message + Environment.NewLine + e.StackTrace + Environment.NewLine, "SERVICES CLIENT : ToDoOnFirstCommit");
            }
        }
    }
}
