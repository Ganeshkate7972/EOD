

using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SimpleHeartBeatService
{
    class SendMailService
    {
        // This function write log to LogFile.text when some error occurs.  


        System.Data.DataTable dt;
        bool ckday;
        bool CheckExcle;
        string todaydate;
        public static void writeLog(string strValue)
        {
            try
            {
                //Logfile
                string path = ConfigurationSettings.AppSettings["LogFileDrive"];
                string logFilePath = path + "\\" + DateTime.Now.ToString("dd-MM-yyyy") + "\\";
                if (Directory.Exists(logFilePath) == false)
                {
                    Directory.CreateDirectory(logFilePath);
                }
                StreamWriter sw;
                string file = String.Format("Services {0}.txt", DateTime.Now.ToString("dd-MM-yyyy"));
                string filePath = Path.Combine(logFilePath, file);
                if (!File.Exists(filePath))
                { sw = File.CreateText(filePath); }
                else
                { sw = File.AppendText(filePath); }
                LogWrite(strValue, sw);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
            }
        }
        private static void LogWrite(string logMessage, StreamWriter w)
        {
            w.WriteLine("{0}", logMessage);
            w.WriteLine();
            w.WriteLine();
        }


        public void sendEMailThroughOUTLOOK()
        {
            string ExcleToCreateHtmlBodyPath = ConfigurationSettings.AppSettings["ExcleToCreateHtmlBodyPath"];
            SendMailService.writeLog("ExcleToCreateHtmlBodyPath" + ExcleToCreateHtmlBodyPath);
            string ExcleToCreateHtmlBodyPathExt = ConfigurationSettings.AppSettings["ExcleToCreateHtmlBodyPathExt"];
            SendMailService.writeLog("ExcleToCreateHtmlBodyPathExt" + ExcleToCreateHtmlBodyPathExt);
            string filePath = ExcleToCreateHtmlBodyPath;
            SendMailService.writeLog("ExcleToCreateHtmlBodyPath" + filePath);
            string fileExt = ExcleToCreateHtmlBodyPathExt;
            SendMailService.writeLog("ExcleToCreateHtmlBodyPathExt" + fileExt);
            string EnableToCreateHtmlBody = ConfigurationSettings.AppSettings["EnableToCreateHtmlBody"];
            if (EnableToCreateHtmlBody=="1")
            {
                dt = ReadExcel(filePath, fileExt);
                SendMailService.writeLog("ReadExcel" + dt);
                CheckExcle = CheckUpdateExcle(dt);
                if (CheckExcle)
                {
                    SendMailService.writeLog("CheckExcle" + CheckExcle);
                    return;
                }
            }

            ckday = CheckSunday();
            if (ckday)
            {
                SendMailService.writeLog("ckday" + ckday);
                return;
            }

            string body = ConvertDataTableToHTML(dt);
            SendMailService.writeLog("body" + body);
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = body;
                //Add an attachment.
                string AttachmentEnable = ConfigurationSettings.AppSettings["AttachmentEnable"];
                SendMailService.writeLog("AttachmentEnable" + AttachmentEnable);
                string AttachmentPath = ConfigurationSettings.AppSettings["AttachmentPath"];
                if (AttachmentEnable == "1")
                {
                    String sDisplayName = "Attachment";
                    int iPosition = (int)oMsg.Body.Length + 1;
                    int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                    //now attached the file
                    Outlook.Attachment oAttach = oMsg.Attachments.Add
                                                (AttachmentPath, iAttachType, iPosition, sDisplayName);
                }
                SendMailService.writeLog("AttachmentPath" + AttachmentPath);

                //Subject line
                string MailSubject = ConfigurationSettings.AppSettings["MailSubject"];
                SendMailService.writeLog("MailSubject" + MailSubject);
                string MailSubjectTodaydateEnable = ConfigurationSettings.AppSettings["MailSubjectTodaydateEnable"];
                if (MailSubjectTodaydateEnable == "1") { todaydate = System.DateTimeOffset.Now.ToString("dd/MM/yyyy"); SendMailService.writeLog("MailSubject" + MailSubject); };
                oMsg.Subject = MailSubject + " " + todaydate;

                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                string Recips = ConfigurationSettings.AppSettings["Recips"];
                SendMailService.writeLog("Recips" + Recips);
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(Recips);
                string EnableCC = ConfigurationSettings.AppSettings["EnableCC"];
                string RecipsCC = ConfigurationSettings.AppSettings["RecipsCC"];
                SendMailService.writeLog("EnableCC" + EnableCC);
                if (EnableCC == "1")
                {
                    SendMailService.writeLog("RecipsCC" + RecipsCC);
                    Outlook.Recipient oCC = (Outlook.Recipient)oRecips.Add(RecipsCC);
                    oCC.Type = (int)Outlook.OlMailRecipientType.olCC;
                }

                oRecip.Resolve();
                // Send.
                oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                SendMailService.writeLog("Send Mail" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));
            }//end of try block
            catch (Exception ex)
            {
                SendMailService.writeLog("Send Mail Error"+ ex.Message + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));
            }//end of catch
        }//end of Email Method



        public static string ConvertDataTableToHTML(System.Data.DataTable dt)
        {
            string html = "";
            try
            {
                html = "<table border=\"" + 1 + "\" >";
                //add header row
                System.Data.DataRow firstRow = dt.Rows[0];
                html += "<tr>";
                string fc1 = Convert.ToString(firstRow[0]);
                string fc2 = Convert.ToString(firstRow[1]);
                string fc3 = Convert.ToString(firstRow[2]);
                string fc4 = Convert.ToString(firstRow[3]);
                string fc5 = Convert.ToString(firstRow[4]);
                string fc6 = Convert.ToString(firstRow[5]);
                string fc7 = Convert.ToString(firstRow[6]);
                if (fc1 == "")
                {
                    fc1 = "Date";
                }
                if (fc2 == "")
                {
                    fc2 = "Client Name";
                }
                if (fc3 == "")
                {
                    fc3 = "Project";
                }
                if (fc4 == "")
                {
                    fc4 = "Task Detail";
                }
                if (fc5 == "")
                {
                    fc5 = "Assign on";
                }
                if (fc6 == "")
                {
                    fc6 = "Completed on";
                }
                if (fc7 == "")
                {
                    fc7 = "Status";
                }
                html += "<td>" + fc1 + "</td>";
                html += "<td>" + fc2 + "</td>";
                html += "<td>" + fc3 + "</td>";
                html += "<td>" + fc4 + "</td>";
                html += "<td>" + fc5 + "</td>";
                html += "<td>" + fc6 + "</td>";
                html += "<td>" + fc7 + "</td>";


                html += "</tr>";


                DataRow lastRow = dt.Rows[dt.Rows.Count - 1];
                html += "<tr>";
                string lc1 = Convert.ToString(lastRow[0]);
                string lc2 = Convert.ToString(lastRow[1]);
                string lc3 = Convert.ToString(lastRow[2]);
                string lc4 = Convert.ToString(lastRow[3]);
                string lc5 = Convert.ToString(lastRow[4]);
                string lc6 = Convert.ToString(lastRow[5]);
                string lc7 = Convert.ToString(lastRow[6]);
                html += "<td>" + lc1 + "</td>";
                html += "<td>" + lc2 + "</td>";
                html += "<td>" + lc3 + "</td>";
                html += "<td>" + lc4 + "</td>";
                html += "<td>" + lc5 + "</td>";
                html += "<td>" + lc6 + "</td>";
                html += "<td>" + lc7 + "</td>";
                html += "</tr>";

                html += "</table>";

            }
            catch (Exception ex)
            {
                SendMailService.writeLog("ConvertDataTableToHTML ERROR" + ex.Message + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));
            }
            return html;
        }

        public System.Data.DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(conn))
            {
                try
                {
                    System.Data.OleDb.OleDbDataAdapter oleAdpt = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex) 
                {
                    SendMailService.writeLog("ReadExcel ERROR" + ex.Message + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));
                }
            }
            return dtexcel;
        }

        public bool CheckSunday()
        {
            var result = false;
            try
            {

                DayOfWeek day = DateTime.Now.DayOfWeek;
                if (day == DayOfWeek.Sunday)
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                SendMailService.writeLog("CheckSunday ERROR" + ex.Message + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));
            }

            return result;
        }

        public bool CheckUpdateExcle(System.Data.DataTable dt)
        {
            var result = true;
            try
            {

                System.Data.DataRow lastRow = dt.Rows[dt.Rows.Count - 1];
                string lc1 = Convert.ToString(lastRow[0]);
                var d1 = Convert.ToDateTime(lc1).ToString("dd/MM/yyyy");
                if (d1 == DateTime.Now.ToString("dd/MM/yyyy"))
                {
                    result = false;
                }

            }
            catch (Exception ex)
            {

                SendMailService.writeLog("CheckUpdateExcle ERROR" + ex.Message + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ffff"));

            }
            return result;
        }




    }
}
