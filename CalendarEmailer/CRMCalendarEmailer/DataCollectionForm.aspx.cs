///////////////////////////////   DataCollectionForm   /////////////////////////////////////////
//Summary of Work: Email form that generates .ics files for Outlook calendars
//Created by: Sage Aucoin       Created Date: 10/29/2014
////////////////////////////////////////////////////////////////////////////////////////////////



using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace CRMCalendarEmailer
{
    public partial class DataCollectionForm : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            SubmitForm(sender, e);
        }


        ///////////////////////////////   SubmitForm()   ///////////////////////////////////////////////
        //Summary of Work: Grabs all form information and sends email.
        //Created by: Sage Aucoin       Created Date: 10/29/2014
        //PreCondition: Form is filled out with appropriate information. Web.Config is configured
        //              with appropriate 'sender' data including smtp email, password, port, etc.
        //PostCondition: Email is sent to user indicated on form with appropriate data from the 
        //               address specified in the web.config.
        //
        //Misc Note: Changing the string filename varible will give you different attachments. 
        //           Removing it altogether will remove all attachments.
        ////////////////////////////////////////////////////////////////////////////////////////////////
        protected void SubmitForm(object sender, EventArgs e)
        {
            if (Request.QueryString["sendTo"] != "" && Request.QueryString["sendTo"] != null)
            {
                //TO-DO: Add better conditioning to indicate 'Required Fields'
                var sendTo = Request.QueryString["sendTo"];
                var subject = Request.QueryString["subject"];
                var location = Request.QueryString["location"];
                var startDateTime = Convert.ToDateTime(Request.QueryString["startDateTime"]);
                var endDateTime = startDateTime.AddMinutes(30);
                var message = Request.QueryString["message"];
                var smtpEmail = ConfigurationManager.AppSettings["smtpEmail"];
                var smtpHost = ConfigurationManager.AppSettings["smtpHost"];
                var smtpUserName = ConfigurationManager.AppSettings["smtpUserName"];
                var smtpPassword = ConfigurationManager.AppSettings["smtpPassword"];
                //Currently isn't used anywhere, but will leave this here in case we want to use it later.
                //var smtpDisplayName = ConfigurationManager.AppSettings["smtpDisplayName"];

                //Composes the body of the email.
                message = string.Format(
                    "{0}\n" +
                    "\n\nOrganizer: {1}" +
                    "\n\nLocation: {2}" +
                    "\n\nStart Date: {3}" +
                    "\n\nEnd Date: {4}", message, sendTo, location, startDateTime.ToString("f"),
                    endDateTime.ToString("f"));
                try
                {
                    var appointment = new Appointment();

                    //Create the iCal file. This can be changed for other attachments as needed.
                    string fileName = appointment.MakeHourEvent(subject, location, startDateTime.Date,
                        startDateTime.ToString("HH:mm"), endDateTime.ToString("HH:mm"));

                    //Attach the iCal file.
                    var attachment = new Attachment(fileName);

                    //Send the email
                    SendMail(smtpEmail, sendTo, subject, message, attachment, smtpHost, smtpUserName, smtpPassword);

                    //Delete the iCal file after send regardless of pass/fail to eliminate footsteps.
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                    }
                }
                catch (Exception)
                {
                    Response.Write("<script>alert('There was an error. Please contact IDS support.');</script>");
                    throw;
                }


                Response.Write("A calendar reminder has been sent to your email.");
            }
            else if (Request.QueryString["sendTo"] == null || Request.QueryString["sendTo"] == "")
            {
                Response.Write("Meeting has been submitted but no email was sent. Please confirm that you have an email configured.");
            }
        }


        ///////////////////////////////   SendMail()   /////////////////////////////////////////////////
        //Summary of Work: Sends email using data passed
        //Created by: Sage Aucoin       Created Date: 10/29/2014
        //PreCondition:  Requires send/receive email addresses. Send email requires valid SMTP
        //               host server, email address, and password.
        //PostCondition: Email is sent and cleared from memory. 
        ////////////////////////////////////////////////////////////////////////////////////////////////
        public void SendMail(string from, string to, string subject, string body, Attachment attachment, string smtpHost, string smtpUserName, string smtpPassword)
        {
            
            var basicCredential = new NetworkCredential(smtpUserName, smtpPassword);
            //Create Email Object
            var mail = new MailMessage(from, to, subject, body);
            //Attach file.
            //TO-DO: Add behavior for no attachment.
            mail.Attachments.Add(attachment);
            //Create new smtp object with credentials provided in web.config.
            var smtp = new SmtpClient(smtpHost)
            {
                Host = smtpHost,
                UseDefaultCredentials = false,
                Credentials = basicCredential, //Do not use default credentials since only username/password is needed.
                //Setup on the host, increase timeout to 5 minutes
                //60 seconds * 5 minutes * 1000 miliseconds == 5 minutes
                Timeout = (60*5*1000)   
            };
            
            smtp.Send(mail);
            mail.Dispose();
        }
    }
}