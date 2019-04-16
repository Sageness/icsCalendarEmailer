using System;
using System.Globalization;
using System.Web;
using System.IO;

namespace CRMCalendarEmailer
{
    ///////////////////////////////////   Appointment.cs   ///////////////////////////////////////////////
    //Summary of Class: Functions in this class are all written with the goal of creating an iCal file
    //                  that can be attached to an email. Currently only supports hourly events, but 
    //                  will eventually be adapted to support daily events.
    //*****EXAMPLE iCAL FILE*****
    //BEGIN:VCALENDAR
    //VERSION:2.0
    //PRODID:-//hacksw/handcal//NONSGML v1.0//EN
    //BEGIN:VEVENT
    //UID:uid1@example.com
    //DTSTAMP:19970714T170000Z
    //ORGANIZER;CN=John Doe:MAILTO:john.doe@example.com
    //DTSTART:19970714T170000Z
    //DTEND:19970715T035959Z
    //SUMMARY:Bastille Day Party
    //END:VEVENT
    //END:VCALENDAR
    //Created by: Sage Aucoin       Created Date: 10/29/2014
    ///////////////////////////////////////////////////////////////////////////////////////////////////////
    public class Appointment
    {
        private StreamWriter _writer;

        ///////////////////////////////////   GetFormatedDate   ///////////////////////////////////////////////
        //Summary of Work: Grabs the user-inputted date and splits it into the appropriate format needed for
        //                 iCal.
        //Created by: Sage Aucoin       Created Date: 10/29/2014
        //Precondition: Date is submitted
        //Postcondition: Date is formatted as string in YYMMdd format.
        ///////////////////////////////////////////////////////////////////////////////////////////////////////
        public string GetFormatedDate(DateTime date)
        {
            var yy = date.Year.ToString(CultureInfo.InvariantCulture);
            var mm = date.Month.ToString(CultureInfo.InvariantCulture);
            var dd = date.Day.ToString(CultureInfo.InvariantCulture);
            //Checks month being before October and adds a 0 to the front if it is to maintain iCal format. 
            if (date.Month < 10) mm = "0" + date.Month.ToString(CultureInfo.InvariantCulture);
            else mm = date.Month.ToString(CultureInfo.InvariantCulture);
            //Checks day being before 10 and adds a 0 to the front if it is to maintain iCal format. 
            if (date.Day < 10) dd = "0" + date.Day.ToString(CultureInfo.InvariantCulture);
            else dd = date.Day.ToString(CultureInfo.InvariantCulture);
            return yy + mm + dd;
        }

        ///////////////////////////////////   GetFormattedTime   //////////////////////////////////////////////
        //Summary of Work: Grabs the user-inputted Time and splits it into the appropriate format needed for
        //                 iCal.
        //Created by: Sage Aucoin       Created Date: 10/29/2014
        //Precondition: Time is submitted
        //Postcondition: Time is formatted as string in hhmmss format.
        ///////////////////////////////////////////////////////////////////////////////////////////////////////
        public string GetFormattedTime(string time)
        {
            var times = time.Split(':');
            var hh = times[0];
            var mm = times[1];
            return hh + mm + "00";

        }

        ///////////////////////////////////   MakeHourEvent   /////////////////////////////////////////////////
        //Summary of Work: Creates an iCal event based on hourly times. End time is set to 0:30 ahead of start
        //Created by: Sage Aucoin       Created Date: 10/29/2014
        //Precondition: Date & Time are both properly formatted and required data is passed.
        //Postcondition: iCal file is created for the date and time given.
        ///////////////////////////////////////////////////////////////////////////////////////////////////////
        public string MakeHourEvent(string subject, string location, DateTime date, string startTime, string endTime)
        {
            var path = HttpContext.Current.Server.MapPath(@"iCal");
            string filePath = path + subject + ".ics";
            _writer = new StreamWriter(filePath);
            _writer.WriteLine("BEGIN:VCALENDAR");
            _writer.WriteLine("VERSION:2.0");
            _writer.WriteLine("PRODID:-//hacksw/handcal//NONSGML v1.0//EN");
            _writer.WriteLine("BEGIN:VEVENT");
            string startDateTime = GetFormatedDate(date) + "T" + GetFormattedTime(startTime);
            string endDateTime = GetFormatedDate(date) + "T" + GetFormattedTime(endTime);
            _writer.WriteLine("DTSTART:" + startDateTime);
            _writer.WriteLine("DTEND:" + endDateTime);
            _writer.WriteLine("SUMMARY:" + subject);
            _writer.WriteLine("LOCATION:" + location);
            _writer.WriteLine("END:VEVENT");
            _writer.WriteLine("END:VCALENDAR");
            _writer.Close();
            return filePath;
        }
    }
}