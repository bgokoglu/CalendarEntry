using System;
using System.IO;
using System.Text;

namespace CalendarEntry
{
    public static class CalendarHelper
    {
        /// <summary>
        /// Creates Outlook entry
        /// </summary>
        /// <param name="entry">Appointment</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream CreateOutlookEntry(Appointment entry)
        {
            return CreateIcs(entry);
        }
        
        /// <summary>
        /// Creates iCal entry
        /// </summary>
        /// <param name="entry">Appointment</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream CreateiCalEntry(Appointment entry)
        {
            return CreateIcs(entry);
        }

        /// <summary>
        /// Creates Google Calendar link
        /// </summary>
        /// <param name="entry">Appointment</param>
        /// <returns>string</returns>
        public static string CreateGoogleCalendarLink(Appointment entry)
        {
            const string url = "https://www.google.com/calendar/render?action=TEMPLATE&hl=en&text={0}&dates={1}/{2}&location={3}&details={4}sf=true&output=xml";
            return string.Format(url, entry.Title, GetFormattedDate(entry.StartDateTime), GetFormattedDate(entry.EndDateTime), entry.Location, entry.Description);
        }

        /// <summary>
        /// Creates Yahoo Calendar link
        /// </summary>
        /// <param name="entry">Appointment</param>
        /// <returns>string</returns>
        public static string CreateYahooCalendarLink(Appointment entry)
        {
            const string url = "http://calendar.yahoo.com/?v=60&view=d&type=20&title={0}&st={1}&et={2}&desc={4}&in_loc={3}";
            return string.Format(url, entry.Title, GetFormattedDate(entry.StartDateTime), GetFormattedDate(entry.EndDateTime), entry.Location, entry.Description);
        }

        private static string GetFormattedDate(DateTime dt)
        {
            return dt.ToString("yyyyMMddTHHmmss");
        }

        private static MemoryStream CreateIcs(Appointment entry)
        {
            var sb = new StringBuilder();

            sb.AppendLine(string.Format("BEGIN:VCALENDAR{0}", Environment.NewLine));
            sb.AppendLine(string.Format("VERSION:2.0{0}", Environment.NewLine));
            sb.AppendLine(string.Format("PRODID:-//hacksw/handcal//NONSGML v1.0//EN{0}", Environment.NewLine));
            sb.AppendLine(string.Format("BEGIN:VEVENT{0}", Environment.NewLine));
            sb.AppendLine(string.Format("DTSTART:{0}{1}", GetFormattedDate(entry.StartDateTime), Environment.NewLine));
            sb.AppendLine(string.Format("DTEND:{0}{1}", GetFormattedDate(entry.EndDateTime), Environment.NewLine));
            sb.AppendLine(string.Format("{0}{1}", string.Format("SUMMARY:{0}", entry.Title), Environment.NewLine));
            sb.AppendLine(string.Format("{0}{1}", string.Format("LOCATION:{0}", entry.Location), Environment.NewLine));
            sb.AppendLine(string.Format("{0}{1}", string.Format("DESCRIPTION:{0}", entry.Description), Environment.NewLine));
            sb.AppendLine("END:VEVENT" + Environment.NewLine);
            sb.AppendLine("END:VCALENDAR");

            var ms = new MemoryStream();
            var enc = new UTF8Encoding();
            var arrBytData = enc.GetBytes(sb.ToString());
            ms.Write(arrBytData, 0, arrBytData.Length);
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }
    }
}
