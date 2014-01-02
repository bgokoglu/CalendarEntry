using System;

namespace CalendarEntry
{
    public class Appointment
    {
        public string Title { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public string Location { get; set; }
        public string Description { get; set; }
    }
}
