Sample usage for outlook/ical

var stream = CalendarHelper.CreateIcs(appointment);
var result = new FileStreamResult(stream, "text/calendar") { FileDownloadName = string.Format("{0}.ics", appointment.Title) };