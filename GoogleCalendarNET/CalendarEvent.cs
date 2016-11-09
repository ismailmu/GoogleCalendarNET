using System;

using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace GoogleCalendarNET
{
    public class CalendarEvent
    {
        static string TIME_ZONE = "Asia/Jakarta";
        public void insert(CalendarService service,string summary,DateTime startDateTime,DateTime endDateTime)
        {
            Event myEvent = new Event{
                Summary = summary,
                Start =  new EventDateTime() {
                    DateTime = startDateTime,
                    TimeZone = TIME_ZONE
                },
                End = new EventDateTime() {
                    DateTime = endDateTime,
                    TimeZone = TIME_ZONE
                }
                ,Recurrence = new String[] {
                    "RRULE:FREQ=DAILY;COUNT=30"
                }
            };

            Event recurringEvent = service.Events.Insert(myEvent, "primary").Execute();
        }
    }//end class
}//end namespace
