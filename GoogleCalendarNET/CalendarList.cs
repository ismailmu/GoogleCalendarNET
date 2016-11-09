using System;

using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace GoogleCalendarNET
{
    public class CalendarList
    {
        public void Get(CalendarService service)
        {
            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();
            Console.WriteLine("Upcoming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine("---------------------------------------------------");
                    Console.WriteLine("{0} ({1})", eventItem.Summary, when);
                    Console.WriteLine();
                }
                
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
            Console.WriteLine();
        }
    }//end class
}//end namespace
