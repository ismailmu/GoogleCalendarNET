using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;


namespace GoogleCalendarNET
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "GoogleCalendarNET";
        static string conn_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=template.xlsx;Extended Properties=\"Excel 8.0;HDR=YES\"";
        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/googlecalendarnet.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            using (OleDbConnection conn = new OleDbConnection(conn_string))
            {
                conn.Open();
                using (OleDbCommand comm = new OleDbCommand("SELECT * FROM [Sheet3$]",conn))
                {
                    using (OleDbDataReader reader = comm.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        CalendarEvent myEvent = new CalendarEvent();
                        while (reader.Read())
                        {
                            string summary = reader[0].ToString() + "\r\n" + reader[1].ToString();
                            myEvent.insert(service, summary, Convert.ToDateTime(reader[2]), Convert.ToDateTime(reader[3]));
                        }
                    }
                }
            }

            CalendarList list = new CalendarList();
            list.Get(service);

            Console.WriteLine("------DONE------");
            Console.ReadLine();

        }//end static
    }//end class
}//end namespace9