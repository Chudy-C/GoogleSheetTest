using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheets
{
    class Program
    {

        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "test";
        static readonly string SpreadsheetId = "1m0QWQlsi4N_P_ge11d2D8DM6el8zOr3TZQm-w0ZB3-o";

        static readonly string sheet = "ASPnet";

        static SheetsService service;


        static void Main(string[] args)
        {
            GoogleCredential credential;

            using (var stream = new FileStream("zglaszanie-awarii-e37cb8c264ca.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //ReadEntries();
            CreateEntry();
        }
        
        static void ReadEntries()
        {

            var range = $"{sheet}!A1:D10";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;

            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    Console.WriteLine("{0} | {1} | {2} | {3}", row[0], row[1], row[2], row[3]);
                }
            }
            else Console.WriteLine("No data was found");
        }

        static void CreateEntry()
        {
            var range = $"{sheet}!A:D";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { DateTime.Now, "DTE", "TEST VisualStudio1", DateTime.Now };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();


        }

    }
}
