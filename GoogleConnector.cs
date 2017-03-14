using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Data;

namespace SpreadSheetConnector
{
    public class GoogleConnector
    {
        private SheetsService service;
        private static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private static string CredsPath;
        public bool Authorized { get; set; }

        public GoogleConnector(string clientCredsPath)
        {
            CredsPath = clientCredsPath;
        }

        public Task Connect()
        {
            return new Task(() =>
            {
                try
                {
                    UserCredential credential;

                    using (var stream =
                        new FileStream(CredsPath, FileMode.Open, FileAccess.Read))
                    {
                        string credPath = "/";

                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                    }

                    service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential
                    });

                    Authorized = true;
                }
                catch (Google.GoogleApiException)
                {
                    Authorized = false;
                }
            });
        }

        public void Append(DataTable table, string spreadsheetID)
        {
            IList<IList<object>> matrix = new List<List<object>>();
            foreach(var row in table.Rows)
            {
                List<object> list = new List<object>();

                matrix.Add(new List<object>())
            }
            service.Spreadsheets.Values.Append(new ValueRange().
        }

        public void Overwrite(DataTable table, string spreadsheetID)
        {

        }
    }
}
