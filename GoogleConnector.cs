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
        public Task Connect(string clientCredsPath)
        {
            return Task.Factory.StartNew(() =>
            {
                try
                {
                    CredsPath = clientCredsPath;
                    UserCredential credential;

                    using (var stream =
                        new FileStream(CredsPath, FileMode.Open, FileAccess.Read))
                    {
                        string credPath = new FileInfo(CredsPath).DirectoryName;

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

        public async Task<int?> Append(DataTable table, string spreadsheetID, int deleteHeaders)
        {
            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>>();
            for (int i = deleteHeaders; i < table.Rows.Count; i++)
            {
                valueRange.Values.Add(new List<object>(table.Rows[i].ItemArray));
            }
            SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                new SpreadsheetsResource.ValuesResource.GetRequest(service, spreadsheetID, "A:A");
            ValueRange response = getRequest.Execute();
            int responseRowsLength = 1;
            if (response.Values != null)
            {
               responseRowsLength = response.Values.Count;
            }
            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest =
                new SpreadsheetsResource.ValuesResource.AppendRequest(
                    service,
                    valueRange,
                    spreadsheetID,
                    string.Format("A{0}:A", responseRowsLength));
            appendRequest.ValueInputOption =
                SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            return (await appendRequest.ExecuteAsync()).Updates.UpdatedRows;
        }

        public void ClearSpreadsheet(string spreadsheetID)
        {
            service.Spreadsheets.Values.Clear(new ClearValuesRequest(),
                spreadsheetID,
                "Sheet1").Execute();
        }

        public async Task<int?> Overwrite(DataTable table, string spreadsheetID, int deleteHeaders)
        {
            ClearSpreadsheet(spreadsheetID);
            return await Append(table, spreadsheetID, deleteHeaders);
        }
    }
}
