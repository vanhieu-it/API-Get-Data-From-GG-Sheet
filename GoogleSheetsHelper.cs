using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

namespace api_key
{
    class GoogleSheetsHelper
    {
        private static SheetsService GetSheetsService()
        {
            // Đường dẫn đến tệp JSON Service Account
            string credentialsPath = "C:/Users/Hieu/Desktop/api-key/key/key.json";

            // Đọc tệp JSON Service Account
            GoogleCredential credential;
            using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            // Tạo dịch vụ Google Sheets
            var sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "YourAppName",
            });

            return sheetsService;
        }

        public static bool CheckStringInGoogleSheets(string spreadsheetId, string sheetName, string columnName, string columnName2, string searchString)
        {
            try
            {
                // Tạo dịch vụ Google Sheets
                var sheetsService = GetSheetsService();

                // Lấy dữ liệu từ cột cụ thể
                string range = $"{sheetName}!{columnName}:{columnName2}";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                    sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;

                // Kiểm tra xem chuỗi có tồn tại trong cột và ngày hết hạn đã qua hay chưa
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row.Count > 0 && row[0].ToString() == searchString)
                        {
                            string expiryDateString = row[1].ToString();
                            DateTime.TryParseExact(expiryDateString, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime expiryDate);
                            // So sánh với ngày hiện tại
                            if (DateTime.Now > expiryDate)
                            {
                                Console.WriteLine("Chuỗi tồn tại nhưng đã hết hạn!");
                                return false;
                            }
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking string in Google Sheets: {ex.Message}");
                return false;
            }
        }
    }
}

  

