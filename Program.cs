using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using api_key;

class Program
{
    static void Main(string[] args)
    {

        string spreadsheetId = "your-spreadsheetId";
        string sheetName = "Key";
        string columnName = "A"; // Đặt tên cột cần kiểm tra
        string columnName2 = "B"; // Đặt tên cột cần kiểm tra
        //string searchString = "YourSearchString"; // Chuỗi cần kiểm tra
        Console.Write("Nhập chuỗi cần kiểm tra: ");
        string searchString = Console.ReadLine();

        bool result = GoogleSheetsHelper.CheckStringInGoogleSheets(spreadsheetId, sheetName, columnName, columnName2, searchString);

        if (result)
        {
            Console.WriteLine("Chuỗi tồn tại trong Google Sheets!");
        }
        else
        {
            Console.WriteLine("Chuỗi không tồn tại trong Google Sheets.");
        }
    }
}
