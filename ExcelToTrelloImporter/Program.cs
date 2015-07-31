using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using TrelloNet;

namespace ExcelToTrelloImporter
{
    /// <summary>
    ///     find the board ids from the below url
    ///     https://trello.com/1/members/me/boards?fields=name;
    /// </summary>
    internal class Program
    {
        private const string File = @"C:\dev\ExcelToTrelloImporter\ExcelToTrelloImporter\userstories.xlsx";

        private static void Main(string[] args)
        {
            var list = ExtractDevCards();

            ITrello trello = new Trello("7b17eb1ed849a91c051da9c924f93cfb");
            // var url = trello.GetAuthorizationUrl("userstorydataloader", Scope.ReadWrite);
            trello.Authorize("bdf4ef9325312874025ccce4a197c7794720e36079d78be31d90278bb792225e");

            var board = trello.Boards.WithId("55a8cdfd9536d1d4a332691f");
            var backlog = trello.Lists
                .ForBoard(board)
                .FirstOrDefault(a => a.Name == "Backlog");
            foreach (var devCard in list)
            {
                Console.WriteLine(devCard.ToString());
                var cc = new NewCard(devCard.ToString(), backlog);
                cc.Desc = string.Format("Feature:{0} Priority:{1} Estimated Hours:{2} Notes:{3}", devCard.Feature,devCard.Priority, devCard.EstimatedHours, devCard.Notes);
                trello.Cards.Add(cc);
            }
        }

        private static List<DevCard> ExtractDevCards()
        {
            var pck = new ExcelPackage(new FileInfo(File));

            var worksheet = pck.Workbook.Worksheets.First();

            Console.WriteLine(worksheet.Name);

            var list = new List<DevCard>();

            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;
            for (var row = start.Row; row <= end.Row; row++)
            {
                if (row <= 1) continue;
                var dc = new DevCard();
                for (var col = start.Column; col <= end.Column; col++)
                {
                    // ... Cell by cell...
                    object cellValue = worksheet.Cells[row, col].Text; // This got me the actual value I needed.
                    Debug.WriteLine(cellValue);

                    switch (col)
                    {
                        case 1:
                            dc.Epic = Convert.ToString(cellValue);
                            break;
                        case 2:
                            dc.Feature = Convert.ToString(cellValue);
                            break;
                        case 3:
                            dc.AsA = Convert.ToString(cellValue);
                            break;
                        case 4:
                            dc.IWantTo = Convert.ToString(cellValue);
                            break;
                        case 5:
                            dc.SoThat = Convert.ToString(cellValue);
                            break;
                        case 6:
                            dc.Priority = Convert.ToString(cellValue);
                            break;
                        case 7:
                            var ss = Convert.ToString(cellValue);
                            var no = string.IsNullOrEmpty(ss) ? "5" : ss;
                            dc.EstimatedHours = Convert.ToInt16(no);
                            break;
                        case 8:
                            dc.Notes = Convert.ToString(cellValue);
                            break;
                    }
                }
                list.Add(dc);
            }
            return list;
        }
    }
}