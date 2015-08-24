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
        private const string File =
            @"C:\Dropbox\FGF CloudLending\5. Requirements\FGF Application form\User Stories.xlsx";

        private static void Main(string[] args)
        {
            var list = ExtractDevCards();

            ITrello trello = new Trello("7b17eb1ed849a91c051da9c924f93cfb");
            // var url = trello.GetAuthorizationUrl("userstorydataloader", Scope.ReadWrite);
            // Process.Start(url.AbsoluteUri);
            trello.Authorize("88b7bf860f1b63bcf1338e69fba56e1dbe0470db8b5e20d7567d2ae93b4da232");

            var board = trello.Boards.WithId("55a8cdfd9536d1d4a332691f");
            var backlog = trello.Lists
                .ForBoard(board)
                .FirstOrDefault(a => a.Name == "Backlog");
            foreach (var devCard in list)
            {
                Console.WriteLine(devCard.ToString());
                var cardname = GetCardName(devCard);

                var cc = new NewCard(cardname, backlog);
                var msg = devCard.ToString();
                msg += Environment.NewLine +
                       string.Format("Feature:{0} Priority:{1} Estimated Hours:{2} Notes:{3}",
                           devCard.Feature + Environment.NewLine, devCard.Priority + Environment.NewLine,
                           devCard.EstimatedHours + Environment.NewLine, devCard.Notes + Environment.NewLine);
                cc.Desc = msg;
               // trello.Cards.Add(cc);
            }
            var pos = 1;
            var lbls = trello.Labels.ForBoard(board);
            foreach (var bl in trello.Cards.ForList(backlog))
            {
                var ddd = list.FirstOrDefault(x => GetCardName(x) == bl.Name);
                bl.Pos = pos++;
              
                SetTrelloLabel(ddd, bl, lbls);

               // bl.Checklists.Add(new Card.Checklist {Name = "Acceptance Criteria"});
                try
                {
                    trello.Cards.Update(bl);
                }
                catch (Exception)
                {
                    
                    
                }
                
            }
        }

        private static void SetTrelloLabel(DevCard ddd, Card bl, IEnumerable<Label> lbls)
        {
            switch (ddd.SoThat)
            {
                case "usability ":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Usability"));
                    break;
                case "loan servicability":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Loan servicability"));
                    break;
                case "validation":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Validation"));
                    break;
                case "compliance":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Compliance"));
                    break;
                case "prevent bad loans":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Prevent bad loans"));
                    break;
                case "loan affordability":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Loan affordability"));
                    break;
                case "customer data is captured":
                    bl.Labels.Add(lbls.Single(y => y.Name == "Backend"));
                    break;
                default:
                    bl.Labels.Add(lbls.Single(y => y.Name == "Infrastructure"));
                    break;
            }
        }

        private static string GetCardName(DevCard devCard)
        {
            string cardname = null;
            if (string.IsNullOrEmpty(devCard.IWantTo))
            {
                cardname = devCard.ToString();
            }
            else
            {
                cardname = devCard.Feature + ":   " + devCard.IWantTo;
            }

            cardname += " [" + devCard.EstimatedHours + "]";
            return cardname;
        }

        private static List<DevCard> ExtractDevCards()
        {
            var pck = new ExcelPackage(new FileInfo(File));

            var worksheet = pck.Workbook.Worksheets.First(x => x.Name == "Backlog");

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