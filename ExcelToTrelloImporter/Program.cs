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
        private static IEnumerable<Label> _lbls;
        //private const string File =
        //    @"C:\work\Dropbox\FGF CloudLending\5. Requirements\FGF Application form\User Stories (chris mckelt's conflicted copy 2015-08-25).xlsx";

        private const string File =
            @"C:\work\Dropbox\FGF CloudLending\5. Requirements\FGF Application form\User Stories (chris mckelt's conflicted copy 2015-08-25).xlsx";


        private static void Main(string[] args)
        {
            var list = ExtractDevCards();
            var count = 1;

            ITrello trello = new Trello("7b17eb1ed849a91c051da9c924f93cfb");
            //var url = trello.GetAuthorizationUrl("userstorydataloader", Scope.ReadWrite);
            //Process.Start(url.AbsoluteUri);
            trello.Authorize("88b7bf860f1b63bcf1338e69fba56e1dbe0470db8b5e20d7567d2ae93b4da232");

            var board = trello.Boards.WithId("55a8cdfd9536d1d4a332691f");
            var backlog = trello.Lists
                .ForBoard(board)
                .FirstOrDefault(a => a.Name == "Backlog");

            _lbls = trello.Labels.ForBoard(board);
            AddCards(list, backlog, trello, count);
          
            foreach (var s in trello.Cards.ForList(backlog))
            {
                s.Badges.Votes = 10;
                var cl = trello.Checklists.Add("Acceptance Criteria", board);
                cl.CheckItems.Add(new CheckItem() { Id = Guid.NewGuid().ToString(), Name = "Enter test 1 here...", Pos = 1 });

                trello.Checklists.AddCheckItem(cl, "Add test criteria here...");

                trello.Cards.AddChecklist(s,cl);
                trello.Cards.Update(s);
                trello.Checklists.Update(cl);
            }

            
        }

        private static void AddCards(List<DevCard> list, List backlog, ITrello trello, int count)
        {
            foreach (var devCard in list)
            {
                Console.WriteLine(devCard.ToString());
                var cardname = GetCardName(devCard);

                var cc = new NewCard(cardname, backlog);
                var msg = devCard.ToString();
                msg += Environment.NewLine + Environment.NewLine +
                       string.Format("Feature:{0} Priority:{1} Estimated Hours:{2} Notes:{3}",
                           devCard.Feature + Environment.NewLine, devCard.Priority + Environment.NewLine,
                           devCard.EstimatedHours + Environment.NewLine, devCard.Notes + Environment.NewLine);
                cc.Desc = msg;

                var card = trello.Cards.Add(cc);

                SetTrelloLabel(devCard, card, _lbls);
                SetPriority(devCard, card, count);
                trello.Cards.Update(card);
                count ++;
            }
        }

        private static void SetPriority(DevCard devCard, Card card, int count)
        {
            switch (devCard.Priority.ToLowerInvariant())
            {
                case "must":
                    card.Pos = 1*count;
                    SetLabel(card, () => devCard.Priority.ToLowerInvariant().Contains("must"), "Must");
                    break;
                case "should":
                    card.Pos = 2*count;
                    SetLabel(card, () => devCard.Priority.ToLowerInvariant().Contains("should"), "Should");
                    break;
                case "could":
                    card.Pos = 3*count;
                    SetLabel(card, () => devCard.Priority.ToLowerInvariant().Contains("could"), "Could");
                    break;
                default:
                    card.Pos = 4*count;
                    break;
            }
        }

        private static void SetLabel(Card card, Func<bool> cardContains, string label)
        {
            var yep = cardContains();
            if (yep)
                card.Labels.Add(_lbls.Single(y => y.Name == label));
        }

        private static void SetTrelloLabel(DevCard ddd, Card bl, IEnumerable<Label> lbls)
        {
            foreach (var label in lbls)
            {
                if (ddd.ToString().ToLowerInvariant().Contains(label.Name.ToLowerInvariant()))
                {
                    bl.Labels.Add(label);
                }
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