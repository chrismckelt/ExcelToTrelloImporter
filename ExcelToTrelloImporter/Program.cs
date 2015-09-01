using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
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
        private static ITrello _trello;
        private static List<Card> _cards;
        private static Dictionary<Card, DevCard> _dic;
        private static int _count = 1;

        private const string File =
            @"C:\dev\ExcelToTrelloImporter\ExcelToTrelloImporter\UserStories.xlsx";

        //private const string File =
        //     @"C:\work\Dropbox\FGF CloudLending\5. Requirements\FGF Application form\User Stories.xlsx";


        private static void Main(string[] args)
        {
            const string milestone = "Screens";
            var list = ExtractDevCards().Where(a => a.Milestone == milestone && a.EstimatedHours > 0).ToList();

            _trello = new Trello("7b17eb1ed849a91c051da9c924f93cfb");
            var url = _trello.GetAuthorizationUrl("userstorydataloader", Scope.ReadWrite);
            //Process.Start(url.AbsoluteUri);
            _trello.Authorize("db2c728bfd1b4cca3e07c0176e6ac3208fd4f363f383f9e0a2ac74081da4cd95");

            var board = _trello.Boards.WithId("55a8cdfd9536d1d4a332691f");
            var backlog = _trello.Lists
                .ForBoard(board)
                .FirstOrDefault(a => a.Name == "Backlog");

            _lbls = _trello.Labels.ForBoard(board);

            AddCards(list, backlog, _trello, _count);
            Thread.Sleep(5000);
            AddAcceptanceCriteria(backlog, board);
            Thread.Sleep(5000);
            AddLabels(backlog, board);
        }

        private static void AddLabels(List backlog, Board board)
        {
            foreach (var card in _trello.Cards.ForList(backlog))
            {
                if (!_dic.Any(a => a.Key.Name == card.Name)) continue;
                var xxx = _dic.FirstOrDefault(a => a.Key.Name == card.Name);
                if (xxx.Key!=null)
                {
                    SetPriority(xxx.Value, card, _count);
                    _count++;
 
                    SetTrelloLabel(xxx.Value, card);

                   // _trello.Cards.Update(card);
                    Thread.Sleep(500);
                }
            }
        }

        private static void AddAcceptanceCriteria(List backlog, Board board)
        {
            foreach (var card in _trello.Cards.ForList(backlog))
            {
                var cl = _trello.Checklists.Add("Acceptance Criteria", board);
                cl.CheckItems.Add(new CheckItem {Id = Guid.NewGuid().ToString(), Name = "Enter test 1 here...", Pos = 1});

                _trello.Checklists.AddCheckItem(cl, "Add test criteria here...");

                _trello.Cards.AddChecklist(card, cl);
               
                _trello.Cards.Update(card);
                _trello.Checklists.Update(cl);
            }
        }

        private static void AddCards(List<DevCard> list, List backlog, ITrello trello, int count)
        {
            _cards = new List<Card>();
            _dic = new Dictionary<Card, DevCard>();

            foreach (var devCard in list)
            {
                Console.WriteLine(devCard.ToString());
                var cardname = GetCardName(devCard);

                var cc = new NewCard(cardname, backlog);
                var msg = devCard.ToString();
                msg += Environment.NewLine + Environment.NewLine +
                       string.Format("Feature:{0} Priority:{1}    {2}",
                           devCard.Feature + Environment.NewLine, devCard.Priority + Environment.NewLine,
                           devCard.Notes + Environment.NewLine);
                cc.Desc = msg;
     
                var card =
                  trello.Cards.ForList(backlog).FirstOrDefault(a => a.Name.ToLowerInvariant() == cardname.ToLowerInvariant());

                if (card == null)
                {
                    card = trello.Cards.Add(cc);
                }

                _cards.Add(card);
                _dic.Add(card, devCard);
            }
        }


        private static void SetPriority(DevCard devCard, Card card, int count)
        {
            switch (devCard.Priority.ToLowerInvariant())
            {
                case "must":
                    card.Pos = 1*count;
                     break;
                case "should":
                    card.Pos = 2*count;
                     break;
                case "could":
                    card.Pos = 3*count;
                    break;
                default:
                    card.Pos = 4*count;
                    break;
            }
        }


        private static void SetTrelloLabel(DevCard ddd, Card bl)
        {
            foreach (var label in _lbls)
            {
                if (bl.Labels.Any(x => x.Name.ToLowerInvariant() == label.Name.ToLowerInvariant())) continue;

                if (ddd.ToString().ToLowerInvariant().Contains(label.Name.ToLowerInvariant()))
                {
                    bl.Labels.Add(label);
                     _trello.Cards.AddLabel(bl, label.Color.Value);
                }

                if (ddd.Priority.ToLowerInvariant().Contains(label.Name.ToLowerInvariant()))
                {
                    bl.Labels.Add(label);
                    _trello.Cards.AddLabel(bl, label.Color.Value);
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

            cardname = "(" + devCard.EstimatedHours + ") " + cardname;
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
                if (row > 50) break;

                var dc = new DevCard();
                for (var col = start.Column; col <= end.Column; col++)
                {
                    // ... Cell by cell...
                    object cellValue = worksheet.Cells[row, col].Text; // This got me the actual value I needed.
                    Debug.WriteLine(cellValue);
                    if (col > 50) break;
                    switch (col)
                    {
                        case 1:
                            dc.Milestone = Convert.ToString(cellValue);
                            break;
                        case 3:
                            dc.Feature = Convert.ToString(cellValue);
                            break;
                        case 4:
                            dc.AsA = Convert.ToString(cellValue);
                            break;
                        case 5:
                            dc.IWantTo = Convert.ToString(cellValue);
                            break;
                        case 6:
                            dc.SoThat = Convert.ToString(cellValue);
                            break;
                        case 7:
                            dc.Priority = Convert.ToString(cellValue);
                            break;
                        case 8:
                            var ss = Convert.ToString(cellValue);
                            var no = string.IsNullOrEmpty(ss) ? "5" : ss;
                            dc.EstimatedHours = Convert.ToInt16(no);
                            break;
                        case 9:
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