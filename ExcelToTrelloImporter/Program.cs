using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
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

       // private const string File = @"C:\dev\ExcelToTrelloImporter\ExcelToTrelloImporter\UserStories.xlsx";

          private const string File = @"C:\work\Dropbox\FGF CloudLending\5. Requirements\FGF Application form\user stories_v0.8_cm.xlsx";

        private static ExcelPackage _pck;

        public static Checklist SetChecklist(DevCard card)
        {
            var checklist = _trello.Checklists.Add("Acceptance Criteria", _board);
            checklist.CheckItems.Add(new CheckItem { Id = Guid.NewGuid().ToString(), Name = "Enter test 1 here...", Pos = 1 });
            
            string text = "Add test criteria here...";
            _trello.Checklists.AddCheckItem(checklist, !string.IsNullOrEmpty(card.AcceptanceCriteria) ? card.AcceptanceCriteria : text);
            return checklist;
        }

        private static void Main(string[] args)
        {
            if (!System.IO.File.Exists(File)) throw new FileNotFoundException(File);

            _pck = new ExcelPackage(new FileInfo(File));

            GetColumnIndexes();


            const string milestone = "Screens";
            _devCards = ExtractDevCards().Where(a => a.Milestone == milestone && a.EstimatedHours > 0).ToList();

            _trello = new Trello("7b17eb1ed849a91c051da9c924f93cfb");
            var url = _trello.GetAuthorizationUrl("userstorydataloader", Scope.ReadWrite);
            //Process.Start(url.AbsoluteUri);
            _trello.Authorize("db2c728bfd1b4cca3e07c0176e6ac3208fd4f363f383f9e0a2ac74081da4cd95");

            _board = _trello.Boards.WithId("55a8cdfd9536d1d4a332691f");
            _backlog = _trello.Lists
                .ForBoard(_board)
                .FirstOrDefault(a => a.Name == "Backlog");

            _lbls = _trello.Labels.ForBoard(_board);

            AddCards(_devCards, _backlog, _trello, _count);
            Thread.Sleep(5000);
            AddAcceptanceCriteria(_backlog, _board);
            Thread.Sleep(5000);
            AddLabels(_backlog, _board);
        }

        private static void AddLabels(List backlog, Board board)
        {
            foreach (var card in _trello.Cards.ForList(backlog))
            {
                if (!_dic.Any(a => a.Key.Name == card.Name)) continue;
                var xxx = _dic.FirstOrDefault(a => a.Key.Name == card.Name);
                if (xxx.Key != null)
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
            foreach (var card in _trello.Cards.ForList(backlog).Where(a=>a.Checklists.Count==0))
            {
                var gg = _devCards.FirstOrDefault(a => GetCardName(a) == card.Name);
                var cl = SetChecklist(gg);

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
                    card.Pos = 1 * count;
                    break;
                case "should":
                    card.Pos = 2 * count;
                    break;
                case "could":
                    card.Pos = 3 * count;
                    break;
                default:
                    card.Pos = 4 * count;
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

        public static List<int> Rows = new List<int>();
        private static Dictionary<string, int> _columnIndexes;
        private static Board _board;
        private static List _backlog;
        private static List<DevCard> _devCards;

        private static List<DevCard> ExtractDevCards()
        {
                var worksheet = _pck.Workbook.Worksheets.First(x => x.Name == "Backlog");

                Console.WriteLine(worksheet.Name);

                var list = new List<DevCard>();

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                for (var row = start.Row; row <= end.Row; row++)
                {
                    if (row <= 1) continue;
                    if (row > 50) break;
                    Rows.Add(row);
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
                        default:
                                int icol;
                                if (_columnIndexes.TryGetValue("UAC", out icol))
                                    if (icol == col)
                                    {
                                        dc.AcceptanceCriteria = Convert.ToString(cellValue);
                                    }
                                break;
                        }
                    }
                    list.Add(dc);
                }
                return list;

        }

        private static Dictionary<string, int> GetColumnIndexes()
        {
            _columnIndexes = new Dictionary<string, int>();
             var workbook = _pck.Workbook;

            if (!workbook.Worksheets.Any()) return null;
            var worksheet = workbook.Worksheets.First(x => x.Name.Contains("Backlog"));
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;
            for (var col = start.Column; col <= end.Column; col++)
            {
                string txt = worksheet.Cells[1, col].Text;
                if (!string.IsNullOrEmpty(txt))
                    _columnIndexes.Add(txt, col);

                if (col > 500) break;
            }
            

            return _columnIndexes;
        }

        private static void MarkRowsAsInSprint()
        {
           
            var worksheet = _pck.Workbook.Worksheets.First(x => x.Name == "Backlog");
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;
            for (var row = start.Row; row <= end.Row; row++)
            {
                if (Rows.Contains(row))
                {
                    var txt = GetColumnIndexes().First(a => a.Value == row).Key;
                     
                }
            }
        }
    }
}