using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToTrelloImporter
{
    public class DevCard
    {
        public string Milestone { get; set; } 
        public string Feature { get; set; }
        public string AsA { get; set; } 
        public string IWantTo { get; set; } 

        public string SoThat { get; set; } 

        public string Priority { get; set; } 

        public decimal EstimatedHours { get; set; } 

        public string Notes { get; set; }

        public override string ToString()
        {
            if (AsA.Contains("Marketing"))
            {
                return $"As {AsA}, I want to {IWantTo}, so that {SoThat}";
            }

            if (AsA.Contains("System"))
            {
                return $"As the {AsA}, I want to {IWantTo}, so that {SoThat}";
            }

            return $"As a {AsA}, I want to {IWantTo}, so that {SoThat}";
        }

        
        public string FullString()
        {
            return string.Format("{0}, {1}, {2}, {3}", ToString(), Priority, Notes, Feature);
        }
    }
}
