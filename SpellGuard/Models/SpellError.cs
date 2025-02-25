using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpellGuard.Models
{
    public class SpellError
    {
        public string WordFileName { get; set; }
        public string WrongSpell { get; set; }
        public int PageNumber { get; set; }
        public int LineNumber { get; set; }
        public int Position { get; set; }
        public string SuggestedWords { get; set; }
    }
}
