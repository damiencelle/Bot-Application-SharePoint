using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bot_Application2.Luis
{
    public class LuisJSON
    {
        public string Query { get; set; }

        public Intent[] Intents { get; set; }

        public Entity[] Entities { get; set; }

        public Intent TopScoringIntent { get; set; }
    }
}