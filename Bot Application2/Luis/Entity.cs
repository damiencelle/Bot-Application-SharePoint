using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bot_Application2.Luis
{
    public class Entity
    {
        public string entity { get; set; } // Name is important to deserialize JSON

        public string Type { get; set; }

        public int StartIndex { get; set; }

        public int EndIndex { get; set; }

        public float Score { get; set; }
    }
}