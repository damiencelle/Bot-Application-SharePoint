using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bot_Application2.Luis
{
    public class Intent
    {
        public string intent { get; set; } // Name is important to deserialize JSON

        public float Score { get; set; }
    }
}