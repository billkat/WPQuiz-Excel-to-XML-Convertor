///Questions and Answers Class
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToXMLwpQuiz
{
    public class Questions
    {
        public string answerType { get; set; }
        public string title { get; set; }
        public int points { get; set; }
        public string questionText { get; set; }
        public string correctMsg { get; set; }
        public string incorrectMsg { get; set; }
        public string category { get; set; }
    }

    public class Answers
    {
        public int points { get; set; }
        public string correct { get; set; }
        public string answerText { get; set; }
        public string stortText { get; set; }
    }
}
