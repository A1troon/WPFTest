using System;

namespace WPFTest
{
    class Test
    {
        public string Question { get; set; }
        public int Answer { get; set; }
        public int AnswerChecked { get; set; }
        public string []V {get; set;}

        public Test() { }

        public Test(string question, int answer, string[] v)
        {
            Question = question;
            Answer = answer;
            V = v;
            AnswerChecked = 0;
        }

      
    }
}
