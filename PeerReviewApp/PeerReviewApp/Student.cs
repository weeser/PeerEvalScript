using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeerReviewApp
{
    public enum Roster
    {
        LastName = 1,
        FirstName,
        WID,
        Section,
        Email,
        DegreeProgram,
        ClassLevel,
        Phone
    }
    public class Student
    {
        public List<int> SelfScores { get; set; }
        public Dictionary<string, List<int>> TeamScores { get; set; }
        public bool CompletedSurvey { get; set; }
        public string Name { get; set; }
        public StringBuilder CommentsByStudent { get; set; }
        public StringBuilder CommentsByTeam { get; set; }

        public Student(string name)
        {
            Name = name;
            SelfScores = new List<int>();
            TeamScores = new Dictionary<string, List<int>>();
            CommentsByStudent = new StringBuilder();
            CommentsByTeam = new StringBuilder();
        }
    }
}
