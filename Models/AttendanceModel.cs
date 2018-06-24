using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace ReportCreation_2._0.Models
{
    [DataContract]
    public class AttendanceModel
    {
        [DataMember]
        public string semester { get; set; }
        [DataMember]
        public int year { get; set; }
        [DataMember]
        public string branch { get; set; }
        [DataMember]
        public string department { get; set; }
        [DataMember]
        public string level { get; set; }
        [DataMember]
        public string speciality { get; set; }
        [DataMember]
        public string subject { get; set; }
        [DataMember]
        public string teacher { get; set; }
        [DataMember]
        public int course { get; set; }
        [DataMember]
        public int group { get; set; }
        [DataMember]
        public List<Student> students { get; set; }
    }

    public class Record
    {
        public string date { get; set; }
        public string note { get; set; }
    }

    public class Student
    {
        public string FIO { get; set; }
        public List<Record> records { get; set; }
        public string note { get; set; }
    }
}