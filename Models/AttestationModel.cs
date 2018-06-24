using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace ReportCreation_2._0.Models
{
    [DataContract]
    public class AttestationModel
    {
        [DataMember]
        public string semester { get; set; }
        [DataMember]
        public int year { get; set; }
        [DataMember]
        public string department { get; set; }
        [DataMember]
        public string speciality { get; set; }
        [DataMember]
        public string subject { get; set; }
        [DataMember]
        public string teacher { get; set; }
        [DataMember]
        public List<AttestationRecord> attestationRecords { get; set; }
    }

    public class Mark
    {
        public string FIO { get; set; }
        public int mark { get; set; }
    }

    public class AttestationRecord
    {
        public int course { get; set; }
        public int group { get; set; }
        public string date { get; set; }
        public string contingentOfStudents { get; set; }
        public List<Mark> marks { get; set; }
    }

}