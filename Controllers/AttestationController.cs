using ReportCreation_2._0.Models;
using ReportCreation_2._0.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;

namespace ReportCreation_2._0.Controllers
{
    public class AttestationController : ApiController
    {
        [HttpPost]
        public byte[] DownloadAttestation(AttestationModel json)
        {
            if (ModelState.IsValid)
            {
                AttestationService att = new AttestationService();
                att.SaveWord(json, "C:/AttestationFile.docx");

                var f = System.IO.File.ReadAllBytes("C:/AttestationFile.docx");
                HttpContext.Current.Response.ContentType = "application/docx";

                return f;
            }

            return null;
        }
    }
}
