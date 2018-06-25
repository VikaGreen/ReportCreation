using ReportCreation_2._0.Models;
using ReportCreation_2._0.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;

namespace ReportCreation_2._0.Controllers
{
    public class AttestationController : ApiController
    {
        [HttpPost]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        public HttpResponseMessage DownloadAttestation(AttestationModel json)
        {
            if (ModelState.IsValid)
            {
                AttestationService att = new AttestationService();
                att.SaveWord(json, "C:\\Users\\Ivan\\Desktop\\AttestationFile.docx");

                return Post("C:\\Users\\Ivan\\Desktop\\AttestationFile.docx", "application/docx");
            }

            return null;
        }

        private HttpResponseMessage Post(string path, string contentType)
        {
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            return result;
        }
    }
}
