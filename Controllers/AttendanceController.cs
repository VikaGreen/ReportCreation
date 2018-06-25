using ReportCreation_2._0.Models;
using ReportCreation_2._0.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.ModelBinding;

//using System.Web.Mvc;
//using System.Web.Mvc;

namespace ReportCreation_2._0.Controllers
{
    
    public class AttendanceController : ApiController
    {
        
        [HttpPost]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        public HttpResponseMessage DownloadAttendance(AttendanceModel json)
        {
            if (ModelState.IsValid)
            {                
                AttendanceService att = new AttendanceService();
                att.SaveWord(json, "C:\\Users\\Ivan\\Desktop\\AttendanceFile.docx");


                return Post("C:\\Users\\Ivan\\Desktop\\AttendanceFile.docx", "application/docx");
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
