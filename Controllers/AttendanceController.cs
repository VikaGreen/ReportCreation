using ReportCreation_2._0.Models;
using ReportCreation_2._0.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.ModelBinding;

//using System.Web.Mvc;
//using System.Web.Mvc;

namespace ReportCreation_2._0.Controllers
{
    
    public class AttendanceController : ApiController
    {
        
        [HttpPost]
        public void DownloadAttendance(AttendanceModel json)
        {
            if (ModelState.IsValid)
            {                
                AttendanceService att = new AttendanceService();
                att.SaveWord(json, "C:/AttendanceFile.docx");

               /* string filePath = System.IO.Path.Combine(Server.MapPath("~/Documents"), "xxx.jpg");
                FileInfo file = new FileInfo(filePath);
                HttpContext.Current.Response.AddHeader("Content-Length", file.Length.ToString());
                HttpContext.Current.Response.AddHeader("Connection", "Keep-Alive");
                HttpContext.Current.Response.ContentType = "image/jpeg";
                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + file.Name);
                HttpContext.Current.Response.WriteFile(file.FullName);
                HttpContext.Current.Response.End();*/
                var f = File.ReadAllBytes("C:/AttendanceFile.docx");
                
                HttpContext.Current.Response.ClearContent();
                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=AttendanceFile.docx");
                //HttpContext.Current.Response.AddHeader("Accept", "application/docx");
                HttpContext.Current.Response.ContentType = "application/docx";

                /*string importAddress = "C:/AttendanceFile.docx";
                using (var web = new WebClient())
                using (var importStream = web.OpenRead(importAddress))
                {
                    importStream.CopyTo(HttpContext.Current.Response.OutputStream);
                    
                }*/
                HttpContext.Current.Response.WriteFile("C:/AttendanceFile.docx");
                HttpContext.Current.Response.End();
                


            }

           // return null;
        }

       // [System.Web.Http.HttpPost]
        /*public FileResult GetFileAttendance()
        {
            //_corporation = Corporation.OpenJson(OFD.FileName);
            AttendanceModel att;
            var serializer = new DataContractJsonSerializer(typeof(AttendanceModel));
            using (var fileStream = new FileStream("C:/Users/Администратор/Documents/Visual Studio 2017/Projects/ReportCreation/ReportCreation/Files/attendance.json", FileMode.Open))
            {
                att = (AttendanceModel)serializer.ReadObject(fileStream);
                fileStream.Close();
            }

            Microsoft.Office.Interop.Word._Application word = new Microsoft.Office.Interop.Word.Application();
            var document = word.Documents.Add();
            var fitBehavior = Type.Missing;
            document.Paragraphs.Add();
            document.Paragraphs[1].Range.Text = att.semester;//"Hello world";//att.Semester;
            document.SaveAs("~Files/AttendanceFile.docx");
            document.Close();
            word.Quit();

            //HttpContext.Response.ContentType = "application/docx";
            FileContentResult result = new FileContentResult(System.IO.File.ReadAllBytes("~Files/AttendanceFile.docx"), "application/docx")
            {
                FileDownloadName = "AttendanceFile.docx"
            };

            return result;
        }
        */
    }
}
