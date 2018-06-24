using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Routing;

namespace ReportCreation_2._0
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Конфигурация и службы веб-API

            // Маршруты веб-API


            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

           /* config.Routes.MapHttpRoute(
                name: "AuthentificateRoute",
                routeTemplate: "api/Attendance/DownloadAttendance",
                defaults: new
                {
                    controller = "AttendanceController",
                    action = "DownloadAttendance"
                },
                constraints: new { httpMethod = new HttpMethodConstraint(HttpMethod.Post) }
            );*/

            config.Formatters.Remove(config.Formatters.XmlFormatter);
        }
    }
}
