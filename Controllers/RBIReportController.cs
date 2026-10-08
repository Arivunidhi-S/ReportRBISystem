using Microsoft.SqlServer.Server;
using Stimulsoft.Base.Design;
using Stimulsoft.Controls.Win.DotNetBar;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Export;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Web.Hosting;
using System.Web.Http;
using System.Web.UI.WebControls;

namespace ReportRBISystem.Controllers
{
    [RoutePrefix("api/RBIReport")]
    public class RBIReportController : ApiController
    {
        [HttpGet]
        [Route("test")]
        public IHttpActionResult Test()
        {
            return Ok(new
            {
                Success = true,
                Message = "RBIreport is working",
                StimulsoftVersion = "2011 DLL"
            });
        }


        [HttpGet]
        [Route("EquipmentList")]
        public HttpResponseMessage EquipmentListPdf(int id, string format = "pdf")
        {
            // SQL இன்ஜெக்ஷன் தவிர்ப்பதற்காக SQL parameter-ஆக மாற்றுவது நல்லது, தற்போதைக்கு உங்கள் தற்போதைய குறியீட்டின்படி வைக்கப்பட்டுள்ளது
            string sql = $@"SELECT ROW_NUMBER() OVER(ORDER BY DoshNo DESC) as No, ProcessArea, DoshNo, EqupID, EqupDescription FROM VW__EquipmentAsset WHERE ProcessAreaID={id}";

            string datasourceName = "VW__EquipmentAsset";
            string mrtFileName = "EquipmentList.mrt";

            // 4-வது ஆர்குமெண்ட்டாக 'format' பராமீட்டர் அனுப்பப்படுகிறது
            return ReportHelper.GeneratePdfReport(Request, sql, datasourceName, mrtFileName, format);
        }


        [HttpGet]
        [Route("EquipmentDetails")]
        public HttpResponseMessage EquipmentDetailsPdf(int id, string format = "pdf")
        {
           
            string sql = $@"select * from VW_EquipmentDetails where CompAutoID={id}" ;           
            string datasourceName = "VW_EquipmentDetails";

            string sql2 = $@"select * from VW_InspectionPlanForEquipmentDetails where CompID={id}";
            string datasourceName2 = "VW_InspectionPlan";

            string mrtFileName = "EquipmentDetails.mrt";         
            return ReportHelper.GeneratePdfReport2Val(Request, sql, sql2, datasourceName, datasourceName2, mrtFileName);
        }

        [HttpGet]
        [Route("InspectionDetails")]
        public HttpResponseMessage InspectionDetailsPdf(int id, string format = "pdf")
        {

            string sql = $@"select *,CONVERT(VARCHAR(10), Initialdate, 103) AS [Indate],Initialvalue,CONVERT(VARCHAR(10), InspecDate, 103) AS
                            [Insdate],CONVERT(VARCHAR(10), Previousdate, 103) AS [Predate] from VW_Inspection where CompAutoID={id} and Deleted=0 
                                order by EquAutoID,InspectionPointNo,InspecDate";
            string datasourceName = "VW_Inspection";
            string mrtFileName = "Inspection.mrt";
            return ReportHelper.GeneratePdfReport(Request, sql, datasourceName, mrtFileName, format);
        }

        [HttpGet]
        [Route("InspectionChart")]
        public HttpResponseMessage InspectionChartPdf(int id, string format = "pdf")
        {

            string sql = $@"select MRT,CONVERT(VARCHAR(10), InspecDate, 103) AS InsDate,InspecDate,ReadingValue,NormalThickness,[Equipment ID],DoshNo,CompName,uCR,RemainingLife 
                            from VW_InspectionChart where CompAutoID={id} and   Deleted=0 and LongCRrate is not null order by EqupID,InspecDate";
            string datasourceName = "VW_InspectionChart";
            string mrtFileName = "Chart_Inspection.mrt";
            return ReportHelper.GeneratePdfReport(Request, sql, datasourceName, mrtFileName, format);
        }

        [HttpGet]
        [Route("POFList")]
        public HttpResponseMessage POFListPdf(int id, string format = "pdf")
        {

            string sql = $@"select * from VW_POF where ProcessareaID={id}";
            string datasourceName = "VW_POF";
            string mrtFileName = "POFList.mrt";
            return ReportHelper.GeneratePdfReport(Request, sql, datasourceName, mrtFileName, format);
        }

        [HttpGet]
        [Route("COFFlammableDetails")]
        public HttpResponseMessage COFFlammableDetailsPdf(int id, string format = "pdf")
        {

            // Get connection string from Web.config
           string connectionString = ConfigurationManager.ConnectionStrings["RBIDB"]?.ConnectionString;
            string sql = $@"select * from VW_COF_Flammable where CompID={id}";
            string datasourceName = "COF_Flammable";        

            string mrtFileName = null;

            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmddup = new SqlCommand(sql, conn);
            SqlDataReader readerdup = cmddup.ExecuteReader();

            if (readerdup.Read())
            {
                double ptrans = Convert.ToDouble(readerdup["ptrans"].ToString());
                double oppress = Convert.ToDouble(readerdup["OpPres"].ToString());


                if (readerdup["Fluid"].ToString() == "Liquid")
                {
                    mrtFileName = "COF_Flammable_Liquid.mrt";
                }
                else
                {
                    if (oppress < ptrans)
                    {
                        mrtFileName = "COF_Flammable_Vapour_Sonic.mrt";
                    }
                    else
                    {
                        mrtFileName = "COF_Flammable_Vapour_SubSonic.mrt";
                    }
                }
            }
            readerdup.Close();
            conn.Close();

            return ReportHelper.GeneratePdfReport(Request, sql,  datasourceName, mrtFileName, format);
        }

        [HttpGet]
        [Route("COFNonFlammableDetails")]
        public HttpResponseMessage COFNonFlammableDetailsPdf(int id, string format = "pdf")
        {

            // Get connection string from Web.config
            string connectionString = ConfigurationManager.ConnectionStrings["RBIDB"]?.ConnectionString;
            string sql = $@"select * from VW_COF_NonFlammable where CompID={id}";
            string datasourceName = "COF_NonFlammable";

            string mrtFileName = null;

            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmddup = new SqlCommand(sql, conn);
            SqlDataReader readerdup = cmddup.ExecuteReader();

            if (readerdup.Read())
            {
                double ptrans = Convert.ToDouble(readerdup["ptrans"].ToString());
                double oppress = Convert.ToDouble(readerdup["OpPres"].ToString());


                if (readerdup["Fluid"].ToString() == "Liquid")
                {
                    mrtFileName = "COF_Non_Flammable_Liquid.mrt";
                }
                else
                {
                    if (oppress < ptrans)
                    {
                        mrtFileName = "COF_Non_Flammable_Vapour_Sonic.mrt";
                    }
                    else
                    {
                        mrtFileName = "COF_Non_Flammable_Vapour_SubSonic.mrt";
                    }
                }
            }
            readerdup.Close();
            conn.Close();

            return ReportHelper.GeneratePdfReport(Request, sql, datasourceName, mrtFileName, format);
        }

        //[HttpGet]
        //[Route("EquipmentList")]
        //public HttpResponseMessage EquipmentListPdf(int id)
        //{
        //    StiReport report = null;

        //    try
        //    {
        //        // Get connection string from Web.config
        //        string connectionString = ConfigurationManager.ConnectionStrings["RBIDB"]?.ConnectionString;

        //        if (string.IsNullOrWhiteSpace(connectionString))
        //        {
        //            return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, "RBIDB connection string was not found.");
        //        }

        //        // Load Students data               
        //        DataSet dataSet = new DataSet();
        //        string datasourceName = "VW__EquipmentAsset";
        //        dataSet.DataSetName = datasourceName;
        //        string sql = @"SELECT ROW_NUMBER() OVER(ORDER BY DoshNo DESC) as No,ProcessArea,DoshNo,EqupID,EqupDescription FROM VW__EquipmentAsset where ProcessAreaID="+id;
        //        //command.Parameters.AddWithValue("@id", id);

        //        SqlDataAdapter adapter = new SqlDataAdapter(sql, connectionString);
        //        dataSet.Tables.Add(datasourceName);
        //        adapter.Fill(dataSet, datasourceName);
        //        string reportPath = HostingEnvironment.MapPath("~/Reports/EquipmentList.mrt");

        //        if (string.IsNullOrWhiteSpace(reportPath) || !File.Exists(reportPath))
        //        {
        //            return Request.CreateErrorResponse(HttpStatusCode.NotFound, "EquipmentList.mrt report file was not found.");
        //        }

        //        // Load report
        //        report = new StiReport();
        //        report.Load(reportPath);

        //        // Remove the embedded database connection
        //        report.Dictionary.Databases.Clear();

        //        // Add current SQL connection to the report
        //        report.Dictionary.Databases.Add(new StiSqlDatabase("Connection", connectionString));

        //        report.Dictionary.DataSources.Clear();
        //        // Register data
        //        report.RegData(datasourceName, dataSet);

        //        // Synchronize report dictionary
        //        report.Dictionary.Synchronize();

        //        // Compile and render report
        //        report.Compile();
        //        report.Render();

        //        // Export PDF directly into memory         
        //        using (MemoryStream memoryStream = new MemoryStream())
        //        {
        //            StiPdfExportService pdfExportService = new StiPdfExportService();
        //            pdfExportService.ExportPdf(report, memoryStream);

        //            // ⭐ மிக முக்கிய மாற்றம்: மெமரி ஸ்ட்ரீமை ஆரம்ப இடத்திற்கு கொண்டு செல்ல வேண்டும்
        //            memoryStream.Position = 0;

        //            byte[] pdfBytes = memoryStream.ToArray();

        //            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        //            response.Content = new ByteArrayContent(pdfBytes);

        //            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

        //            // ⭐ டவுன்லோட் ஆகாமல் iframe-க்குள் தெரிய 'inline' என்று மாற்றப்பட்டுள்ளது
        //            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline")
        //            {
        //                FileName = "EquipmentList.pdf"
        //            };

        //            // CORS Headers (Blazor-ல் fetch வேலை செய்ய)
        //            response.Headers.Add("Access-Control-Allow-Origin", "*");
        //            response.Headers.Add("Access-Control-Allow-Methods", "GET");
        //            response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Accept");

        //            return response;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        return Request.CreateErrorResponse(
        //            HttpStatusCode.InternalServerError,
        //            ex);
        //    }
        //    finally
        //    {
        //        if (report != null)
        //        {
        //            report.Dispose();
        //        }
        //    }
        //}


        [HttpGet]
        [Route("students-pdf")]
        public HttpResponseMessage StudentsPdf()
        {
            StiReport report = null;

            try
            {
                string connectionString =
                    ConfigurationManager.ConnectionStrings["StudentDb"]
                    ?.ConnectionString;

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    return Request.CreateErrorResponse(
                        HttpStatusCode.InternalServerError,
                        "StudentDb connection string is missing.");
                }

                // Create DataSet
                DataSet dataSet = new DataSet("StudentsDataSet");

                DataTable studentsTable = new DataTable("Students");

                string sql = @"
            SELECT
                Id,
                Name,
                Age,
                Department,
                Email
            FROM Students";

                using (SqlConnection connection =
                       new SqlConnection(connectionString))
                using (SqlCommand command =
                       new SqlCommand(sql, connection))
                using (SqlDataAdapter adapter =
                       new SqlDataAdapter(command))
                {
                    adapter.Fill(studentsTable);
                }

                if (studentsTable.Rows.Count == 0)
                {
                    return Request.CreateErrorResponse(
                        HttpStatusCode.NotFound,
                        "No student records found.");
                }

                dataSet.Tables.Add(studentsTable);

                // Get report path
                string reportPath =
                    System.Web.Hosting.HostingEnvironment.MapPath(
                        "~/Reports/Students.mrt");

                if (string.IsNullOrWhiteSpace(reportPath) ||
                    !File.Exists(reportPath))
                {
                    return Request.CreateErrorResponse(
                        HttpStatusCode.NotFound,
                        "Students.mrt file was not found.");
                }

                // Create report
                report = new StiReport();

                // Load MRT file
                report.Load(reportPath);

                // Remove SQL connection stored in MRT
                report.Dictionary.Databases.Clear();
                report.Dictionary.Databases.Add(new StiSqlDatabase("Connection", connectionString));
                // Register DataSet with the same name as the report DataSource
                report.RegData("Students", dataSet);

                // Synchronize dictionary
                report.Dictionary.Synchronize();

                // Compile and render
                report.Compile();
                report.Render();

                // Export PDF
                using (MemoryStream stream = new MemoryStream())
                {
                    StiPdfExportService pdfExportService =
                        new StiPdfExportService();

                    pdfExportService.ExportPdf(report, stream);

                    stream.Position = 0;

                    byte[] pdfBytes = stream.ToArray();

                    HttpResponseMessage response =
                        new HttpResponseMessage(HttpStatusCode.OK);

                    response.Content =
                        new ByteArrayContent(pdfBytes);

                    response.Content.Headers.ContentType =
                        new MediaTypeHeaderValue("application/pdf");

                    response.Content.Headers.ContentDisposition =
                        new ContentDispositionHeaderValue("attachment")
                        {
                            FileName = "Students.pdf"
                        };

                    return response;
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(
                    HttpStatusCode.InternalServerError,
                    ex);
            }
            finally
            {
                if (report != null)
                {
                    report.Dispose();
                }
            }
        }

        [HttpGet]
        [Route("EquipLst-pdf")]
        public HttpResponseMessage EquipLstPdf()
        {
            StiReport report = null;

            try
            {
                string connectionString =
                    ConfigurationManager.ConnectionStrings["RBIDB"]
                    ?.ConnectionString;

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    return Request.CreateErrorResponse(
                        HttpStatusCode.InternalServerError,
                        "RBIDB connection string is missing.");
                }

                // Create DataSet
                DataSet dataSet = new DataSet("StudentsDataSet");

                DataTable studentsTable = new DataTable("VW__EquipmentAsset");

                string sql = @"SELECT ROW_NUMBER() OVER(ORDER BY DoshNo DESC) as No,ProcessArea,DoshNo,EqupID,EqupDescription FROM VW__EquipmentAsset where companyid=2 and ProcessAreaID=4";

                using (SqlConnection connection =
                       new SqlConnection(connectionString))
                using (SqlCommand command =
                       new SqlCommand(sql, connection))
                using (SqlDataAdapter adapter =
                       new SqlDataAdapter(command))
                {
                    adapter.Fill(studentsTable);
                }

                if (studentsTable.Rows.Count == 0)
                {
                    return Request.CreateErrorResponse(
                        HttpStatusCode.NotFound,
                        "No student records found.");
                }

                dataSet.Tables.Add(studentsTable);

                // Get report path
                string reportPath =
                    System.Web.Hosting.HostingEnvironment.MapPath(
                        "~/Reports/EquipmentList.mrt");

                if (string.IsNullOrWhiteSpace(reportPath) ||
                    !File.Exists(reportPath))
                {
                    return Request.CreateErrorResponse(
                        HttpStatusCode.NotFound,
                        "Students.mrt file was not found.");
                }

                // Create report
                report = new StiReport();

                // Load MRT file
                report.Load(reportPath);

                // Remove SQL connection stored in MRT
                report.Dictionary.Databases.Clear();
                report.Dictionary.Databases.Add(new StiSqlDatabase("Connection", connectionString));
                report.Dictionary.DataSources.Clear();
                // Register DataSet with the same name as the report DataSource
                report.RegData("VW__EquipmentAsset", dataSet);

                // Synchronize dictionary
                report.Dictionary.Synchronize();

                // Compile and render
                report.Compile();
                report.Render();

                // Export PDF
                using (MemoryStream stream = new MemoryStream())
                {
                    StiPdfExportService pdfExportService =
                        new StiPdfExportService();

                    pdfExportService.ExportPdf(report, stream);

                    stream.Position = 0;

                    byte[] pdfBytes = stream.ToArray();

                    HttpResponseMessage response =
                        new HttpResponseMessage(HttpStatusCode.OK);

                    response.Content =
                        new ByteArrayContent(pdfBytes);

                    response.Content.Headers.ContentType =
                        new MediaTypeHeaderValue("application/pdf");

                    response.Content.Headers.ContentDisposition =
                        new ContentDispositionHeaderValue("attachment")
                        {
                            FileName = "EquipLst.pdf"
                        };

                    return response;
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(
                    HttpStatusCode.InternalServerError,
                    ex);
            }
            finally
            {
                if (report != null)
                {
                    report.Dispose();
                }
            }
        }
    }
}
