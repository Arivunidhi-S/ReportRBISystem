using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Configuration;
using System.Web.Hosting;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Export;

public class ReportHelper
{
    public static HttpResponseMessage GeneratePdfReport(HttpRequestMessage request, string sql, string datasourceName, string mrtFileName)
    {
        StiReport report = null;

        try
        {
            // Web.config-ல் இருந்து கனெக்ஷன் ஸ்டிரிங் எடுத்தல்
            string connectionString = WebConfigurationManager.ConnectionStrings["RBIDB"]?.ConnectionString;

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                return request.CreateErrorResponse(HttpStatusCode.InternalServerError, "RBIDB connection string was not found.");
            }

            // டேட்டா செட் தயார் செய்தல் (Dynamic Datasource Name)
            DataSet dataSet = new DataSet();
            dataSet.DataSetName = datasourceName;

            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connectionString))
            {
                dataSet.Tables.Add(datasourceName);
                adapter.Fill(dataSet, datasourceName);
            }

            // MRT ஃபைல் பாத் அமைத்தல் (Dynamic File Name)
            string reportPath = HostingEnvironment.MapPath($"~/Reports/{mrtFileName}");

            if (string.IsNullOrWhiteSpace(reportPath) || !File.Exists(reportPath))
            {
                return request.CreateErrorResponse(HttpStatusCode.NotFound, $"{mrtFileName} report file was not found.");
            }

            // ரிப்போர்ட் லோடிங் மற்றும் செட்டப்
            report = new StiReport();
            report.Load(reportPath);
            report.Dictionary.Databases.Clear();
            report.Dictionary.Databases.Add(new StiSqlDatabase("Connection", connectionString));
            report.Dictionary.DataSources.Clear();

            // டேட்டா ரெஜிஸ்டர் செய்தல்
            report.RegData(datasourceName, dataSet);
            report.Dictionary.Synchronize();
            report.Compile();
            report.Render();

            // PDF எக்ஸ்போர்ட் மெமரி ஸ்ட்ரீம்
            using (MemoryStream memoryStream = new MemoryStream())
            {
                StiPdfExportService pdfExportService = new StiPdfExportService();
                pdfExportService.ExportPdf(report, memoryStream);

                memoryStream.Position = 0;
                byte[] pdfBytes = memoryStream.ToArray();

                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new ByteArrayContent(pdfBytes);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

                // iframe-ல் தெரிய 'inline' செட்டிங்
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline")
                {
                    FileName = mrtFileName.Replace(".mrt", ".pdf") // எ.கா: EquipmentList.pdf
                };

                // Blazor-க்கான CORS ஹெடர்கள்
                response.Headers.Add("Access-Control-Allow-Origin", "*");
                response.Headers.Add("Access-Control-Allow-Methods", "GET");
                response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Accept");

                return response;
            }
        }
        catch (Exception ex)
        {
            return request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
        }
        finally
        {
            if (report != null)
            {
                report.Dispose();
            }
        }
    }

    public static HttpResponseMessage GeneratePdfReport2Val(HttpRequestMessage request, string sql, string sql2, string datasourceName, string datasourceName2, string mrtFileName)
    {
        StiReport report = null;

        try
        {
            // Web.config-ல் இருந்து கனெக்ஷன் ஸ்டிரிங் எடுத்தல்
            string connectionString = WebConfigurationManager.ConnectionStrings["RBIDB"]?.ConnectionString;

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                return request.CreateErrorResponse(HttpStatusCode.InternalServerError, "RBIDB connection string was not found.");
            }

            // டேட்டா செட் தயார் செய்தல் (Dynamic Datasource Name)
            DataSet dataSet = new DataSet();
            dataSet.DataSetName = datasourceName;

            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connectionString))
            {
                dataSet.Tables.Add(datasourceName);
                adapter.Fill(dataSet, datasourceName);
            }

            DataSet dataSet2 = new DataSet();
            dataSet2.DataSetName = datasourceName2;

            using (SqlDataAdapter adapter2 = new SqlDataAdapter(sql2, connectionString))
            {
                dataSet2.Tables.Add(datasourceName2);
                adapter2.Fill(dataSet2, datasourceName2);
            }

            // MRT ஃபைல் பாத் அமைத்தல் (Dynamic File Name)
            string reportPath = HostingEnvironment.MapPath($"~/Reports/{mrtFileName}");

            if (string.IsNullOrWhiteSpace(reportPath) || !File.Exists(reportPath))
            {
                return request.CreateErrorResponse(HttpStatusCode.NotFound, $"{mrtFileName} report file was not found.");
            }

            // ரிப்போர்ட் லோடிங் மற்றும் செட்டப்
            report = new StiReport();
            report.Load(reportPath);
            report.Dictionary.Databases.Clear();
            report.Dictionary.Databases.Add(new StiSqlDatabase("Connection", connectionString));
            report.Dictionary.DataSources.Clear();

            // டேட்டா ரெஜிஸ்டர் செய்தல்
            report.RegData(datasourceName, dataSet);
            report.RegData(datasourceName2, dataSet2);
            report.Dictionary.Synchronize();
            report.Compile();
            report.Render();

            // PDF எக்ஸ்போர்ட் மெமரி ஸ்ட்ரீம்
            using (MemoryStream memoryStream = new MemoryStream())
            {
                StiPdfExportService pdfExportService = new StiPdfExportService();
                pdfExportService.ExportPdf(report, memoryStream);

                memoryStream.Position = 0;
                byte[] pdfBytes = memoryStream.ToArray();

                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new ByteArrayContent(pdfBytes);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

                // iframe-ல் தெரிய 'inline' செட்டிங்
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline")
                {
                    FileName = mrtFileName.Replace(".mrt", ".pdf") // எ.கா: EquipmentList.pdf
                };

                // Blazor-க்கான CORS ஹெடர்கள்
                response.Headers.Add("Access-Control-Allow-Origin", "*");
                response.Headers.Add("Access-Control-Allow-Methods", "GET");
                response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Accept");

                return response;
            }
        }
        catch (Exception ex)
        {
            return request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
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
