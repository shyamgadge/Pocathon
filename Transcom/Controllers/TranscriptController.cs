using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Transcom.Models;

namespace Transcom.Controllers
{
    public class TranscriptController : Controller
    {

        [HttpPost]
        public async Task<IActionResult> Index(string urlName, string videoId)
        {

            string textTrack = "https://euno-1.api.microsoftstream.com/api/videos/d3813380-988e-4445-809e-9893ad987179/texttracks?api-version=1.4-private";

            //string apiUrl = "https://euno-1.api.microsoftstream.com/api/videos/fa51e2eb-175e-4c72-97aa-d5c30b8594df/texttracks?api-version=1.4-private";

            string vttUrl = "";
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(textTrack);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5NzE2MDQsIm5iZiI6MTYwNjk3MTYwNCwiZXhwIjoxNjA2OTc1NTA0LCJhY3IiOiIxIiwiYWlvIjoiQVNRQTIvOFJBQUFBbFB6R0plN29UVnlQbExnd1M5ZWh2OUFucTRJOUdjWm9tYkhkUUFhbEhHUT0iLCJhbXIiOlsicHdkIl0sImFwcGlkIjoiY2Y1M2ZjZTgtZGVmNi00YWViLThkMzAtYjE1OGU3YjFjZjgzIiwiYXBwaWRhY3IiOiIyIiwiZmFtaWx5X25hbWUiOiJHYWRnZSIsImdpdmVuX25hbWUiOiJTaHlhbSIsImlwYWRkciI6IjEwMy4xMTAuMjU0LjE0MCIsIm5hbWUiOiJHYWRnZSwgU2h5YW0gKENhcGl0YSBTb2Z0d2FyZSkiLCJvaWQiOiIzNDYxMDZmMy03YWE5LTRhNmYtOGNkMi1kZTNlNjNkOWU3OTIiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjM4NTc0OTg3LTI5MzUzODY4MTktMjA5MzY4NjEwLTI1NTc1MDQiLCJwdWlkIjoiMTAwMzIwMDA0OTIyQUJDQyIsInJoIjoiMC5BQUFBZzYzYUh1LXlQVWlCOFN4SWFDOUE3T2o4VThfMjN1dEtqVEN4V09leHo0TUNBT2cuIiwic2NwIjoiYWNjZXNzX21pY3Jvc29mdHN0cmVhbV9zZXJ2aWNlIiwic3ViIjoiVy15ZlNoMHZQZEt4Q2NCR1hMVUlEeG5jUHJTdXBRNnBpSGt0VXc5alAyYyIsInRpZCI6IjFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYyIsInVuaXF1ZV9uYW1lIjoiUDEwNDk3Mjc3QGNhcGl0YS5jby51ayIsInVwbiI6IlAxMDQ5NzI3N0BjYXBpdGEuY28udWsiLCJ1dGkiOiJvVloyanBLQ3FFcWNvNEpwY2lJekFBIiwidmVyIjoiMS4wIn0.LEGT6_xzUxop792Bhw5MKnFqSZTyokNxy0RDd0Yi76FCV6142sv0Id8UveEOkaEK99_dXrswNHNNgL9IfQsD9Ped_TEbiVcGRTTnlOJCclhBcgFKvrhcruyI4zeuckasyqUbB37Bain8-IFuBBhufZpbnEuTHbbSsAlbuzmiOvc-tfytwJmtt7BHz5lOJABAAeBN7IzOAZoNPJfkE59LAchbOg-AkHkTY3T8lwWl0Yvhq50u0clxcuuPamYTIOTufF8uPFmntgPCX09pi2qq3UspSnVpSf1R92JkZ7zs5BAL_IB9gkiKcwRSpO85InSWRe4YIeq_0x3vaFCBU4Zhiw");
                HttpResponseMessage response = await client.GetAsync(textTrack);
                if (response.IsSuccessStatusCode)
                {
                    var data = await response.Content.ReadAsStringAsync();
                    var data1 = JsonConvert.DeserializeObject<dynamic>(data);
                    //var d =  JObject.Parse(data);

                    //dynamic helpDeserialized = JsonConvert.DeserializeObject(data);
                    //dynamic seedDataJson = JsonConvert.DeserializeObject(helpDeserialized.jsonData.Value);
                    vttUrl = "https://euno-1-content.api.microsoftstream.com/api/texttracks/1/5cf25b09-b8b6-48b5-9c34-781aebc9adcd/MWNmMTk3YTAtZTI4Yi00NGMzLWFmYzEtNzgyNGU4NWEyZjcyLnYwLnZ0dA?validTill=2020-12-04T00%3a00%3a00.0000000Z&aadUserOId=346106f3-7aa9-4a6f-8cd2-de3e63d9e792&encoding=base64&api-version=1.4-private&signature=V%2fTO0HPmmKHcqLToLwZyFWlGEVMWXFGSoTNvlMoq6nE%3d";
                }
            }

            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(vttUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5NzE2MDQsIm5iZiI6MTYwNjk3MTYwNCwiZXhwIjoxNjA2OTc1NTA0LCJhY3IiOiIxIiwiYWlvIjoiQVNRQTIvOFJBQUFBbFB6R0plN29UVnlQbExnd1M5ZWh2OUFucTRJOUdjWm9tYkhkUUFhbEhHUT0iLCJhbXIiOlsicHdkIl0sImFwcGlkIjoiY2Y1M2ZjZTgtZGVmNi00YWViLThkMzAtYjE1OGU3YjFjZjgzIiwiYXBwaWRhY3IiOiIyIiwiZmFtaWx5X25hbWUiOiJHYWRnZSIsImdpdmVuX25hbWUiOiJTaHlhbSIsImlwYWRkciI6IjEwMy4xMTAuMjU0LjE0MCIsIm5hbWUiOiJHYWRnZSwgU2h5YW0gKENhcGl0YSBTb2Z0d2FyZSkiLCJvaWQiOiIzNDYxMDZmMy03YWE5LTRhNmYtOGNkMi1kZTNlNjNkOWU3OTIiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjM4NTc0OTg3LTI5MzUzODY4MTktMjA5MzY4NjEwLTI1NTc1MDQiLCJwdWlkIjoiMTAwMzIwMDA0OTIyQUJDQyIsInJoIjoiMC5BQUFBZzYzYUh1LXlQVWlCOFN4SWFDOUE3T2o4VThfMjN1dEtqVEN4V09leHo0TUNBT2cuIiwic2NwIjoiYWNjZXNzX21pY3Jvc29mdHN0cmVhbV9zZXJ2aWNlIiwic3ViIjoiVy15ZlNoMHZQZEt4Q2NCR1hMVUlEeG5jUHJTdXBRNnBpSGt0VXc5alAyYyIsInRpZCI6IjFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYyIsInVuaXF1ZV9uYW1lIjoiUDEwNDk3Mjc3QGNhcGl0YS5jby51ayIsInVwbiI6IlAxMDQ5NzI3N0BjYXBpdGEuY28udWsiLCJ1dGkiOiJvVloyanBLQ3FFcWNvNEpwY2lJekFBIiwidmVyIjoiMS4wIn0.LEGT6_xzUxop792Bhw5MKnFqSZTyokNxy0RDd0Yi76FCV6142sv0Id8UveEOkaEK99_dXrswNHNNgL9IfQsD9Ped_TEbiVcGRTTnlOJCclhBcgFKvrhcruyI4zeuckasyqUbB37Bain8-IFuBBhufZpbnEuTHbbSsAlbuzmiOvc-tfytwJmtt7BHz5lOJABAAeBN7IzOAZoNPJfkE59LAchbOg-AkHkTY3T8lwWl0Yvhq50u0clxcuuPamYTIOTufF8uPFmntgPCX09pi2qq3UspSnVpSf1R92JkZ7zs5BAL_IB9gkiKcwRSpO85InSWRe4YIeq_0x3vaFCBU4Zhiw");
                HttpResponseMessage response = await client.GetAsync(vttUrl);
                if (response.IsSuccessStatusCode)
                {
                    // Get the vtt file content
                }
            }

            //Create Word file
            string filepath = @"C:\Users\P10494765.AD\Desktop\testing3.docx";
            CreateWordprocessingDocument(filepath);

            //using (var client = new WebClient())
            //{

            //    client.Headers.Clear();
            //    //client.Headers.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            //    //wb.Headers.Add(HttpRequestHeader.Cookie, "somecookie");
            //    client.Headers.Add(HttpRequestHeader.Cookie, "Authorization_Api=eyJUb2tlbiI6InlaT0ZqMnBVbXd3U3VZQnFWenRqdUJ3TGY2dm04WlJVNVdMazVGck1LWWhNdjNCTTZRc0dCVWFJYmNybStwWjRWQnFEWURRYytEc3VwenVNRTNqRFM0YmgzSEdpc2ZmUmFOU1pOMnIwY1A2ZGpSTHJTVjZTTHAwbmtQREQyQ0p6NlBUL2xHUVEvU2h4SlBuNkxRaW5PK1dQSmllSk5vYkFQN2xiSDlsYitIR2Y2WjFnTDR3dFJDMnVHR0RVNDR4Zm0rU013c2FvdlVrUUMyWDdDSzhwd3B3WWNwbFZlZGRsb1pTYVgrRTU1U1RpdFoyMXh6eGhyNGREYmQ5bEMrMHl4UFVhNm5HNHpuUmJBS1lyUEtvaXR6UEpOOEZlYmRHNVkrN3Y4UGxTaUs3L0xuWVZNRVg5ZEx6OHNqL0JNZThJdDNNTGtEY2c5cGFDOEZZZGdXd3RGSXZQaGF3S0wxdHEzc0pmeGsrajNlTVBrRU1zcDZvVjJpRzhFYWx4ekZrOGQxOHFnZ1FKZVBhTFl2bGhncVkrbXNXeXBaNGdJZUs2UklDUThFV3RJOVptMU91ZWo4bjNjR1hWZytscUJLazFRbHZjNmpuUjhGc3RidkxQQ25SRnVGM3E5dWZtekhpOG1vck5MQi9kbFMxSGNncXdxQktOc2xSVE0vd0xHTk93Ulc5UllpMXZnV1ovWkdWVlltamZRZUJWdVFrRGJtUXVwTkNhSEVlc0llaWtqWXhtSlcwRmprMUErZGllR2UrVVowRGpRaGFVMVRMUVkyQVRXcUZjSjYzMDJvTk83RkNGcFhLU2xrazRBYVJsbUtpT0JaVEJhbkcybVdPdkduN2thNjkxWnFWdzkwRVY0MjdjSFdMMzBOajBWNkZuZVI4T2NsdGRyaUlpb25jSkp3S3R1ODJTeFdqeCs2UXhGYTZrd29HL2dVUWtPV2c5OHp3RVd5Um5UTzlzZm55YXFycUQ5b2hXT2JQQW5mZzF6amo0WVZUZy9rU1NXVzJJSHowQlI5WFVQVENWL1BVeEFBS2dha2hEN1dqdmo5TW1CcWxSelozbUh3V1N6ZTROQUpld2xZZ3M4MHcyM2pqclBNTlV0SGxqQityQnVHVW5WWGpGRndoTW5RdHRnMlVEelNoT3ZlcHBmSEZTZnFubGtXY2E0Y3VRM1lGczgvWHFpUFNxb2xQRVZPeXJoeVRHY2haaUJLcXlSeElLSjFzMDVtSmE5MmVVeGdEL2VVWEpDaHIxNWxNL2hzQzM1MUpyY2czU2dpV2FKTHZqMlpNL3J6VVlGWEd1L0tvcm9ZNllMclNCdWFCUXlDVGVYOTNHOHVocWdiRDJIUkQxQUVKOUgrcncxd3BrdnZNZ01JTWkybVFPajRxK2J5cThNbU5nVzBDdEM3NmVaZGp0Y1p2RVJISDRoc2ZBajFoNVNFKzdYQ1Q5Um5xZ0tMR3Z5OGdLeTlDYTNaSVU2OVhBaXVEVmRoK21Cbmc4YkFieU9PWGVRTFlBVVZSYnZLaGo1MGdZeUx1VkZIR1d4OCtxT25mMG5tL24wYWc0WnZRSEVhOXpNcTkxT2N5Qnc0M3o2eEdGeHVhaitsanlzeUZoU0FOOU82SERzYU1uQ3ZMLzZNd0ZKOEJkUVRPSzlIL3N5UlVqMkt3OFk2VXRIRkhtTFowTk0vUXJ6T1dva1dDbjYrNVVQcWtQcFZ4WnpWTXZOUlFPZmJNaFRsNmFLWUhXaHNYYjdFYTZGUXpWV0NDbWtsdk5NV2xLK0ZxOEVwWHJPa3R2OUxwNEQydmgzNEI2TTlKaWRMU1hNNktKSFFBck5TNElEVjdFdytEZFdYbGV2aVNjZFB0bmFhM0E4ZFRiU25DSFVQMWFzUzIvUHcvanpZdm1TbXR4a1BtSVphREJ4dHFJaUlsaHpOc2FDbUZJNXl1bXdib01ZNDhpSVB4ZjNlSCs2UXQyUTlCbVpyYW04YXVqNEk3TUd0ZmZYVEpnUmpZS1NYZTNOenV6T0ZFUjVoYVFsR3FmVng0TFpDRG9XVWhkOGlsbXNueksvbXVWcXJ0cEUvczFhYU9tekprL0hhNS9lazVSMG8yYS9Mb0dpSVJ3azdOMlZZa2FMeEMrRFF5aCtmRWtWL1gwcjZRTGxsVzFMSkRWMUt1MVd5TngzaVk1WjlrSkdxUjJQZVhucUR6SWE5SVlLb1UrRGRoVm9IVG41WVdUWTE3TktkSTA3K1loejBVNzl5WEI3c2FBdWNyWHM3b1F2cjliNzh4ZXBUTENpQUxOY3VtcHBmTlZVN1J1Q20rTWRqSEhvQlJreFFLYjJzNUhNMkI0Mm9KSVFhYnRHSUtuOGtaU1F4K3kvUHk3YzRQVkg2Tm5LanBUUERVRnU1bnFab3VOVVVha0w1dWxGcmluVmdZRGRMTXNwZnJEeExQYlFHRzEvSWJOV3RWUTBIejVFbFk5WEtTL0V0YjZvOVdlWjJWQThOd0RwdEFXSWRYbXdLUVkvbkM4bFFUMlUwVXRhb3ZueDY0ZFFEQVBlOTlCeWlVOExMT0hlaVVuZzZyM1BwK1RLeDFzZ29weHFuYndVRkNLU01wZnBoNXB5MnAxcUQwSlFhandwazJMcFNlWXV2anlta2lsUzFwSmcrTi9Wc2VJR0xEeHRVUFNNUjQ0T29Cay9pdlFHc1pmODA3UFVKU2ZuaFd5MFlLQm45MldKS1FqU2k0Z2pJSGpQcm1NZGFSOTBXV25Td0Rqb0ZkanUyM0RieUdFSVN3bTV5eU5CM0ExZmRPdTQrUmRMNzZlWGRnZjRZdnJiRlhYajZiQzArR0wrc0FRUlAvckZsbk1rWk1zYkZvZ3VvQ0RMQWVyemJCRCtyYUEzK1RyenF0MTBwMzc2UGhUT1RHYS9rNGtNZE41OVlxdE40TWNwdm44VmpYc21nPT0iLCJUb2tlblNpZ25hdHVyZSI6IkVzT3lsSTduSUFnUXlJK080NU5SVGY1aXRnU3pLSkRVaUxtVWhJRGFpNEE9IiwiRW5jcnlwdGlvbkl2IjoiUEsxWTNCbElZaW5PM2hUR3NVRnNEQT09IiwiRW5jcnlwdGlvbktleVNpZ25hdHVyZSI6InlnZWZleUpiSFlRbEsycmVtNnprVzBzbEtQVmFpT1hmKy9FOTU2cEFzOGM9In0%3d; Signature_Api=U%252bKy%252baA2zjJfx6wa9XkiiaWLn20wrxc2YL5XjhFkXLM%253d; UserSession_Api=signature=puBRc30o6_eZCcfaaZpUw8Y50YRotvidSc26tW32Cdc&payload=eyJFbmNyeXB0ZWRQYXlsb2FkIjoiMUQ5Zy9RWXA4UFBFUlY1T29XSVRRSXo4eENPY3JRb1lvcHZkQ1NhYmtzaWhSSlhrTkZYYm9lbU5DRWtoWkJ1Z2NTYzFTKzhjWm9idTNWNk5EemRGdTdPaXoreEJZT2l0QmVGa1lMMHhZNk0vR1kvQUFmN1hjOWp2ekY0S0xDTXhDQVdFUjVabWhsWS90OXJNalc4ZVQxd05JU1hGZ25DUDV0V3RxRHBmUEdoREdIaGd4V0d4SFhzUTFGc0JQUFYzRVhRYThvWEtpLzUxRnFGNnFsZlMyZ0RqQWw2WE03TENBNEZYeTdMTnIwanVuaktYQllscFkrdHdVVUFIM0ZjOHhmQ25hVHVyK3MybW9Id2JCRENRY29obHhnajcrQWhGOFREa0h4NHNIZE50V2NXSnhsdmxnMlJIcms0S0l4OEtzNVZ3Q0c4bUwzaFBlTDhXbVhkazhnPT0iLCJFbmNyeXB0aW9uS2V5SGFzaCI6InlnZWZleUpiSFlRbEsycmVtNnprVzBzbEtQVmFpT1hmKy9FOTU2cEFzOGM9IiwiRW5jcnlwdGlvbktleUhhc2hDb2RlIjoxMjExNzY4MTMzLCJFbmNyeXB0aW9uSXYiOiJGQ2ZnOFAxSU9OTWZLSEtUUElHY29nPT0ifQ");
            //    var content = client.DownloadData(textTrack);
            //    using (var stream = new MemoryStream(content))
            //    {
            //        FileStream file = new FileStream("C:\\file.txt", FileMode.Create, FileAccess.Write);
            //        stream.WriteTo(file);
            //        file.Close();
            //    }
            //}

            //var parser = new SubtitlesParser.Classes.Parsers.SubParser();

            //var allFiles = BrowseTestSubtitlesFiles();
            //foreach (var file in allFiles)
            //{
            //    var fileName = Path.GetFileName(file);
            //    using (var fileStream = new FileStream(@file, FileMode.Open, FileAccess.Read))
            //    {
            //        try
            //        {
            //            var mostLikelyFormat = parser.GetMostLikelyFormat(fileName);
            //            var items = parser.ParseStream(fileStream, Encoding.UTF8, mostLikelyFormat);
            //            if (items.Any())
            //            {
            //                Console.WriteLine("Parsing of file {0}: SUCCESS ({1} items - {2}% corrupted)",
            //                    fileName, items.Count, (items.Count(it => it.StartTime <= 0 || it.EndTime <= 0) * 100) / items.Count);
            //                /*foreach (var item in items)
            //                {
            //                    Console.WriteLine(item);
            //                }*/
            //                /*var duplicates =
            //                    items.GroupBy(it => new {it.StartTime, it.EndTime}).Where(grp => grp.Count() > 1);
            //                Console.WriteLine("{0} duplicate items", duplicates.Count());*/

            //            }
            //            else
            //            {
            //                //Console.WriteLine("Parsing of file {0}: SUCCESS (No items found!)", fileName, items.Count);
            //            }

            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //    Console.WriteLine("----------------------");
            //}

            //Console.ReadLine();


            return View();
        }


        public static void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                AddSettingsToMainDocumentPart(mainPart);

                // Create the document structure and add some text.
                Document doc = new Document();

                HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>("rId7");
                GenerateHeader(headerPart);

                Body body = new Body();
                // 1 paragrafo
                Paragraph para = new Paragraph();

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Bold" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };
                FontSize fontSize1 = new FontSize() { Val = "50" };
                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(fontSize1);
                paragraphProperties1.Append(justification1);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run = new Run();
                RunProperties runProperties1 = new RunProperties();

                Text text = new Text() { Text = "The OpenXML SDK rocks!" };

                // siga a ordem 
                run.Append(runProperties1);
                run.Append(text);
                para.Append(paragraphProperties1);
                para.Append(run);
                body.Append(para);

                // 2 paragrafo
                Paragraph para2 = new Paragraph();

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Bold" };
                Justification justification2 = new Justification() { Val = JustificationValues.Right };
                FontSize fontSize = new FontSize() { Val = "50" };
                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
               
                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);
                paragraphProperties2.Append(fontSize);

                Run run2 = new Run();
                RunProperties runProperties3 = new RunProperties();
                Text text2 = new Text();
                text2.Text = "Teste aqui";

                run2.AppendChild(new Break());
                run2.AppendChild(runProperties3);
                run2.AppendChild(text2);
                run2.AppendChild(new Break());

                para2.Append(paragraphProperties2);
                para2.Append(run2);

                // todos os 2 paragrafos no main body

                body.Append(para2);


                doc.Append(body);

                wordDocument.MainDocumentPart.Document = doc;

                wordDocument.Close();
            }
        }

        private static void GenerateHeader(HeaderPart part)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00B22010", RsidRunAdditionDefault = "00B22010" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = "This is a ";

            run1.Append(text1);

            Run run2 = new Run();
            Text text2 = new Text();
            text2.Text = "header";

            run2.Append(text2);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(bookmarkStart1);
            paragraph1.Append(bookmarkEnd1);

            header1.Append(paragraph1);
            part.Header = header1;
        }

        private static void AddSettingsToMainDocumentPart(MainDocumentPart part)
        {
            DocumentSettingsPart settingsPart = part.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
               new Compatibility(
                   new CompatibilitySetting()
                   {
                       Name = new EnumValue<CompatSettingNameValues>(CompatSettingNameValues.CompatibilityMode),
                       Val = new StringValue("16"),
                       Uri = new StringValue("http://schemas.microsoft.com/office/word")
                   }
               )
            );
            settingsPart.Settings.Save();
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
