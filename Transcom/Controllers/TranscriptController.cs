using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Transcom.Models;
using SubtitlesParser;
using Newtonsoft.Json.Linq;

namespace Transcom.Controllers
{
    public class TranscriptController : Controller
    {

        

        [HttpPost]
        public async Task<IActionResult> Index(string urlName)
        {
            string title = "Transcom";
            List<string> sentences = new List<string>();
            string[] s = urlName.Split('/');
            string videoId = s.Last();
            string token = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDcwMjA3MDQsIm5iZiI6MTYwNzAyMDcwNCwiZXhwIjoxNjA3MDI0NjA0LCJhY3IiOiIxIiwiYWlvIjoiQVNRQTIvOFJBQUFBZUZqYmtEWFpYNlNYUGE4K2tCckxVUFhJa3Z4bWEyME12d3krTHdiWFF4dz0iLCJhbXIiOlsicHdkIl0sImFwcGlkIjoiY2Y1M2ZjZTgtZGVmNi00YWViLThkMzAtYjE1OGU3YjFjZjgzIiwiYXBwaWRhY3IiOiIyIiwiZmFtaWx5X25hbWUiOiJHYWRnZSIsImdpdmVuX25hbWUiOiJTaHlhbSIsImlwYWRkciI6IjEwMy4xMTAuMjU0LjE0MCIsIm5hbWUiOiJHYWRnZSwgU2h5YW0gKENhcGl0YSBTb2Z0d2FyZSkiLCJvaWQiOiIzNDYxMDZmMy03YWE5LTRhNmYtOGNkMi1kZTNlNjNkOWU3OTIiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjM4NTc0OTg3LTI5MzUzODY4MTktMjA5MzY4NjEwLTI1NTc1MDQiLCJwdWlkIjoiMTAwMzIwMDA0OTIyQUJDQyIsInJoIjoiMC5BQUFBZzYzYUh1LXlQVWlCOFN4SWFDOUE3T2o4VThfMjN1dEtqVEN4V09leHo0TUNBT2cuIiwic2NwIjoiYWNjZXNzX21pY3Jvc29mdHN0cmVhbV9zZXJ2aWNlIiwic3ViIjoiVy15ZlNoMHZQZEt4Q2NCR1hMVUlEeG5jUHJTdXBRNnBpSGt0VXc5alAyYyIsInRpZCI6IjFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYyIsInVuaXF1ZV9uYW1lIjoiUDEwNDk3Mjc3QGNhcGl0YS5jby51ayIsInVwbiI6IlAxMDQ5NzI3N0BjYXBpdGEuY28udWsiLCJ1dGkiOiJtSktFOHNadmxrLWV4dzkyMXhzN0FBIiwidmVyIjoiMS4wIn0.KnEP7oB7aZXDoJEKHYoX3vi4O4jcw0u__FXof6Us5jbXkghLSlYlhHT86VIPEUdY6qTijmoCn6JaD9I8vfwO_FmX4f7PxZl7R2j5XvGfEZ-6RsGax6uJ0HfaTGCbp1dQD2xovY6o2Kn7yuP57leNqY2XFzhs5k_ZyHrGdHaFyDUTSpDJQhtInH9VpCfds5y3TxAubJSJTUE4VJnItlQHrIJyaBO7Zrmo42-SH9wHCnoasHTHwKHBbhM5g_UpTREhaUK6-FSv9Hp2jZEfc3yZAzo7v-wUSjukTT7R1EsUSz3hplLu5WKt1iheGRPywW6pzji2BF8xlq_nYly5XWV0QA";

            string textTrackUrl = "https://euno-1.api.microsoftstream.com/api/videos/" + videoId + "/texttracks?api-version=1.4-private";
            string titleUrl = "https://euno-1.api.microsoftstream.com/api/videos/" + videoId + "?$expand=creator,tokens,status,liveEvent,extensions&api-version=1.4-private";

            string vttUrl = "";
         
            HttpResponseMessage urlResponse = await GetHttpResponse(textTrackUrl, token);
            if (urlResponse.IsSuccessStatusCode)
            {
                var textTracksResponseString = await urlResponse.Content.ReadAsStringAsync();
                TextTracksResponse textTracksResponse = JsonConvert.DeserializeObject<TextTracksResponse>(textTracksResponseString);
                vttUrl = textTracksResponse?.value?[0].url;
            }

           
            HttpResponseMessage titleResponse = await GetHttpResponse(titleUrl, token);
            if (titleResponse.IsSuccessStatusCode)
            {
                var textTracksResponseString = await titleResponse.Content.ReadAsStringAsync();
                var responseJObject = JObject.Parse(textTracksResponseString);
                title = (string)responseJObject["name"];
            }

            sentences = ExtractVttContent( token, vttUrl);

            //Create Word file
            string fileName = title + ".docx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            CreateWordprocessingDocument(path, title, sentences);
            TempData["Path"] = "";
            TempData["Path"] = path;
            return View();
        }

        private async Task<HttpResponseMessage> GetHttpResponse(string url, string token)
        {
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", token);
                HttpResponseMessage response = await client.GetAsync(url);

                return response;
            }
        }

        private static List<string> ExtractVttContent( string token, string vttUrl)
        {
            List<string> sentences = new List<string>();
            using (var client = new WebClient())
            {

                client.Headers.Clear();

                client.Headers.Add("authorization", token);
                var content = client.DownloadData(vttUrl);
                using (var stream = new MemoryStream(content))
                {
                    FileStream file = new FileStream("C:\\file.txt", FileMode.Create, FileAccess.Write);
                    stream.WriteTo(file);
                    file.Close();

                    SubtitlesParser.Classes.Parsers.SubParser parser = new SubtitlesParser.Classes.Parsers.SubParser();

                    using (var fileStream = new FileStream("C:\\file.txt", FileMode.Open, FileAccess.Read))
                    {
                        try
                        {
                            var mostLikelyFormat = parser.GetMostLikelyFormat(fileStream.Name);
                            var items = parser.ParseStream(fileStream, Encoding.UTF8, mostLikelyFormat);

                            string sentence = string.Empty;
                            foreach (var item in items)
                            {
                                foreach (var line in item.Lines)
                                {
                                    foreach (var word in line.ToCharArray())
                                    {
                                        sentence = sentence + word;

                                        if (word.Equals('.') || word.Equals('?'))
                                        {
                                            sentences.Add(sentence);
                                            sentence = string.Empty;
                                        }
                                    }
                                    sentence = sentence + " ";
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            return sentences;
        }

        [HttpGet]
        public async Task<IActionResult> Download(string path)
        {
            var memory = new MemoryStream();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return File(memory, GetMimeTypes()[ext], Path.GetFileName(path));
        }

        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
        {
            {".txt", "text/plain"},
            {".pdf", "application/pdf"},
            {".doc", "application/vnd.ms-word"},
            {".docx", "application/vnd.ms-word"},
            {".png", "image/png"},
            {".jpg", "image/jpeg"}
        };
        }

        public void CreateWordprocessingDocument(string filepath, string title, List<string> conversation)
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
                //paragraphProperties1.Append(fontSize1);
                paragraphProperties1.Append(justification1);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run = new Run();
                RunProperties runProperties1 = new RunProperties();
                FontSize fontSize2 = new FontSize() { Val = "36" };

                runProperties1.Append(fontSize2);
                Text text = new Text() { Text = title };

                // siga a ordem 
                run.Append(runProperties1);
                run.Append(text);
                para.Append(paragraphProperties1);
                para.Append(run);
                body.Append(para);
                //Paragraph titlePara = CreateParagraph( title, JustificationValues.Center);
                //body.Append(titlePara);

                // "Here are the discussed points:" Paragraph
                Paragraph pointsHeaderPara = CreateParagraph("Here are the discussed points:", JustificationValues.Left);

                body.Append(pointsHeaderPara);

                foreach (var line in conversation)
                {
                    Paragraph paragraph1 = CreateParagraphForBullets(line);
                    body.Append(paragraph1);
                }

                doc.Append(body);

                wordDocument.MainDocumentPart.Document = doc;

                wordDocument.Close();
            }
        }

        private Paragraph CreateParagraph(string text, JustificationValues justification)
        {
            Paragraph paragraph = new Paragraph(
                        new ParagraphProperties(
                          new ParagraphStyleId() { Val = "Bold" },
                          new Justification() { Val = justification },
                          new ParagraphMarkRunProperties(),
                          new FontSize() { Val = "12" }
                          ),
                        new Run(
                          new RunProperties(),
                          new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

            return paragraph;
        }

        private Paragraph CreateParagraph(string text)
        {
            Paragraph paragraph = new Paragraph(
            new ParagraphProperties(
              new NumberingProperties(
                new NumberingLevelReference() { Val = 0 },
                new NumberingId() { Val = 2 })),
            new Run(
              new RunProperties(),
              new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            return paragraph;
        }
        private Paragraph CreateParagraphForBullets(string text)
        {
            Paragraph paragraph = new Paragraph(
            new ParagraphProperties(
              new NumberingProperties(
                new NumberingLevelReference() { Val = 0 },
                new NumberingId() { Val = 2 })),
            new Run(
              new RunProperties(),
              new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            return paragraph;
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
