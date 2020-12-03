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

namespace Transcom.Controllers
{
    public class TranscriptController : Controller
    {
        [HttpPost]
        public async Task<IActionResult> Index(string urlName)
        {
            string[] s = urlName.Split('/');
            string videoId = s.Last();

            string textTrack = "https://euno-1.api.microsoftstream.com/api/videos/" + videoId + "/texttracks?api-version=1.4-private";

            string vttUrl = "";
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(textTrack);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDcwMDgxNDYsIm5iZiI6MTYwNzAwODE0NiwiZXhwIjoxNjA3MDEyMDQ2LCJhY3IiOiIxIiwiYWlvIjoiRTJSZ1lHQU1YYVBvSk11Y3pEcjFmMGFJd3dKYmNXdDNIZVdnKzk4TUhOZFBmZFYvTkJrQSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiJjZjUzZmNlOC1kZWY2LTRhZWItOGQzMC1iMTU4ZTdiMWNmODMiLCJhcHBpZGFjciI6IjIiLCJmYW1pbHlfbmFtZSI6IkRoYW5kZSIsImdpdmVuX25hbWUiOiJEaXBhayIsImlwYWRkciI6IjgyLjIwMy4zMy4xMzQiLCJuYW1lIjoiRGhhbmRlLCBEaXBhayAoQ2FwaXRhIFNvZnR3YXJlKSIsIm9pZCI6ImJhN2RiNzVhLTJlZjUtNDI1Zi04ZWM2LWY5ODZhMWQzMjYyMiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMzg1NzQ5ODctMjkzNTM4NjgxOS0yMDkzNjg2MTAtMjQ3Nzg2MCIsInB1aWQiOiIxMDAzM0ZGRkFGRENFMjJFIiwicmgiOiIwLkFBQUFnNjNhSHUteVBVaUI4U3hJYUM5QTdPajhVOF8yM3V0S2pUQ3hXT2V4ejRNQ0FJOC4iLCJzY3AiOiJhY2Nlc3NfbWljcm9zb2Z0c3RyZWFtX3NlcnZpY2UiLCJzdWIiOiJOVm43Z0RZVk5hTjVRbjd0TWhLVWdJbVVwOFBTWVU4UEZYeGxJVEdielFnIiwidGlkIjoiMWVkYWFkODMtYjJlZi00ODNkLTgxZjEtMmM0ODY4MmY0MGVjIiwidW5pcXVlX25hbWUiOiJQMTA0NzkxNTZAY2FwaXRhLmNvLnVrIiwidXBuIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInV0aSI6ImhTVEVYdUVYZ0VhSk1xVGtKaVJOQUEiLCJ2ZXIiOiIxLjAifQ.O2gET2TM84G7yfXNJPFA_EG-Y9gj5xKrPXMRODRdmfPe6MtOIs3zATI_YXmi_jGAP5aDERG7sIrva7x058SlhJaL5V7LLJk-QP6Bniug6xYhSg4MglmenbRiSjZDy51fuMxOQkiKKmCWLoQGfsa8NdbKvNHKrR8FfOTXy_6H3JbxbhNYOSwgLNSIthMewesqLrEcR8itrRysR18ueldLhhZpA1iioPnA-WVqubeTXBfC7lMhbV9rw7cLo4jQ4Hgd-vJCzqpVuB9znkBGeEGqjZ1L2L8Ig2WxGkj7Jn8rjGj832MZZv1cfQld4MaF8vVm6H7YqXhfMpzYGOjBkuGiIg");
                HttpResponseMessage response = await client.GetAsync(textTrack);
                if (response.IsSuccessStatusCode)
                {
                    var textTracksResponseString = await response.Content.ReadAsStringAsync();
                    TextTracksResponse textTracksResponse = JsonConvert.DeserializeObject<TextTracksResponse>(textTracksResponseString);
                    vttUrl = textTracksResponse?.value?[0].url;
                }
            }

            using (var client = new WebClient())
            {

                client.Headers.Clear();

                client.Headers.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDcwMDgxNDYsIm5iZiI6MTYwNzAwODE0NiwiZXhwIjoxNjA3MDEyMDQ2LCJhY3IiOiIxIiwiYWlvIjoiRTJSZ1lHQU1YYVBvSk11Y3pEcjFmMGFJd3dKYmNXdDNIZVdnKzk4TUhOZFBmZFYvTkJrQSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiJjZjUzZmNlOC1kZWY2LTRhZWItOGQzMC1iMTU4ZTdiMWNmODMiLCJhcHBpZGFjciI6IjIiLCJmYW1pbHlfbmFtZSI6IkRoYW5kZSIsImdpdmVuX25hbWUiOiJEaXBhayIsImlwYWRkciI6IjgyLjIwMy4zMy4xMzQiLCJuYW1lIjoiRGhhbmRlLCBEaXBhayAoQ2FwaXRhIFNvZnR3YXJlKSIsIm9pZCI6ImJhN2RiNzVhLTJlZjUtNDI1Zi04ZWM2LWY5ODZhMWQzMjYyMiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMzg1NzQ5ODctMjkzNTM4NjgxOS0yMDkzNjg2MTAtMjQ3Nzg2MCIsInB1aWQiOiIxMDAzM0ZGRkFGRENFMjJFIiwicmgiOiIwLkFBQUFnNjNhSHUteVBVaUI4U3hJYUM5QTdPajhVOF8yM3V0S2pUQ3hXT2V4ejRNQ0FJOC4iLCJzY3AiOiJhY2Nlc3NfbWljcm9zb2Z0c3RyZWFtX3NlcnZpY2UiLCJzdWIiOiJOVm43Z0RZVk5hTjVRbjd0TWhLVWdJbVVwOFBTWVU4UEZYeGxJVEdielFnIiwidGlkIjoiMWVkYWFkODMtYjJlZi00ODNkLTgxZjEtMmM0ODY4MmY0MGVjIiwidW5pcXVlX25hbWUiOiJQMTA0NzkxNTZAY2FwaXRhLmNvLnVrIiwidXBuIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInV0aSI6ImhTVEVYdUVYZ0VhSk1xVGtKaVJOQUEiLCJ2ZXIiOiIxLjAifQ.O2gET2TM84G7yfXNJPFA_EG-Y9gj5xKrPXMRODRdmfPe6MtOIs3zATI_YXmi_jGAP5aDERG7sIrva7x058SlhJaL5V7LLJk-QP6Bniug6xYhSg4MglmenbRiSjZDy51fuMxOQkiKKmCWLoQGfsa8NdbKvNHKrR8FfOTXy_6H3JbxbhNYOSwgLNSIthMewesqLrEcR8itrRysR18ueldLhhZpA1iioPnA-WVqubeTXBfC7lMhbV9rw7cLo4jQ4Hgd-vJCzqpVuB9znkBGeEGqjZ1L2L8Ig2WxGkj7Jn8rjGj832MZZv1cfQld4MaF8vVm6H7YqXhfMpzYGOjBkuGiIg");
                var content = client.DownloadData(vttUrl);
                using (var stream = new MemoryStream(content))
                {
                    FileStream file = new FileStream("C:\\file.txt", FileMode.Create, FileAccess.Write);
                    stream.WriteTo(file);
                    file.Close();

                    var parser = new SubtitlesParser.Classes.Parsers.SubParser();

                    using (var fileStream = new FileStream("C:\\file.txt", FileMode.Open, FileAccess.Read))
                    {
                        try
                        {
                            var mostLikelyFormat = parser.GetMostLikelyFormat(fileStream.Name);
                            var items = parser.ParseStream(fileStream, Encoding.UTF8, mostLikelyFormat);
                            List<string> sentences = new List<string>();
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
            //Create Word file
            string fileName = System.IO.Path.GetRandomFileName() + ".docx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            CreateWordprocessingDocument(path);
            TempData["Path"] = "";
            TempData["Path"] = path;
            return View();
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
                //paragraphProperties1.Append(fontSize1);
                paragraphProperties1.Append(justification1);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run = new Run();
                RunProperties runProperties1 = new RunProperties();
                FontSize fontSize2 = new FontSize() { Val = "36" };

                runProperties1.Append(fontSize2);
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
                NumberingFormat format = new NumberingFormat() { Val = NumberFormatValues.Bullet };

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);
                paragraphProperties2.Append(fontSize);
                paragraphProperties2.Append(format);

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
                body.Append(para2);
                // todos os 2 paragrafos no main body


                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00EA7FFB", RsidParagraphProperties = "00EA7FFB", RsidRunAdditionDefault = "00EA7FFB" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "ListParagraph" };

                NumberingProperties numberingProperties3 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference3 = new NumberingLevelReference() { Val = 0 };
                NumberingId numberingId1 = new NumberingId() { Val = 4 };

                numberingProperties3.Append(numberingLevelReference3);
                numberingProperties3.Append(numberingId1);

                paragraphProperties3.Append(paragraphStyleId3);
                paragraphProperties3.Append(numberingProperties3);

                Run run1 = new Run();
                Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text1.Text = "Item New Name ";

                run1.AppendChild(text1);
                run1.AppendChild(new Break());
                run1.AppendChild(new Text("Item New Name "));
                run1.AppendChild(new Break());
                run1.AppendChild(new Text("Item New Name 3"));
                run1.AppendChild(new Break());
                run1.AppendChild(new Text("Item New Name 4"));
                run1.AppendChild(new Break());
                run1.AppendChild(new Text("Item New Name 5"));
                run1.AppendChild(new Break());

                paragraph1.Append(paragraphProperties3);
                paragraph1.Append(run1);
                body.Append(paragraph1);




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
