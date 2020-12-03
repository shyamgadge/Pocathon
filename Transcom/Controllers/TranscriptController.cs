using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Transcom.Models;
using SubtitlesParser;

namespace Transcom.Controllers
{
    public class TranscriptController : Controller
    {
        [HttpPost]
        public async Task<IActionResult> Index(string urlName, string videoId)
        {
            string textTrack = "https://euno-1.api.microsoftstream.com/api/videos/d3813380-988e-4445-809e-9893ad987179/texttracks?api-version=1.4-private";

            string vttUrl = "";
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(textTrack);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5OTk5MTYsIm5iZiI6MTYwNjk5OTkxNiwiZXhwIjoxNjA3MDAzODE2LCJhY3IiOiIxIiwiYWlvIjoiRTJSZ1lJZytkV2RCbjlGMS8yVnhEOFJZRFhZNzhiQmVpYk02SkoxNk5tUi8zY2NUT2hrQSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiJjZjUzZmNlOC1kZWY2LTRhZWItOGQzMC1iMTU4ZTdiMWNmODMiLCJhcHBpZGFjciI6IjIiLCJmYW1pbHlfbmFtZSI6IkRoYW5kZSIsImdpdmVuX25hbWUiOiJEaXBhayIsImlwYWRkciI6IjgyLjIwMy4zMy4xMzQiLCJuYW1lIjoiRGhhbmRlLCBEaXBhayAoQ2FwaXRhIFNvZnR3YXJlKSIsIm9pZCI6ImJhN2RiNzVhLTJlZjUtNDI1Zi04ZWM2LWY5ODZhMWQzMjYyMiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMzg1NzQ5ODctMjkzNTM4NjgxOS0yMDkzNjg2MTAtMjQ3Nzg2MCIsInB1aWQiOiIxMDAzM0ZGRkFGRENFMjJFIiwicmgiOiIwLkFBQUFnNjNhSHUteVBVaUI4U3hJYUM5QTdPajhVOF8yM3V0S2pUQ3hXT2V4ejRNQ0FJOC4iLCJzY3AiOiJhY2Nlc3NfbWljcm9zb2Z0c3RyZWFtX3NlcnZpY2UiLCJzdWIiOiJOVm43Z0RZVk5hTjVRbjd0TWhLVWdJbVVwOFBTWVU4UEZYeGxJVEdielFnIiwidGlkIjoiMWVkYWFkODMtYjJlZi00ODNkLTgxZjEtMmM0ODY4MmY0MGVjIiwidW5pcXVlX25hbWUiOiJQMTA0NzkxNTZAY2FwaXRhLmNvLnVrIiwidXBuIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInV0aSI6IlVFb205RzRRWVVPOTVPU21PTGxJQUEiLCJ2ZXIiOiIxLjAifQ.MfOWzI1HzJmHmJTZA0Tew8rV095qkzNSN0VQ0bOIXbh6qVkoDu6B2fa6Mdg2IJddWANvKNJf80H4dZeACLMKRjhKGCz4fJMJh44Y-yoksWkYQSB_81Fh0CHQeJzckvcUjopO98VzqyXIkWGPLao-K5OmwHAhk7gHo4oA9Iixh7icG2MgobjM9le_BAfo7y3UbsGj2iQ5pwPUuTgOnhFPQ3HhdrULxfLvwPHhnYd5xxPpr2-wQffjSjuClCZn5KjZgdzRNwgiC8ZijtcFtBScuSjco1h_D4BxrT0vM-MSVXbl786FyYOzdYwwN1JcFEwFyi9qFEwcY2uUc4ylO6-Kew");
                HttpResponseMessage response = await client.GetAsync(textTrack);
                if (response.IsSuccessStatusCode)
                {
                    var textTracksResponseString = await response.Content.ReadAsStringAsync();
                    TextTracksResponse textTracksResponse = JsonConvert.DeserializeObject<TextTracksResponse>(textTracksResponseString);
                    vttUrl = textTracksResponse?.value?[0].url;
                }
            }

            //using (HttpClient client = new HttpClient())
            //{
            //    client.BaseAddress = new Uri(vttUrl);
            //    client.DefaultRequestHeaders.Accept.Clear();
            //    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            //    client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5ODU4NjAsIm5iZiI6MTYwNjk4NTg2MCwiZXhwIjoxNjA2OTg5NzYwLCJhY3IiOiIxIiwiYWlvIjoiQVNRQTIvOFJBQUFBb2t0ZlFQQUNtbi9ROE1hWUp1MTduMldaOTE3QVlpY2wxRGkvS0NxdVRzTT0iLCJhbXIiOlsicHdkIl0sImFwcGlkIjoiY2Y1M2ZjZTgtZGVmNi00YWViLThkMzAtYjE1OGU3YjFjZjgzIiwiYXBwaWRhY3IiOiIyIiwiZmFtaWx5X25hbWUiOiJEaGFuZGUiLCJnaXZlbl9uYW1lIjoiRGlwYWsiLCJpcGFkZHIiOiI4Mi4yMDMuMzMuMTM0IiwibmFtZSI6IkRoYW5kZSwgRGlwYWsgKENhcGl0YSBTb2Z0d2FyZSkiLCJvaWQiOiJiYTdkYjc1YS0yZWY1LTQyNWYtOGVjNi1mOTg2YTFkMzI2MjIiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjM4NTc0OTg3LTI5MzUzODY4MTktMjA5MzY4NjEwLTI0Nzc4NjAiLCJwdWlkIjoiMTAwMzNGRkZBRkRDRTIyRSIsInJoIjoiMC5BQUFBZzYzYUh1LXlQVWlCOFN4SWFDOUE3T2o4VThfMjN1dEtqVEN4V09leHo0TUNBSTguIiwic2NwIjoiYWNjZXNzX21pY3Jvc29mdHN0cmVhbV9zZXJ2aWNlIiwic3ViIjoiTlZuN2dEWVZOYU41UW43dE1oS1VnSW1VcDhQU1lVOFBGWHhsSVRHYnpRZyIsInRpZCI6IjFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYyIsInVuaXF1ZV9uYW1lIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInVwbiI6IlAxMDQ3OTE1NkBjYXBpdGEuY28udWsiLCJ1dGkiOiJ2MGlNeWQ1VzRFYUNGSHZvSUhKQ0FBIiwidmVyIjoiMS4wIn0.eP7sqenBiO19DhrMFuuyY8-71B3mSzYI-iZxQd1UDHKPuzWmAL2wj0S4zVJhTVUpKVRLqqCVTGGvs8taGhfJ2E_L7OOXvNnzNw7Quc9xc88cfON2v-0A1gnwAPPTmE2JcVbQxg3dhKKvEWDwOkLAkSlaBOvh9oP_9fL9qhYM0qaRkeuIAy9AjuAmEpki4iL4VX6gSoPV7dKmqvgeZDgLLar3wD2LxWJqiXqYmpgHTm5mAFwfhaboZA1_w1IFyDqG7hIO-P6RlSZTx9fHfV3gZEmW8u0FXuBVNV-Rc8FtdITtzJQbDosOlk8S8yH67uFmt_aRJ5zSibo_ueUModSerQ");
            //    HttpResponseMessage response = await client.GetAsync(vttUrl);
            //    if (response.IsSuccessStatusCode)
            //    {
            //        var vttResponseString = await response.Content.ReadAsStringAsync();

            //        var vttResponse = JsonConvert.DeserializeObject<string>(vttResponseString);
            //    }
            //}

            using (var client = new WebClient())
            {

                client.Headers.Clear();
                
                client.Headers.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5OTk5MTYsIm5iZiI6MTYwNjk5OTkxNiwiZXhwIjoxNjA3MDAzODE2LCJhY3IiOiIxIiwiYWlvIjoiRTJSZ1lJZytkV2RCbjlGMS8yVnhEOFJZRFhZNzhiQmVpYk02SkoxNk5tUi8zY2NUT2hrQSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiJjZjUzZmNlOC1kZWY2LTRhZWItOGQzMC1iMTU4ZTdiMWNmODMiLCJhcHBpZGFjciI6IjIiLCJmYW1pbHlfbmFtZSI6IkRoYW5kZSIsImdpdmVuX25hbWUiOiJEaXBhayIsImlwYWRkciI6IjgyLjIwMy4zMy4xMzQiLCJuYW1lIjoiRGhhbmRlLCBEaXBhayAoQ2FwaXRhIFNvZnR3YXJlKSIsIm9pZCI6ImJhN2RiNzVhLTJlZjUtNDI1Zi04ZWM2LWY5ODZhMWQzMjYyMiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMzg1NzQ5ODctMjkzNTM4NjgxOS0yMDkzNjg2MTAtMjQ3Nzg2MCIsInB1aWQiOiIxMDAzM0ZGRkFGRENFMjJFIiwicmgiOiIwLkFBQUFnNjNhSHUteVBVaUI4U3hJYUM5QTdPajhVOF8yM3V0S2pUQ3hXT2V4ejRNQ0FJOC4iLCJzY3AiOiJhY2Nlc3NfbWljcm9zb2Z0c3RyZWFtX3NlcnZpY2UiLCJzdWIiOiJOVm43Z0RZVk5hTjVRbjd0TWhLVWdJbVVwOFBTWVU4UEZYeGxJVEdielFnIiwidGlkIjoiMWVkYWFkODMtYjJlZi00ODNkLTgxZjEtMmM0ODY4MmY0MGVjIiwidW5pcXVlX25hbWUiOiJQMTA0NzkxNTZAY2FwaXRhLmNvLnVrIiwidXBuIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInV0aSI6IlVFb205RzRRWVVPOTVPU21PTGxJQUEiLCJ2ZXIiOiIxLjAifQ.MfOWzI1HzJmHmJTZA0Tew8rV095qkzNSN0VQ0bOIXbh6qVkoDu6B2fa6Mdg2IJddWANvKNJf80H4dZeACLMKRjhKGCz4fJMJh44Y-yoksWkYQSB_81Fh0CHQeJzckvcUjopO98VzqyXIkWGPLao-K5OmwHAhk7gHo4oA9Iixh7icG2MgobjM9le_BAfo7y3UbsGj2iQ5pwPUuTgOnhFPQ3HhdrULxfLvwPHhnYd5xxPpr2-wQffjSjuClCZn5KjZgdzRNwgiC8ZijtcFtBScuSjco1h_D4BxrT0vM-MSVXbl786FyYOzdYwwN1JcFEwFyi9qFEwcY2uUc4ylO6-Kew");
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
                //    //    Console.WriteLine("----------------------");
                //}

                //Console.ReadLine();


                return View();
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
