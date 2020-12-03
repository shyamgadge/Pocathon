using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
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
            //string textTrack = "https://euno-1.api.microsoftstream.com/api/videos/a606d397-8653-4ce5-ba6a-12b5a87a2016/texttracks?api-version=1.4-private";
            string textTrack = "https://euno-1.api.microsoftstream.com/api/videos/" + videoId + "/texttracks?api-version=1.4-private";

            string vttUrl = "";
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(textTrack);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5Nzk0MzUsIm5iZiI6MTYwNjk3OTQzNSwiZXhwIjoxNjA2OTgzMzM1LCJhY3IiOiIxIiwiYWlvIjoiRTJSZ1lMRGJvOU5acEZIRHhQNDk2L0plcFZOTjI3K3gxM3kyc3lxcWxWTS8rV0hLN3U4QSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiJjZjUzZmNlOC1kZWY2LTRhZWItOGQzMC1iMTU4ZTdiMWNmODMiLCJhcHBpZGFjciI6IjIiLCJmYW1pbHlfbmFtZSI6IkRoYW5kZSIsImdpdmVuX25hbWUiOiJEaXBhayIsImlwYWRkciI6IjgyLjIwMy4zMy4xMzQiLCJuYW1lIjoiRGhhbmRlLCBEaXBhayAoQ2FwaXRhIFNvZnR3YXJlKSIsIm9pZCI6ImJhN2RiNzVhLTJlZjUtNDI1Zi04ZWM2LWY5ODZhMWQzMjYyMiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMzg1NzQ5ODctMjkzNTM4NjgxOS0yMDkzNjg2MTAtMjQ3Nzg2MCIsInB1aWQiOiIxMDAzM0ZGRkFGRENFMjJFIiwicmgiOiIwLkFBQUFnNjNhSHUteVBVaUI4U3hJYUM5QTdPajhVOF8yM3V0S2pUQ3hXT2V4ejRNQ0FJOC4iLCJzY3AiOiJhY2Nlc3NfbWljcm9zb2Z0c3RyZWFtX3NlcnZpY2UiLCJzdWIiOiJOVm43Z0RZVk5hTjVRbjd0TWhLVWdJbVVwOFBTWVU4UEZYeGxJVEdielFnIiwidGlkIjoiMWVkYWFkODMtYjJlZi00ODNkLTgxZjEtMmM0ODY4MmY0MGVjIiwidW5pcXVlX25hbWUiOiJQMTA0NzkxNTZAY2FwaXRhLmNvLnVrIiwidXBuIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInV0aSI6InR6Y1dqUXpVQkUtUUpzRlYtQ1U4QUEiLCJ2ZXIiOiIxLjAifQ.AWOcmMMMaj722KiYTuCoVZZ4G5BrxpsK__gsNZmdgod3iGf1AzTaH90q6c9Uu9albeI2NYR79r6GJX0wETCd5XFGkxhsJ3D_Pb5HH4VdG9qatoWoz62EmigoPg3bZgIyEMB7tL2DPqS40OiOnFwAYbOpiGvWa1edarOgMSV6Fro1fbJV1Ue0TZZdeQJ7mjqA176j-efU74M8Uqms_IhnHPZv8olpRxVUeOnhIwc6FMqIA4HbYTOY6XgffRH_1fj4tUlHMhVTBptSpDZSfPvUry1pcaKZq7i0HjYcIAbBXmzGiYPGh_7JyEzn-A2SpbCLy72gE1KSxoRDpmnkINHuew");
                HttpResponseMessage response = await client.GetAsync(textTrack);
                if (response.IsSuccessStatusCode)
                {
                    var textTracksResponseString = await response.Content.ReadAsStringAsync();
                    TextTracksResponse textTracksResponse = JsonConvert.DeserializeObject<TextTracksResponse>(textTracksResponseString);

                    vttUrl = textTracksResponse?.value?[0].url;
                }
            }

            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(vttUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCIsImtpZCI6ImtnMkxZczJUMENUaklmajRydDZKSXluZW4zOCJ9.eyJhdWQiOiJodHRwczovLyoubWljcm9zb2Z0c3RyZWFtLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzFlZGFhZDgzLWIyZWYtNDgzZC04MWYxLTJjNDg2ODJmNDBlYy8iLCJpYXQiOjE2MDY5Nzk0MzUsIm5iZiI6MTYwNjk3OTQzNSwiZXhwIjoxNjA2OTgzMzM1LCJhY3IiOiIxIiwiYWlvIjoiRTJSZ1lMRGJvOU5acEZIRHhQNDk2L0plcFZOTjI3K3gxM3kyc3lxcWxWTS8rV0hLN3U4QSIsImFtciI6WyJwd2QiXSwiYXBwaWQiOiJjZjUzZmNlOC1kZWY2LTRhZWItOGQzMC1iMTU4ZTdiMWNmODMiLCJhcHBpZGFjciI6IjIiLCJmYW1pbHlfbmFtZSI6IkRoYW5kZSIsImdpdmVuX25hbWUiOiJEaXBhayIsImlwYWRkciI6IjgyLjIwMy4zMy4xMzQiLCJuYW1lIjoiRGhhbmRlLCBEaXBhayAoQ2FwaXRhIFNvZnR3YXJlKSIsIm9pZCI6ImJhN2RiNzVhLTJlZjUtNDI1Zi04ZWM2LWY5ODZhMWQzMjYyMiIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMzg1NzQ5ODctMjkzNTM4NjgxOS0yMDkzNjg2MTAtMjQ3Nzg2MCIsInB1aWQiOiIxMDAzM0ZGRkFGRENFMjJFIiwicmgiOiIwLkFBQUFnNjNhSHUteVBVaUI4U3hJYUM5QTdPajhVOF8yM3V0S2pUQ3hXT2V4ejRNQ0FJOC4iLCJzY3AiOiJhY2Nlc3NfbWljcm9zb2Z0c3RyZWFtX3NlcnZpY2UiLCJzdWIiOiJOVm43Z0RZVk5hTjVRbjd0TWhLVWdJbVVwOFBTWVU4UEZYeGxJVEdielFnIiwidGlkIjoiMWVkYWFkODMtYjJlZi00ODNkLTgxZjEtMmM0ODY4MmY0MGVjIiwidW5pcXVlX25hbWUiOiJQMTA0NzkxNTZAY2FwaXRhLmNvLnVrIiwidXBuIjoiUDEwNDc5MTU2QGNhcGl0YS5jby51ayIsInV0aSI6InR6Y1dqUXpVQkUtUUpzRlYtQ1U4QUEiLCJ2ZXIiOiIxLjAifQ.AWOcmMMMaj722KiYTuCoVZZ4G5BrxpsK__gsNZmdgod3iGf1AzTaH90q6c9Uu9albeI2NYR79r6GJX0wETCd5XFGkxhsJ3D_Pb5HH4VdG9qatoWoz62EmigoPg3bZgIyEMB7tL2DPqS40OiOnFwAYbOpiGvWa1edarOgMSV6Fro1fbJV1Ue0TZZdeQJ7mjqA176j-efU74M8Uqms_IhnHPZv8olpRxVUeOnhIwc6FMqIA4HbYTOY6XgffRH_1fj4tUlHMhVTBptSpDZSfPvUry1pcaKZq7i0HjYcIAbBXmzGiYPGh_7JyEzn-A2SpbCLy72gE1KSxoRDpmnkINHuew");
                HttpResponseMessage response = await client.GetAsync(vttUrl);
                if (response.IsSuccessStatusCode)
                {
                    // Get the vtt file content

                }
            }

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
