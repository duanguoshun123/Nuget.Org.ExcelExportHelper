using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace Nuget.Org.Npoi.Web.Infrastructure
{
    public class FileActionResult : IHttpActionResult
    {
        public FileActionResult(Stream stream, string mediaType, string filename, string browser = null)
        {
            this.Stream = stream;
            this.MediaType = mediaType ?? @"application/octet-stream";
            this.Filename = filename ?? "";
            //this.Browser = browser;
        }

        public Stream Stream { get; }
        public string Filename { get; }
        public string MediaType { get; }
        //public string Browser { get; }

        public Task<System.Net.Http.HttpResponseMessage> ExecuteAsync(System.Threading.CancellationToken cancellationToken)
        {
            var response = new HttpResponseMessage
            {
                Content = new StreamContent(Stream) { },
            };
            HeadersSetContentDispositionFilename(response.Content.Headers, false, this.Filename);
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(MediaType);
            return Task.FromResult(response);
        }

        /// <summary>
        /// HeadersSetContentDispositionFilename
        /// </summary>
        /// <param name="headers"></param>
        /// <param name="inline">true为: inline; ，false为: attachment; </param>
        /// <param name="filename"></param>
        /// <param name="browser"></param>
        public void HeadersSetContentDispositionFilename(System.Net.Http.Headers.HttpHeaders headers, bool inline, string filename, string browser = null)
        {
            if (headers is System.Net.Http.Headers.HttpContentHeaders contentHeaders)
            {
                contentHeaders.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue(inline ? "inline" : "attachment")
                {
                    FileNameStar = filename,
                };
            }
            else
            {
                if (headers.Contains(@"Content-Disposition"))
                {
                    headers.Remove(@"Content-Disposition");
                }
                headers.Add(@"Content-Disposition", EncodeFileDownloadNameContentDisposition4(inline, filename));
            }
        }

        // 支持: Firefox Chrome IE
        [Localizable(false)]
        private static string EncodeFileDownloadNameContentDisposition4(bool inline, string filename)
        {
            var type = inline ? "inline" : "attachment";
            return new System.Net.Http.Headers.ContentDispositionHeaderValue(type)
            {
                FileNameStar = filename,
            }.ToString();
        }
    }
}