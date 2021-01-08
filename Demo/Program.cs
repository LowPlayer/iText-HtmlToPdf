using iText.Html2pdf;
using iText.Html2pdf.Resolver.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.StyledXmlParser.Css.Media;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Demo
{
    class Program
    {
        static void Main(String[] args)
        {
            var pdfDest = "hello.pdf";
            var pdfWriter = new PdfWriter(pdfDest);
            var pdf = new PdfDocument(pdfWriter);

            var pageSize = PageSize.A4;       // 设置默认页面大小，css @page规则可覆盖这个
            pdf.SetDefaultPageSize(pageSize);

            var properties = new ConverterProperties();
            properties.SetBaseUri("wwwroot");     // 设置根目录
            properties.SetCharset("utf-8");

            var provider = new DefaultFontProvider(true, true, true);       // 第三个参数为True，以支持系统字体，否则不支持中文
            properties.SetFontProvider(provider);

            var mediaDeviceDescription = new MediaDeviceDescription(MediaType.PRINT);      // 指当前设备类型，如果是预览使用SCREEN
            mediaDeviceDescription.SetWidth(pageSize.GetWidth());

            properties.SetMediaDeviceDescription(mediaDeviceDescription);

            var peoples = new List<People>
            {
                new People { Avatar = "avatar.jpg", Name = "小明", Age = 23, Sex = "男" },
                new People { Avatar = "avatar.jpg", Name = "小王", Age = 18, Sex = "男" },
                new People { Avatar = "avatar.jpg", Name = "小樱", Age = 19, Sex = "女" },
                new People { Avatar = "avatar.jpg", Name = "小兰", Age = 20, Sex = "女" },
            };

            var htmlTemplate = File.ReadAllText("wwwroot/hello.html");

            Dictionary<String, String> dic = null;
            Regex regex = null;

            var start = "<!--template-start-->";
            var end = "<!--template-end-->";

            var match = Regex.Match(htmlTemplate, $@"{start}(.|\s)+?{end}");

            if (match != null && match.Value.Length > start.Length + end.Length)
            {
                var template = match.Value.Substring(start.Length, match.Value.Length - start.Length - end.Length);

                var sb = new StringBuilder(start);

                foreach (var people in peoples)
                {
                    dic = HtmlTemplateDataBuilder.Create(people);
                    regex = new Regex(String.Join("|", dic.Keys));
                    sb.Append(regex.Replace(template, m => dic[m.Value]));
                }

                sb.Append(end);

                htmlTemplate = htmlTemplate.Replace(match.Value, sb.ToString());
            }

            dic = new Dictionary<String, String> { ["{{ListOfNames}}"] = "人员列表" };
            regex = new Regex(String.Join("|", dic.Keys), RegexOptions.IgnoreCase);

            var html = regex.Replace(htmlTemplate, m => dic[m.Value]);
            HtmlConverter.ConvertToPdf(html, pdf, properties);
        }

        struct People
        {
            public String Avatar { get; set; }      // 头像
            public String Name { get; set; }        // 姓名
            public Int32 Age { get; set; }          // 年龄
            public String Sex { get; set; }         // 性别
        }

        public static class HtmlTemplateDataBuilder
        {
            public static Dictionary<String, String> Create(Object obj)
            {
                var props = obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

                var dic = new Dictionary<String, String>();

                foreach (var prop in props)
                {
                    dic.Add("{{" + prop.Name + "}}", prop.GetValue(obj, null).ToString());
                }

                return dic;
            }
        }
    }
}
