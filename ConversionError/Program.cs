using System;
using System.IO;
using GemBox.Document;

namespace ConversionError
{
    class Program
    {
        static void Main(string[] args)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var path = Path.Combine(Directory.GetCurrentDirectory(), "contract1.docx");
            var document = DocumentModel.Load(path);

            var data = new
            {
                first_name = "John"
            };

            document.MailMerge.Execute(data);

            using var stream = new MemoryStream();
            document.Save(stream, SaveOptions.DocxDefault);

            var documentData = stream.ToArray();
        }
    }
}
