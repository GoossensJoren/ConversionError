using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GemBox.Document;

namespace ConversionError
{
    class Program
    {
        static void Main(string[] args)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var path1 = Path.Combine(Directory.GetCurrentDirectory(), "contract1.docx");
            var path2 = Path.Combine(Directory.GetCurrentDirectory(), "contract2.docx");
            var path3 = Path.Combine(Directory.GetCurrentDirectory(), "contract3.docx");

            var document1 = DocumentModel.Load(path1);
            var document2 = DocumentModel.Load(path2);
            var document3 = DocumentModel.Load(path3);

            var data = new
            {
                first_name = "John"
            };

            document1.MailMerge.Execute(data);
            document2.MailMerge.Execute(data);

            var mergedDocument = MergeDocuments(new List<DocumentModel>{ document2, document1 });

            mergedDocument.Save("Output1.docx", SaveOptions.DocxDefault);
            document3.Save("Output2.pdf", SaveOptions.PdfDefault);
        }

        public static DocumentModel MergeDocuments(List<DocumentModel> documents)
        {
            var mergedDocument = new DocumentModel();

            foreach (DocumentModel document in documents)
            {
                var mapping = new ImportMapping(document, mergedDocument, false);

                List<Section> sections = document.Sections.Select(s => mergedDocument.Import(s, true, mapping)).ToList();
                sections.ForEach(
                    s =>
                    {
                        mergedDocument.Sections.Add(s);

                        if (s.Equals(sections.FirstOrDefault()))
                        {
                            s.PageSetup.SectionStart = SectionStart.NewPage;
                        }
                    });

                if (!document.Equals(documents.FirstOrDefault()))
                {
                    continue;
                }

                mergedDocument.DefaultCharacterFormat = document.DefaultCharacterFormat.Clone();
                mergedDocument.DefaultParagraphFormat = document.DefaultParagraphFormat.Clone();
            }

            return mergedDocument;
        }
    }
}
