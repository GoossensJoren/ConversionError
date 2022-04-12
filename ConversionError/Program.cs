using System.Linq;
using GemBox.Document;

namespace ConversionError
{
    class Program
    {
        static void Main(string[] args)
        {
            ComponentInfo.SetLicense("");

            string[] files = { "doc6.docx", "doc7.docx", "doc8.docx", "doc9.docx", "doc10.docx", "doc11.docx", "doc12.docx", "doc13.docx", "doc14.docx", "doc15.docx", "doc16.docx" };

            var destination = new DocumentModel();
            bool first = true;

            foreach (string file in files)
            {
                var source = DocumentModel.Load(file);
                var mapping = new ImportMapping(source, destination, false);

                if (source.Styles.Contains("Hyperlink") && destination.Styles.Contains("Hyperlink"))
                {
                    mapping.SetDestinationStyle(source.Styles["Hyperlink"], destination.Styles["Hyperlink"]);
                }

                var sourceSections = source.Sections;

                Section firstSection = source.Sections.First();
                if (firstSection != null)
                {
                    firstSection.PageSetup.PageStartingNumber = 1;
                }

                foreach (var section in sourceSections)
                {
                    var importedSection = destination.Import(section, true, mapping);
                    destination.Sections.Add(importedSection);
                }

                if (!first)
                {
                    continue;
                }

                destination.DefaultCharacterFormat = source.DefaultCharacterFormat.Clone();
                destination.DefaultParagraphFormat = source.DefaultParagraphFormat.Clone();
                first = false;
            }

            destination.Save("output.docx");
        }
    }
}
