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
            ComponentInfo.SetLicense("");

            string[] files = { "contract3.docx", "contract4.docx", "contract1.docx" };

            var destination = new DocumentModel();
            bool first = true;

            foreach (string file in files)
            {
                var source = DocumentModel.Load(file);
                var mapping = new ImportMapping(source, destination, false);

                var sourceSections = source.Sections;
                var destinationSections = new List<Section>();

                foreach (var section in sourceSections)
                {
                    var importedSection = destination.Import(section, true, mapping);
                    destination.Sections.Add(importedSection);
                    destinationSections.Add(importedSection);
                }

                if (first)
                {
                    destination.DefaultCharacterFormat = source.DefaultCharacterFormat.Clone();
                    destination.DefaultParagraphFormat = source.DefaultParagraphFormat.Clone();
                    first = false;
                }
                else
                {
                    FixDefaultFormatting(sourceSections, destinationSections);
                }
            }

            destination.Save("output.docx");
        }

        static void FixDefaultFormatting(IList<Section> sourceSections, IList<Section> destinationSections)
        {
            for (int sourceIndex = 0; sourceIndex < sourceSections.Count; sourceIndex++)
            {
                var sourceSection = sourceSections[sourceIndex];
                var destinationSection = destinationSections[sourceIndex];

                var sourceParagraphs = sourceSection.GetChildElements(true, ElementType.Paragraph).Cast<Paragraph>().ToList();
                var destinationParagraphs = destinationSection.GetChildElements(true, ElementType.Paragraph).Cast<Paragraph>().ToList();

                for (int paragraphIndex = 0; paragraphIndex < sourceParagraphs.Count; paragraphIndex++)
                {
                    var sourceParagraph = sourceParagraphs[paragraphIndex];
                    var destinationParagraph = destinationParagraphs[paragraphIndex];

                    FixParagraphFormat(sourceParagraph.ParagraphFormat, destinationParagraph.ParagraphFormat, f => f.LeftIndentation, (f1, f2) => f1.LeftIndentation = f2.LeftIndentation);
                    FixParagraphFormat(sourceParagraph.ParagraphFormat, destinationParagraph.ParagraphFormat, f => f.RightIndentation, (f1, f2) => f1.RightIndentation = f2.RightIndentation);
                    FixParagraphFormat(sourceParagraph.ParagraphFormat, destinationParagraph.ParagraphFormat, f => f.SpecialIndentation, (f1, f2) => f1.SpecialIndentation = f2.SpecialIndentation);
                    FixParagraphFormat(sourceParagraph.ParagraphFormat, destinationParagraph.ParagraphFormat, f => f.SpaceAfter, (f1, f2) => f1.SpaceAfter = f2.SpaceAfter);
                    FixParagraphFormat(sourceParagraph.ParagraphFormat, destinationParagraph.ParagraphFormat, f => f.SpaceBefore, (f1, f2) => f1.SpaceBefore = f2.SpaceBefore);

                    var sourceInlines = sourceParagraph.GetChildElements(true, ElementType.Run, ElementType.Field, ElementType.SpecialCharacter).ToList();
                    var destinationInlines = destinationParagraph.GetChildElements(true, ElementType.Run, ElementType.Field, ElementType.SpecialCharacter).ToList();

                    for (int inlineIndex = 0; inlineIndex < sourceInlines.Count; inlineIndex++)
                    {
                        CharacterFormat sourceCharacterFormat, destinationCharacterFormat;
                        switch (sourceInlines[inlineIndex].ElementType)
                        {
                            case ElementType.Run:
                                sourceCharacterFormat = ((Run)sourceInlines[inlineIndex]).CharacterFormat;
                                destinationCharacterFormat = ((Run)destinationInlines[inlineIndex]).CharacterFormat;
                                break;
                            case ElementType.Field:
                                sourceCharacterFormat = ((Field)sourceInlines[inlineIndex]).CharacterFormat;
                                destinationCharacterFormat = ((Field)destinationInlines[inlineIndex]).CharacterFormat;
                                break;
                            case ElementType.SpecialCharacter:
                                sourceCharacterFormat = ((SpecialCharacter)sourceInlines[inlineIndex]).CharacterFormat;
                                destinationCharacterFormat = ((SpecialCharacter)destinationInlines[inlineIndex]).CharacterFormat;
                                break;
                            default:
                                throw new InvalidOperationException();
                        }

                        FixCharacterFormat(sourceCharacterFormat, destinationCharacterFormat, f => f.Bold, (f1, f2) => f1.Bold = f2.Bold);
                        FixCharacterFormat(sourceCharacterFormat, destinationCharacterFormat, f => f.FontName, (f1, f2) => f1.FontName = f2.FontName);
                        FixCharacterFormat(sourceCharacterFormat, destinationCharacterFormat, f => f.Italic, (f1, f2) => f1.Italic = f2.Italic);
                        FixCharacterFormat(sourceCharacterFormat, destinationCharacterFormat, f => f.Size, (f1, f2) => f1.Size = f2.Size);
                    }
                }
            }
        }

        static void FixParagraphFormat(ParagraphFormat source, ParagraphFormat destination, Func<ParagraphFormat, double> getter, Action<ParagraphFormat, ParagraphFormat> setter)
        {
            var sourceValue = getter(source);
            var destinationValue = getter(destination);
            if (sourceValue != destinationValue)
                setter(destination, source);
        }

        static void FixCharacterFormat(CharacterFormat source, CharacterFormat destination, Func<CharacterFormat, object> getter, Action<CharacterFormat, CharacterFormat> setter)
        {
            var sourceValue = getter(source);
            var destinationValue = getter(destination);
            if (!sourceValue.Equals(destinationValue))
                setter(destination, source);
        }

    }
}
