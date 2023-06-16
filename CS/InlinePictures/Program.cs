using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace InlinePictures
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.LoadDocument("Texts\\InlinePictures.rtf", DocumentFormat.Rtf);
                Document document = wordProcessor.Document;

                // Insert an image from a file.
                DocumentRange rangeFound = document.FindAll("Visual Studio Magazine", SearchOptions.CaseSensitive)[0];
                DocumentPosition pos = document.Paragraphs[document.Paragraphs.Get(rangeFound.End).Index + 2].Range.Start;
                Shape imageFromFile = document.Shapes.InsertPicture(pos, DocumentImageSource.FromFile("Pictures\\ReadersChoice.png"));
                imageFromFile.TextWrapping = TextWrappingType.InLineWithText;

                // Insert an image from a stream.
                pos = document.Paragraphs[4].Range.Start;
                string imageToInsert = "information.png";
                Assembly a = Assembly.GetExecutingAssembly();
                Stream imageStream = a.GetManifestResourceStream("InlinePictures.Resources." + imageToInsert);
               Shape imageFromStream =  document.Shapes.InsertPicture(pos, DocumentImageSource.FromStream(imageStream));
                imageFromStream.TextWrapping = TextWrappingType.InLineWithText;

                // Insert an image using its URI.
                string imageUri = "http://i.gyazo.com/798a2ed48a3535c6c8add0ea7a4fc4e6.png";
                SubDocument docHeader = document.Sections[0].BeginUpdateHeader();
                Shape headerImage = docHeader.Shapes.InsertPicture(docHeader.Range.End, DocumentImageSource.FromUri(imageUri, wordProcessor));
                headerImage.TextWrapping = TextWrappingType.InLineWithText;

                // Save the resulting document.
                wordProcessor.SaveDocument("InlinePictures.docx", DocumentFormat.OpenXml);
            }
                Process.Start("InlinePictures.docx");
        }
    }
}