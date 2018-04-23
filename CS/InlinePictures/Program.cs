using System;
using System.Collections.Generic;
using System.Diagnostics;
#region #usings
using System.IO;
using System.Reflection;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion #usings

namespace InlinePictures {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            #region #inlinepictures
            RichEditDocumentServer server = new RichEditDocumentServer();
            server.LoadDocument("Texts\\InlinePictures.rtf", DocumentFormat.Rtf);
            Document doc = server.Document;
            
            // Insert an image from a file.
            DocumentRange rangeFound = doc.FindAll("Visual Studio Magazine", SearchOptions.CaseSensitive)[0];
            DocumentPosition pos = doc.Paragraphs[doc.Paragraphs.Get(rangeFound.End).Index + 2].Range.Start;
            doc.Images.Insert(pos, DocumentImageSource.FromFile("Pictures\\ReadersChoice.png"));
            
            // Insert an image from a stream.
            pos = doc.Paragraphs[4].Range.Start;
            string imageToInsert = "information.png";
            Assembly a = Assembly.GetExecutingAssembly();
            Stream imageStream = a.GetManifestResourceStream("InlinePictures.Resources." + imageToInsert);
            doc.Images.Insert(pos, DocumentImageSource.FromStream(imageStream));
            
            // Insert an image using its URI.
            string imageUri = "http://i.gyazo.com/798a2ed48a3535c6c8add0ea7a4fc4e6.png";
            SubDocument docHeader = doc.Sections[0].BeginUpdateHeader();
            docHeader.Images.Append(DocumentImageSource.FromUri(imageUri, server));
            doc.Sections[0].EndUpdateHeader(docHeader);


            // Insert a barcode.
            DevExpress.BarCodes.BarCode barCode = new DevExpress.BarCodes.BarCode();
            barCode.Symbology = DevExpress.BarCodes.Symbology.QRCode;
            barCode.CodeText = "http://www.devexpress.com";
            barCode.CodeBinaryData = System.Text.Encoding.Default.GetBytes(barCode.CodeText);
            barCode.Module = 0.5;
            SubDocument docFooter = doc.Sections[0].BeginUpdateFooter();
            docFooter.Images.Append(barCode.BarCodeImage);
            doc.Sections[0].EndUpdateFooter(docFooter);
            #endregion #inlinepictures

            #region #getimages
            // Scale down images in the document body.
           ReadOnlyDocumentImageCollection images = server.Document.Images.Get(doc.Range);
            for (int i = 0; i < images.Count; i++)
            {
                images[i].ScaleX /= 4;
                images[i].ScaleY /= 4;
            }
            #endregion #getimages
            // Save the resulting document.
            server.SaveDocument("InlinePictures.docx", DocumentFormat.OpenXml);
            Process.Start("InlinePictures.docx");            
        }
    }
}