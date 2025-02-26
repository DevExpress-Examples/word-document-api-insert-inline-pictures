Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace InlinePictures
    Class Program
        Shared Sub Main(ByVal args As String())
            Using wordProcessor As New RichEditDocumentServer()
                wordProcessor.LoadDocument("Texts\InlinePictures.rtf", DocumentFormat.Rtf)
                Dim document As Document = wordProcessor.Document

                ' Insert an image from a file.
                Dim rangeFound As DocumentRange = document.FindAll("Visual Studio Magazine", SearchOptions.CaseSensitive)(0)
                Dim pos As DocumentPosition = document.Paragraphs(document.Paragraphs.Get(rangeFound.End).Index + 2).Range.Start
                Dim imageFromFile As Shape = document.Shapes.InsertPicture(pos, DocumentImageSource.FromFile("Pictures\ReadersChoice.png"))
                imageFromFile.TextWrapping = TextWrappingType.InLineWithText

                ' Insert an image from a stream.
                pos = document.Paragraphs(4).Range.Start
                Dim imageToInsert As String = "information.png"
                Dim a As Assembly = Assembly.GetExecutingAssembly()
                Dim imageStream As Stream = a.GetManifestResourceStream(imageToInsert)
                Dim imageFromStream As Shape = document.Shapes.InsertPicture(pos, DocumentImageSource.FromStream(imageStream))
                imageFromStream.TextWrapping = TextWrappingType.InLineWithText

                ' Insert an image using its URI.
                Dim imageUri As String = "http://i.gyazo.com/798a2ed48a3535c6c8add0ea7a4fc4e6.png"
                Dim docHeader As SubDocument = document.Sections(0).BeginUpdateHeader()
                Dim headerImage As Shape = docHeader.Shapes.InsertPicture(docHeader.Range.End, DocumentImageSource.FromUri(imageUri, wordProcessor))
                headerImage.TextWrapping = TextWrappingType.InLineWithText

                ' Save the resulting document.
                wordProcessor.SaveDocument("InlinePictures.docx", DocumentFormat.OpenXml)
            End Using

            Dim p As New Process()
            p.StartInfo = New ProcessStartInfo("InlinePictures.docx") With {
                .UseShellExecute = True
            }
            p.Start()
        End Sub
    End Class
End Namespace
