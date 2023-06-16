Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace InlinePictures

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            Using wordProcessor As RichEditDocumentServer = New RichEditDocumentServer()
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
                Dim imageStream As Stream = a.GetManifestResourceStream("InlinePictures.Resources." & imageToInsert)
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

            Call Process.Start("InlinePictures.docx")
        End Sub
    End Module
End Namespace
