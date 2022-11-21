Imports System
Imports System.Diagnostics
#Region "#usings"
Imports System.IO
Imports System.Reflection
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

#End Region  ' #usings
Namespace InlinePictures

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
#Region "#inlinepictures"
            Dim server As RichEditDocumentServer = New RichEditDocumentServer()
            server.LoadDocument("Texts\InlinePictures.rtf", DocumentFormat.Rtf)
            Dim doc As Document = server.Document
            ' Insert an image from a file.
            Dim rangeFound As DocumentRange = doc.FindAll("Visual Studio Magazine", SearchOptions.CaseSensitive)(0)
            Dim pos As DocumentPosition = doc.Paragraphs(doc.Paragraphs.Get(rangeFound.End).Index + 2).Range.Start
            doc.Images.Insert(pos, DocumentImageSource.FromFile("Pictures\ReadersChoice.png"))
            ' Insert an image from a stream.
            pos = doc.Paragraphs(4).Range.Start
            Dim imageToInsert As String = "information.png"
            Dim a As Assembly = Assembly.GetExecutingAssembly()
            Dim imageStream As Stream = a.GetManifestResourceStream("InlinePictures.Resources." & imageToInsert)
            doc.Images.Insert(pos, DocumentImageSource.FromStream(imageStream))
            ' Insert an image using its URI.
            Dim imageUri As String = "http://i.gyazo.com/798a2ed48a3535c6c8add0ea7a4fc4e6.png"
            Dim docHeader As SubDocument = doc.Sections(0).BeginUpdateHeader()
            docHeader.Images.Append(DocumentImageSource.FromUri(imageUri, server))
            doc.Sections(0).EndUpdateHeader(docHeader)
            ' Insert a barcode.
            Dim barCode As DevExpress.BarCodes.BarCode = New DevExpress.BarCodes.BarCode()
            barCode.Symbology = DevExpress.BarCodes.Symbology.QRCode
            barCode.CodeText = "http://www.devexpress.com"
            barCode.CodeBinaryData = Text.Encoding.Default.GetBytes(barCode.CodeText)
            barCode.Module = 0.5
            Dim docFooter As SubDocument = doc.Sections(0).BeginUpdateFooter()
            docFooter.Images.Append(DocumentImageSource.FromImage(barCode.BarCodeImage))
            doc.Sections(0).EndUpdateFooter(docFooter)
#End Region  ' #inlinepictures
#Region "#getimages"
            ' Scale down images in the document body.
            Dim images As ReadOnlyDocumentImageCollection = server.Document.Images.Get(doc.Range)
            For i As Integer = 0 To images.Count - 1
                images(i).ScaleX /= 4
                images(i).ScaleY /= 4
            Next

#End Region  ' #getimages
            ' Save the resulting document.
            server.SaveDocument("InlinePictures.docx", DocumentFormat.OpenXml)
            Call Process.Start("InlinePictures.docx")
        End Sub
    End Module
End Namespace
