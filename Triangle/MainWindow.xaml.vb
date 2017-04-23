Imports Inventor
Imports System.Windows


Class MainWindow

    Dim _invApp As Inventor.Application = Nothing
    Dim _partDoc As PartDocument

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Try ' Try to get an active instance of Inventor

            Try
                _invApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")

            Catch ' If not active, create a new Inventor session

                Dim _invAppType As Type = System.Type.GetTypeFromProgID("Inventor.Application")

                _invApp = System.Activator.CreateInstance(_invAppType)

                'Must be set visible explicitly
                _invApp.Visible = True

            End Try

        Catch
            MsgBox("Error: couldn't create Inventor instance")

        End Try

        If _invApp.Documents.Count = 0 Then
            MsgBox("Need to open a Part Document")
            Return
        End If

        If _invApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            MsgBox("Need to have a Part Document active")
            Return
        End If


        'Dim _surface As WorkSurface
        '_surface = _invApp.ActiveDocument

    End Sub

    Private Sub Start_Click(sender As Object, e As RoutedEventArgs) Handles Start.Click
        _partDoc = _invApp.ActiveDocument
        Dim _triangle As WorkSurface
        _triangle = _partDoc.ComponentDefinition.WorkSurfaces.Item(1)
        Debug.Print(_triangle.SurfaceBodies.Item(1).Edges.Count)
        Dim Svertex As Vertex
        Svertex = _triangle.SurfaceBodies.Item(1).Edges.Item(1).StartVertex
        Debug.Print(Svertex.Point.X)

    End Sub
End Class
