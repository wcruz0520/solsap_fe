Imports System.Threading
Imports System.Windows.Forms

Public Enum DialogType

    SAVE

    OPEN

    FOLDER

End Enum
Public Class SelectFileDialog

    Private shutdownEvent As ManualResetEvent = New ManualResetEvent(False)

    Public Property SelectedFile As String


    Public Property SelectedFolder As String


    Private folder As String

    Private file As String

    Private filter As String

    Private type As DialogType

    Sub New(ByVal folder As String, ByVal file As String, ByVal filter As String, ByVal type As DialogType)

        If (folder Is Nothing _
                   Or file Is Nothing _
                  Or filter Is Nothing) Then
            Throw New ArgumentException("Error de Parametros")
        End If

        Me.folder = folder
        Me.file = file
        Me.filter = filter
        Me.type = type

    End Sub



    Private Sub InternalSelectFileDialog()
        Dim form = New System.Windows.Forms.Form
        form.TopMost = True
        form.Height = 0
        form.Width = 0
        form.WindowState = FormWindowState.Minimized
        form.Visible = True
        Select Case (Me.type)
            Case DialogType.FOLDER
                Me.FolderDialog(form)
            Case DialogType.OPEN
                Me.OpenDialog(form)
            Case DialogType.SAVE
                Me.SaveDialog(form)
        End Select

        Me.shutdownEvent.Set()
    End Sub

    Private Sub FolderDialog(ByVal form As System.Windows.Forms.Form)
        Dim dialog As FolderBrowserDialog = New FolderBrowserDialog
        dialog.Description = "Selecccionar ruta para guardar los archivos seleccionados"
        'dialog.RootFolder = Environment.SpecialFolder.MyComputer
        dialog.ShowNewFolderButton = True
        '----------------------------------------------------------------//
        If (dialog.ShowDialog = DialogResult.OK) Then
            form.Close()
            Me.SelectedFolder = dialog.SelectedPath
        Else
            form.Close()
            Me.SelectedFolder = ""
        End If

    End Sub

    Private Sub OpenDialog(ByVal form As System.Windows.Forms.Form)
        Dim dialog As OpenFileDialog = New OpenFileDialog
        Me.OpenOrSaveDialog(dialog, form)
    End Sub

    Private Sub SaveDialog(ByVal form As System.Windows.Forms.Form)
        Dim dialog As SaveFileDialog = New SaveFileDialog
        Me.OpenOrSaveDialog(dialog, form)
    End Sub

    Private Sub OpenOrSaveDialog(ByVal dialog As FileDialog, ByVal form As System.Windows.Forms.Form)
        dialog.Title = "Seleccione el archivo por favor"
        dialog.Filter = Me.filter
        '"TXT files (*.txt)|*.txt|All files (*.*)|*.*";
        dialog.InitialDirectory = Me.folder
        dialog.FileName = Me.file
        '----------------------------------------------------------------//
        If (dialog.ShowDialog = DialogResult.OK) Then
            form.Close()
            Me.SelectedFile = dialog.FileName
        Else
            form.Close()
            Me.SelectedFile = ""
        End If

    End Sub

    Public Sub Open()
        Dim t As Thread = New Thread(AddressOf InternalSelectFileDialog)
        t.SetApartmentState(ApartmentState.STA)
        t.Start()
        Me.shutdownEvent.WaitOne()
    End Sub
End Class