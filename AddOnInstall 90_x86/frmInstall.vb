'//  SAP MANAGE DI API 6.7 SDK Sample
'//****************************************************************************
'//
'//  File:      frmInstall.vb
'//
'//  Copyright (c) SAP MANAGE
'//
'// INSTALACION DE ADD-ONS PARA SBO.
'//
'//****************************************************************************
'// This sample creates an add-on installer for SBO.
'// An installation for SBO should be build in a spesific way.
'// 1) It should be able to accept a command line parameter from SBO.
'//    This parameter is a string built from 2 strings devided by "|".
'//    The first string is the path recommended by SBO for installation folder.
'//    The second string is the location of "AddOnInstallAPI.dll".
'//    For example, a command line parameter that looks like this:
'//    "C:\MyAddon|C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
'//    Means that the recommended installation folder for this addon is "C:\MyAddon"
'//    and the location of "AddOnInstallAPI.dll" is - 
'//                 "C:\Program Files\SAP Manage\SAP Business One\AddOnInstallAPI.dll"
'// 2) When the installation is complete the installer must call the function 
'//    "EndInstall" from "AddOnInstallAPI.dll" to inform SBO the installation is complete.
'//    This dll contains 3 functions that can be used during the installation.
'//    The functions are: 
'//         1) EndInstall - Signals SBO that the installation is complete.
'//         2) SetAddOnFolder - Use it if you want to change the installation folder.
'//         3) RestartNeeded - Use it if your installation requires a restart, it will cause
'//            the SBO application to close itself after the installation is complete.
'//    All 3 functions return a 32 bit integer. There are 2 possible values for this integer.
'//    0 - Success, 1 - Failure.
'// 3) The installer must be one executable file.
'// 4) After your installer is ready you need to create an add-on registration file.
'//    In order to create it you have a utility - "Add-On Registration Data Creator"
'//    you can find it in -
'//       "..\SAP Manage\SAP Business One SDK\Tools\AddOnRegDataGen\AddOnRegDataGen.exe".
'//    This utility creates a file with the extention 'ard', you will be asked to 
'//    point to this file when you register your addon.

Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Class frmInstall
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblHeadLine As System.Windows.Forms.Label
    'Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDest As System.Windows.Forms.TextBox
    Friend WithEvents chkRestart As System.Windows.Forms.CheckBox
    Friend WithEvents chkDefaultFolder As System.Windows.Forms.CheckBox
    Friend WithEvents cmdInstall As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox

    Friend WithEvents FileWatcher As System.IO.FileSystemWatcher
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInstall))
        Me.cmdInstall = New System.Windows.Forms.Button()
        Me.lblHeadLine = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDest = New System.Windows.Forms.TextBox()
        Me.chkRestart = New System.Windows.Forms.CheckBox()
        Me.chkDefaultFolder = New System.Windows.Forms.CheckBox()
        Me.FileWatcher = New System.IO.FileSystemWatcher()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.FileWatcher, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdInstall
        '
        Me.cmdInstall.Location = New System.Drawing.Point(480, 218)
        Me.cmdInstall.Name = "cmdInstall"
        Me.cmdInstall.Size = New System.Drawing.Size(115, 37)
        Me.cmdInstall.TabIndex = 1
        Me.cmdInstall.Text = "Instalación"
        '
        'lblHeadLine
        '
        Me.lblHeadLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.lblHeadLine.Location = New System.Drawing.Point(19, 18)
        Me.lblHeadLine.Name = "lblHeadLine"
        Me.lblHeadLine.Size = New System.Drawing.Size(586, 28)
        Me.lblHeadLine.TabIndex = 2
        Me.lblHeadLine.Text = "Instalación - GS.EDOC.SAP"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(29, 146)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(547, 27)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Installation Folder recieved from SBO application"
        '
        'txtDest
        '
        Me.txtDest.Enabled = False
        Me.txtDest.Location = New System.Drawing.Point(29, 165)
        Me.txtDest.Name = "txtDest"
        Me.txtDest.Size = New System.Drawing.Size(566, 22)
        Me.txtDest.TabIndex = 5
        '
        'chkRestart
        '
        Me.chkRestart.Location = New System.Drawing.Point(29, 229)
        Me.chkRestart.Name = "chkRestart"
        Me.chkRestart.Size = New System.Drawing.Size(125, 28)
        Me.chkRestart.TabIndex = 6
        Me.chkRestart.Text = "Ask for a restart"
        '
        'chkDefaultFolder
        '
        Me.chkDefaultFolder.Checked = True
        Me.chkDefaultFolder.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDefaultFolder.Location = New System.Drawing.Point(29, 202)
        Me.chkDefaultFolder.Name = "chkDefaultFolder"
        Me.chkDefaultFolder.Size = New System.Drawing.Size(288, 27)
        Me.chkDefaultFolder.TabIndex = 7
        Me.chkDefaultFolder.Text = "Use path supplied by SBO"
        '
        'FileWatcher
        '
        Me.FileWatcher.EnableRaisingEvents = True
        Me.FileWatcher.SynchronizingObject = Me
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(252, 49)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(343, 88)
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        '
        'frmInstall
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(627, 277)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.chkDefaultFolder)
        Me.Controls.Add(Me.chkRestart)
        Me.Controls.Add(Me.txtDest)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblHeadLine)
        Me.Controls.Add(Me.cmdInstall)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmInstall"
        Me.Text = "Instalación - AddOn "
        CType(Me.FileWatcher, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Data members"
    Private sAddonName As String = "slnGSEDOC_SAP_EC"
    Private sInstallName As String = "AddOnInstaller.exe"
    Private strDll As String ' The path of "AddOnInstallAPI.dll"
    Private strDest As String ' Installation target path
    Private bFileCreated As Boolean ' True if the file was created
#End Region

#Region "Declarations"
    ' Declaring the functions inside "AddOnInstallAPI.dll"

    'EndInstall - Signals SBO that the installation is complete.
    Declare Function EndInstall Lib "AddOnInstallAPI.dll" () As Int32
    'SetAddOnFolder - Use it if you want to change the installation folder.
    Declare Function SetAddOnFolder Lib "AddOnInstallAPI.dll" (ByVal srrPath As String) As Int32
    'RestartNeeded - Use it if your installation requires a restart, it will cause
    'the SBO application to close itself after the installation is complete.
    Declare Function RestartNeeded Lib "AddOnInstallAPI.dll" () As Int32
#End Region

#Region "Methods"

    ' Read the addon path from the registry
    Public Function ReadPath() As String
        Dim sAns As String
        Dim sErr As String = ""

        sAns = RegValue(RegistryHive.LocalMachine, "SOFTWARE", sAddonName, sErr)
        ReadPath = sAns
        If Not (sAns <> "") Then
            MessageBox.Show("This error occurred: " & sErr)
        End If
    End Function

    ' This Function reads values to the registry
    Public Function RegValue(ByVal Hive As RegistryHive, _
          ByVal Key As String, ByVal ValueName As String, _
          Optional ByRef ErrInfo As String = "") As String

        Dim objParent As RegistryKey
        Dim objSubkey As RegistryKey
        Dim sAns As String
        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.DynData
                objParent = Registry.DynData
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
        End Select

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If

        Catch ex As Exception
            ErrInfo = ex.Message
        Finally

            'if no error but value is empty, populate errinfo
            If ErrInfo = "" And sAns = "" Then
                ErrInfo = _
                   "No value found for requested registry key"
            End If
        End Try
        Return sAns
    End Function

    ' This Function writes values to the registry
    Public Function WriteToRegistry(ByVal _
    ParentKeyHive As RegistryHive, _
    ByVal SubKeyName As String, _
    ByVal ValueName As String, _
    ByVal Value As Object) As Boolean

        Dim objSubKey As RegistryKey
        Dim sException As String
        Dim objParentKey As RegistryKey
        Dim bAns As Boolean

        Try
            Select Case ParentKeyHive
                Case RegistryHive.ClassesRoot
                    objParentKey = Registry.ClassesRoot
                Case RegistryHive.CurrentConfig
                    objParentKey = Registry.CurrentConfig
                Case RegistryHive.CurrentUser
                    objParentKey = Registry.CurrentUser
                Case RegistryHive.DynData
                    objParentKey = Registry.DynData
                Case RegistryHive.LocalMachine
                    objParentKey = Registry.LocalMachine
                Case RegistryHive.PerformanceData
                    objParentKey = Registry.PerformanceData
                Case RegistryHive.Users
                    objParentKey = Registry.Users
            End Select

            'Open 
            objSubKey = objParentKey.OpenSubKey(SubKeyName, True)
            'create if doesn't exist
            If objSubKey Is Nothing Then
                objSubKey = objParentKey.CreateSubKey(SubKeyName)
            End If


            objSubKey.SetValue(ValueName, Value)
            bAns = True
        Catch ex As Exception
            bAns = False

        End Try

        Return True

    End Function

    ' This function extracts the given add-on into the path specified
    Private Sub ExtractFile(ByVal path As String)

        Dim tabla_dll As New ArrayList
        tabla_dll.Add(sAddonName & ".exe")
        ' tabla_dll.Add("slnGSEDOC_SAP.xml")

        tabla_dll.Add("itextsharp.dll")
        tabla_dll.Add("itextsharp.licensekey.dll")
        tabla_dll.Add("itextsharp.pdfa.dll")
        tabla_dll.Add("itextsharp.xtra.dll")

        tabla_dll.Add("Spire.License.dll")
        tabla_dll.Add("Spire.Pdf.dll")

        tabla_dll.Add("Entidades.dll")
        tabla_dll.Add("frmAcercaDe.srf")
        tabla_dll.Add("frmConfClave.srf")
        tabla_dll.Add("frmConfMenu.srf")
        tabla_dll.Add("frmParametrosAddon.srf")
        tabla_dll.Add("frmProxy.srf")

        tabla_dll.Add("frmConsultaOrdenes.srf")
        tabla_dll.Add("frmDocumento.srf")
        tabla_dll.Add("frmDocumentoNC.srf")
        tabla_dll.Add("frmDocumentoRE.srf")
        tabla_dll.Add("frmDocumentosEnviados.srf")
        tabla_dll.Add("frmDocumentosIntegrados.srf")
        tabla_dll.Add("frmDocumentosRecibidos.srf")
        tabla_dll.Add("frmLogEmision.srf")
        tabla_dll.Add("frmMapeo.srf")
        tabla_dll.Add("frmParametrosRecepcion.srf")
        tabla_dll.Add("frmSubirArchivo.srf")
        tabla_dll.Add("Functions.dll")
        tabla_dll.Add("imagen.jpg")
        'tabla_dll.Add("logo1.png")
        tabla_dll.Add("logo11.png")
        tabla_dll.Add("LogoSS.png")
        tabla_dll.Add("Negocio.dll")
        'tabla_dll.Add("Param.srf")
        tabla_dll.Add("Utilitario.dll")

        tabla_dll.Add("frmProcesoLote.srf")
        tabla_dll.Add("frmImpresionPorBloque.srf")
        tabla_dll.Add("frmValidarUsuario.srf")
        tabla_dll.Add("frmProcesoLoteManamer.srf")

        tabla_dll.Add("frmDocumentosRecibidosXML.srf")
        tabla_dll.Add("frmDocumentoXML.srf")
        tabla_dll.Add("frmDocumentoNCXML.srf")
        tabla_dll.Add("frmDocumentoREXML.srf")
        tabla_dll.Add("frmProcesoLoteXML.srf")

        tabla_dll.Add("Newtonsoft.Json.dll")
        tabla_dll.Add("frmProcesoLoteC.srf")
        tabla_dll.Add("frmListaAEnviar.srf")

        Dim AddonExeFile As IO.FileStream
        Dim thisExe As System.Reflection.Assembly
        Dim sTargetPath, sSourcePath, var_aux As String
        Dim largo As Integer
        Dim i, k As Integer
        Dim file As System.IO.Stream
        Dim buffer() As Byte
        thisExe = System.Reflection.Assembly.GetExecutingAssembly()

        'SI SE AGREGA UN FORMULARIO O EXEC CAMBIAR EN NUMERO
        For i = 0 To tabla_dll.Count - 1
            file = Nothing
            sTargetPath = path & "/" & tabla_dll(i)
            largo = Len(sTargetPath)
            k = (largo + 1) - 4
            var_aux = Mid(sTargetPath, k, 4)
            sSourcePath = Replace(sTargetPath, var_aux, ".tmp")
            file = thisExe.GetManifestResourceStream("AddOnInstaller." & tabla_dll(i))
            If IO.File.Exists(sSourcePath) Then
                IO.File.Delete(sSourcePath)
            End If
            AddonExeFile = IO.File.Create(sSourcePath)
            ReDim buffer(file.Length)
            file.Read(buffer, 0, file.Length)
            AddonExeFile.Write(buffer, 0, file.Length)
            AddonExeFile.Close()

            If IO.File.Exists(sTargetPath) Then
                IO.File.Delete(sTargetPath)
            End If
            ' Change file extension to exe
            IO.File.Move(sSourcePath, sTargetPath)
        Next

        'sTargetPath = path
        ''****************************** 
        'Shell(sTargetPath & "\CreaUsuarioAddon.exe", AppWinStyle.NormalFocus, True)
        ''****************************** 

    End Sub

    ' This procedure deletes the addon files
    Private Sub UnInstall()
        Dim path As String
        path = ReadPath() ' Reads the addon path from the registry
        If path <> "" Then
            ''-------------------------------------------------------
            'Try
            '    ' Delete the addon EXE file
            '    If IO.File.Exists(path & "\" & sAddonName & ".exe") Then
            '        IO.File.Delete(path & "\" & sAddonName & ".exe")
            '        MessageBox.Show(path & "\" & sAddonName & ".exe was deleted")
            '    Else
            '        MessageBox.Show(path & "\" & sAddonName & ".exe was not found")
            '    End If
            'Catch
            '    MessageBox.Show(path & " - ERROR UNINSTALLING")
            'End Try
            ''-------------------------------------------------------
            'Try
            '    ' Delete the dll1 file
            '    If IO.File.Exists(path & "\Interop.SAPbobsCOM.dll") Then
            '        IO.File.Delete(path & "\Interop.SAPbobsCOM.dll")
            '        MessageBox.Show(path & "\Interop.SAPbobsCOM.dll was deleted")
            '    Else
            '        MessageBox.Show(path & "\Interop.SAPbobsCOM.dll was not found")
            '    End If
            'Catch
            '    MessageBox.Show(path & " - ERROR UNINSTALLING")
            'End Try
            ''-------------------------------------------------------
            'Try
            '    ' Delete the dll2 file
            '    If IO.File.Exists(path & "\Interop.SAPbouiCOM.dll") Then
            '        IO.File.Delete(path & "\Interop.SAPbouiCOM.dll")
            '        MessageBox.Show(path & "\Interop.SAPbouiCOM.dll was deleted")
            '    Else
            '        MessageBox.Show(path & "\Interop.SAPbouiCOM.dll was not found")
            '    End If
            'Catch
            '    MessageBox.Show(path & " - ERROR UNINSTALLING")
            'End Try
            ''-------------------------------------------------------
            'Try
            '    ' Delete the dll3 file
            '    If IO.File.Exists(path & "\Interop.Scripting.dll") Then
            '        IO.File.Delete(path & "\Interop.Scripting.dll")
            '        MessageBox.Show(path & "\Interop.Scripting.dll was deleted")
            '    Else
            '        MessageBox.Show(path & "\Interop.Scripting.dll was not found")
            '    End If
            'Catch
            '    MessageBox.Show(path & " - ERROR UNINSTALLING")
            'End Try
            ''-------------------------------------------------------
            'Try
            '    ' Delete the screen form file
            '    If IO.File.Exists(path & "\SetupImpresionele.exe") Then
            '        IO.File.Delete(path & "\SetupImpresionele.exe")
            '        MessageBox.Show(path & "\SetupImpresionele.exe was deleted")
            '    Else
            '        MessageBox.Show(path & "\SetupImpresionele.exe was not found")
            '    End If
            'Catch
            '    MessageBox.Show(path & " - ERROR UNINSTALLING")
            'End Try
            ''-------------------------------------------------------
            'Try
            '    ' Delete the screen form file
            '    If IO.File.Exists(path & "\CreaUsuarioAddon.exe") Then
            '        IO.File.Delete(path & "\CreaUsuarioAddon.exe")
            '        MessageBox.Show(path & "\CreaUsuarioAddon.exe was deleted")
            '    Else
            '        MessageBox.Show(path & "\CreaUsuarioAddon.exe was not found")
            '    End If
            'Catch
            '    MessageBox.Show(path & " - ERROR UNINSTALLING")
            'End Try
            ''-------------------------------------------------------

            Dim DirFullInfo As New System.IO.DirectoryInfo(path)
            Dim sfs As System.IO.FileInfo()
            Dim sf As System.IO.FileInfo
            Dim i As Integer
            sfs = DirFullInfo.GetFiles
            For i = 0 To sfs.Length - 1
                sf = sfs.GetValue(i)
                If sf.Name <> sInstallName Then
                    Try
                        ' Delete the screen form file
                        If IO.File.Exists(path & "\" & sf.Name) And sf.Extension.Trim.ToUpper <> ".DAT" Then
                            IO.File.Delete(path & "\" & sf.Name)
                            ' MessageBox.Show(path & "\" & sf.Name & " was deleted")
                        ElseIf sf.Extension.Trim.ToUpper <> ".DAT" Then
                            MessageBox.Show(path & "\" & sf.Name & " was not found")
                        End If
                    Catch
                        MessageBox.Show(path & "\" & sf.Name & " - ERROR UNINSTALLING")
                    End Try
                End If
            Next

            MessageBox.Show("( " & Format(i, "###0") & " )" & "  Additional Files Deleted from" & vbCrLf & path)
            '-------------------------------------------------------
        Else
            MessageBox.Show("Path not found")
        End If
        ' Terminate the application
        GC.Collect()
        End
    End Sub

    ' This procedure copies the addon exe file to the installation folder        
    Private Function Install() As Boolean
        Dim resp As Boolean = True
        Try
            Environment.CurrentDirectory = strDll ' For Dll function calls will work

            If chkDefaultFolder.Checked = False Then ' Change the installation folder
                SetAddOnFolder(txtDest.Text)
                strDest = txtDest.Text
            End If

            If Not (IO.Directory.Exists(strDest)) Then
                IO.Directory.CreateDirectory(strDest) ' Create installation folder
            End If

            FileWatcher.Path = strDest
            FileWatcher.EnableRaisingEvents = True

            ExtractFile(strDest) ' Extract add-on to installation folder

            While bFileCreated = False
                Application.DoEvents()
                'Don't continue running until the file is copied...
            End While

            If chkRestart.Checked Then
                RestartNeeded() ' Inform SBO the restart is needed
            End If
            EndInstall() ' Inform SBO the installation ended
            'Write installation Folder to registry
            Dim bAns As Boolean

            'WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", "path", "c:\folder")
            bAns = WriteToRegistry(RegistryHive.LocalMachine, "SOFTWARE", sAddonName, strDest)
            MessageBox.Show("Finished Installing", "Installation ended", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Windows.Forms.Application.Exit() ' Exit the installer
            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Addon Installer")
            Return False
        End Try
    End Function

#End Region

#Region "Events"
    Private Sub frmInstall_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.lblHeadLine.Text = "Instalación - AddOn"
            'Me.Label1.Text = "Usted se dispone a instalar el AddOn: " & sAddonName
            Me.Label2.Text = "Carpeta de Instalación de AddOn en SBO"
            Me.chkDefaultFolder.Text = "Usar carpeta entregada por SBO"
            Me.chkRestart.Text = "Reiniciar"
            Me.cmdInstall.Text = "Instalar"
            'Dim strAppPath As String
            ' The command line parameters, seperated by '|' will be broken to this array
            Dim strCmdLineElements(2) As String
            Dim strCmdLine As String ' The whole command line
            Dim NumOfParams As Integer 'The number of parameters in the command line (should be 2)
            NumOfParams = Environment.GetCommandLineArgs.Length

            If NumOfParams = 2 Then
                strCmdLine = Environment.GetCommandLineArgs.GetValue(1)
                If strCmdLine.ToUpper = "/U" Then
                    UnInstall()
                End If
                strCmdLineElements = strCmdLine.Split("|")

                ' Get Install destination Folder
                strDest = strCmdLineElements.GetValue(0)
                txtDest.Text = strDest

                ' Get the "AddOnInstallAPI.dll" path
                strDll = strCmdLineElements.GetValue(1)
                strDll = strDll.Remove((strDll.Length - 19), 19) ' Only the path is needed
            Else
                MessageBox.Show("This installer must be run from Sap Business One", _
                                "Incorrect installation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Windows.Forms.Application.Exit()
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmdInstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInstall.Click

        If Install() = True Then
            'Create_SQL_User()
        End If

    End Sub

    Private Sub Create_SQL_User()
        Shell(Application.StartupPath & "\CreaUsuarioAddon.exe", AppWinStyle.MaximizedFocus)
    End Sub

    Private Sub chkDefaultFolder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDefaultFolder.CheckedChanged
        txtDest.Enabled = Not (chkDefaultFolder.Checked)
    End Sub

    ' This event happens when the addon exe file is renamed to exe extention
    Private Sub FileWatcher_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles FileWatcher.Renamed
        bFileCreated = True
        FileWatcher.EnableRaisingEvents = False
    End Sub

    Public Sub ShowError(ByVal ex As Exception)
        MsgBox(ex.Message & vbNewLine & "Source:" & ex.StackTrace, MsgBoxStyle.Information, "Addon Installer")
    End Sub
#End Region

End Class
