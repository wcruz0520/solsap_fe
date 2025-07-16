<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    'Se puede modificar usando el Diseñador de componentes.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ServiceProcessInstaller1 = New System.ServiceProcess.ServiceProcessInstaller()
        Me.GSEDOC_Procesador_BETA = New System.ServiceProcess.ServiceInstaller()
        '
        'ServiceProcessInstaller1
        '
        Me.ServiceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessInstaller1.Password = Nothing
        Me.ServiceProcessInstaller1.Username = Nothing
        '
        'GSEDOC_Procesador_BETA
        '
        Me.GSEDOC_Procesador_BETA.Description = "Procesador de Facturas Integradas"
        Me.GSEDOC_Procesador_BETA.DisplayName = "GSEDOC_Procesador_BETA"
        Me.GSEDOC_Procesador_BETA.ServiceName = "GSEDOC_Procesador_BETA"
        Me.GSEDOC_Procesador_BETA.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstaller1, Me.GSEDOC_Procesador_BETA})

    End Sub
    Friend WithEvents ServiceProcessInstaller1 As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents GSEDOC_Procesador_BETA As System.ServiceProcess.ServiceInstaller

End Class
