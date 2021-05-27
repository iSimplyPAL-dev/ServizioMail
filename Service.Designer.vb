Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TegServiceEngine
  Inherits System.ServiceProcess.ServiceBase

  'UserService esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

  ' Il punto di ingresso principale del processo
  <MTAThread()> _
  <System.Diagnostics.DebuggerNonUserCode()> _
  Shared Sub Main(ByVal args() As String)
    Dim bStartService As Boolean = False

    Try
      'Verifica se l'applicazione sia già registrata nel registro eventi di sistema
      If System.Diagnostics.EventLog.SourceExists(ProjectInstaller.EVLOG_SOURCE) Then
        'Rimuove l'applicazione dal registro eventi in cui si trova, in modo che alla prima scrittura venga associata al registro Application
        EventLog.DeleteEventSource(ProjectInstaller.EVLOG_SOURCE, ".")
      End If
    Catch ex As Exception
      TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("ALARM: {0}", ex.ToString), EventLogEntryType.Warning)
    End Try

    If args.Length > 0 Then
      'Traccia nel log i parametri ricevuti sulla command-line
      Dim sParams As String = ""
      For i As Integer = 0 To args.Length - 1
        sParams &= args(i) & "; "
      Next
      TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("Startup Parameters: {0}", sParams), EventLogEntryType.Information)
    End If

    'Verifica se sono stati passati dei parametri e se il primo parametro inizi con un trattino o cun una barra
    If args.Length > 0 AndAlso (args(0).Substring(0, 1) = "-" OrElse args(0).Substring(0, 1) = "/") Then
      Select Case args(0).ToLower.Substring(1)
        Case "i", "install"
          Try
            'Prima di procedere con la registrazione del servizio, verifica che non sia già registrato in qualche registro eventi di sistema
            If System.Diagnostics.EventLog.SourceExists(ProjectInstaller.EVLOG_SOURCE) Then
              'Rimuove l'applicazione dal registro eventi in cui si trova, in modo che alla prima scrittura venga associata al registro Application
              EventLog.DeleteEventSource(ProjectInstaller.EVLOG_SOURCE)
            End If
          Catch ex As Exception
            TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("ALARM: {0}", ex.ToString), EventLogEntryType.Warning)
          End Try

          'Richiama la classe che permette la registrazione dell'applicazione come servizio Windows
          If SelfInstaller.InstallMe() Then
            Try
              'Verifica in quale registro eventi è stato associato il programma dopo la registrazione
              If System.Diagnostics.EventLog.SourceExists(ProjectInstaller.EVLOG_SOURCE) Then
                Dim sLogName As String = EventLog.LogNameFromSourceName(ProjectInstaller.EVLOG_SOURCE, ".")
                If sLogName <> ProjectInstaller.EVLOG_LOG Then
                  EventLog.DeleteEventSource(ProjectInstaller.EVLOG_SOURCE)
                  If Not System.Diagnostics.EventLog.SourceExists(ProjectInstaller.EVLOG_SOURCE) Then
                    EventLog.CreateEventSource(ProjectInstaller.EVLOG_SOURCE, ProjectInstaller.EVLOG_LOG)
                  End If
                End If
              Else
                EventLog.CreateEventSource(ProjectInstaller.EVLOG_SOURCE, ProjectInstaller.EVLOG_LOG)
              End If
            Catch ex As Exception
              TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("ALARM: {0}", ex.ToString), EventLogEntryType.Warning)
            End Try
          End If
        Case "u", "uninstall"
          If SelfInstaller.UninstallMe() Then
            Try
              If System.Diagnostics.EventLog.SourceExists(ProjectInstaller.EVLOG_SOURCE) Then
                EventLog.DeleteEventSource(ProjectInstaller.EVLOG_SOURCE)
              End If
            Catch ex As Exception
              TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("ALARM: {0}", ex.ToString), EventLogEntryType.Warning)
            End Try
          End If
        Case "c", "console"
          TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, "Cannot run service as a console application.", EventLogEntryType.Warning)
        Case Else
          TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("Unknown startup mode: {0}", args(0)), EventLogEntryType.Error)
      End Select
    ElseIf args.Length = 0 Then
      'Avvia il programma come servizio solo se non è stato passato nessun parametro
      bStartService = True
    Else
      TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, "Unknown parameters. Service aborted.", EventLogEntryType.Error)
      Throw New ApplicationException("ABORT!")
    End If

    If bStartService Then
      'Verifica se ci sia un debugger attivo, in tal caso presume di essere in debug con l'ambiente di sviluppo
      If System.Diagnostics.Debugger.IsAttached Then
        TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("Entering debug mode for {0}", ProjectInstaller.EVLOG_SOURCE), EventLogEntryType.Information)
        Dim test As New TegServiceEngine()
        test.StartService(True)
      Else
        TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, String.Format("Entering running mode for {0}", ProjectInstaller.EVLOG_SOURCE), EventLogEntryType.Information)
        ' All'interno di uno stesso processo è possibile eseguire più servizi di Windows NT.
        ' Per aggiungere un servizio al processo, modificare la riga che segue in modo
        ' da creare un secondo oggetto servizio. Ad esempio,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ServicesToRun = New System.ServiceProcess.ServiceBase() {New TegServiceEngine}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
      End If
    End If
  End Sub

  'Richiesto da Progettazione componenti
  Private components As System.ComponentModel.IContainer

  ' NOTA: la procedura che segue è richiesta da Progettazione componenti
  ' Può essere modificata in Progettazione componenti.  
  ' Non modificarla nell'editor del codice.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.EventLog1 = New System.Diagnostics.EventLog
    CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).BeginInit()
    '
        'ReminderServiceMail
    '
        Me.ServiceName = "ReminderServiceMail"
    CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).EndInit()

  End Sub
  Friend WithEvents EventLog1 As System.Diagnostics.EventLog

End Class
