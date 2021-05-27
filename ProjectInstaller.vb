Imports System.ComponentModel
Imports System.Configuration.Install

Public Class ProjectInstaller
    Public Const EVLOG_DISPLAY As String = "Mail Reminder Service"
    Public Const EVLOG_SOURCE As String = "ReminderServiceMail"
  Public Const EVLOG_LOG As String = "Application"

  Public Sub New()
    MyBase.New()

    'Chiamata richiesta da Progettazione componenti.
    InitializeComponent()

    'Aggiungere il codice di inizializzazione dopo la chiamata a InitializeComponent
    Dim sSource As String = EVLOG_SOURCE
    Try
      If System.Diagnostics.EventLog.SourceExists(sSource) Then
        Dim sLogName As String = EventLog.LogNameFromSourceName(sSource, ".")
        If sLogName <> EVLOG_LOG Then
          EventLog.DeleteEventSource(sSource)
          If Not System.Diagnostics.EventLog.SourceExists(sSource) Then
            EventLog.CreateEventSource(sSource, EVLOG_LOG)
          End If
        End If
      Else
        If Not System.Diagnostics.EventLog.Exists(EVLOG_LOG) Then
          EventLog.CreateEventSource(sSource, EVLOG_LOG)
        End If
      End If

    Catch ex As ArgumentException
      If Not System.Diagnostics.EventLog.SourceExists(EVLOG_SOURCE) Then
        EventLog.CreateEventSource(EVLOG_SOURCE, EVLOG_LOG)
      End If
    Catch ex As Exception
      TegServiceEngine.WriteEventLog(ProjectInstaller.EVLOG_SOURCE, ex.ToString, EventLogEntryType.Error)
    End Try

  End Sub

  Private Sub ProjectInstaller_AfterInstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles Me.AfterInstall
    TegServiceEngine.WriteEventLog(EVLOG_SOURCE, String.Format("'{0}' successfully installed as service named '{1}'", Application.ExecutablePath, Me.ServiceInstaller1.ServiceName), EventLogEntryType.Information)
  End Sub

  Private Sub ProjectInstaller_AfterRollback(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles Me.AfterRollback
    TegServiceEngine.WriteEventLog(EVLOG_SOURCE, String.Format("'{0}' installation failed !", Application.ExecutablePath), EventLogEntryType.Error)
  End Sub

  Private Sub ProjectInstaller_AfterUninstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles Me.AfterUninstall
    TegServiceEngine.WriteEventLog(EVLOG_SOURCE, String.Format("'{0}' successfully uninstalled.", Me.ServiceInstaller1.ServiceName), EventLogEntryType.Information)
  End Sub

  Private Sub ProjectInstaller_BeforeInstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles Me.BeforeInstall
    Me.ServiceInstaller1.ServiceName = EVLOG_SOURCE
    Me.ServiceInstaller1.DisplayName = EVLOG_DISPLAY
  End Sub

  Private Sub ProjectInstaller_BeforeUninstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles Me.BeforeUninstall
    Me.ServiceInstaller1.ServiceName = EVLOG_SOURCE
    Me.ServiceInstaller1.DisplayName = EVLOG_DISPLAY
  End Sub
End Class
