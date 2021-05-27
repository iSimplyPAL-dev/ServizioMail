
Imports System.ServiceProcess
Imports log4net
Imports log4net.Config
Imports System.IO
'''---------------------------------------------------------------------------------------------------
'''Autore.......: Antonello Lo Bianco
'''Data ........: 02/05/2014
'''Descrizione..: Creazione della classe di gestione del servizio.
'''Progetto ....: ReminderServiceMail
'''Revisioni....: Indicare qui di seguito le modifiche apportate
'''---------------------------------+-------------------------+---------------------------------------
'''Nome e Cognome                   Data                      Descrizione revisione
'''---------------------------------+-------------------------+---------------------------------------
'''Antonello Lo Bianco              02/05/2014                Creazione classe.
'''---------------------------------+-------------------------+---------------------------------------
Public Class TegServiceEngine
    Inherits System.ServiceProcess.ServiceBase

    Private WithEvents objWorker As Worker
    Private WithEvents StopTimer As System.Timers.Timer

    Public Sub StartService(ByVal Debug As Boolean)
        Try

            Dim pathfileinfo As String = My.Settings.pathfileconflog4net
            Dim fileconfiglog4net As FileInfo = New FileInfo(pathfileinfo)

            XmlConfigurator.ConfigureAndWatch(fileconfiglog4net)

            objWorker_WorkState(String.Format("{0} v.{1}", Application.ExecutablePath, My.Application.Info.Version), EventLogEntryType.Information)
            objWorker_WorkState("Starting service...", EventLogEntryType.Information)

            'Imposta la directory corrente a quella dove si trova il programma
            System.IO.Directory.SetCurrentDirectory(GetMyDir)

            'Crea un nuovo thread nel quale far girare il ciclo continuo della classe Worker
            objWorker = New Worker()
            Dim th As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf objWorker.DoWork))
            th.Name = objWorker.InstanceName
            th.Start()
            'Resta in attesa un momento prima di verificare che la classe Worker sia in esecuzione correttamente
            System.Threading.Thread.Sleep(500)
            'Verifica lo stato della classe Worker, per capire se il servizio sia attivo oppure no
            If objWorker IsNot Nothing AndAlso objWorker.IsRunning Then
                objWorker_WorkState("Worker class loaded succesfully.", EventLogEntryType.Information)
            ElseIf objWorker IsNot Nothing Then
                objWorker_WorkState("Worker class is NOT running !", EventLogEntryType.Warning)
            Else
                objWorker_WorkState("Worker class load failed !", EventLogEntryType.Warning)
            End If
            If objWorker IsNot Nothing Then objWorker_WorkState("Service started.", EventLogEntryType.Information)
            If Debug Then
                Do
                    System.Threading.Thread.Sleep(100)
                    System.Windows.Forms.Application.DoEvents()
                Loop Until objWorker Is Nothing
            End If
        Catch ex As Exception
            objWorker_WorkState(String.Format("Error starting service... {0}", ex.Message), EventLogEntryType.Error)
        End Try
    End Sub

    Protected Overrides Sub OnContinue()
        objWorker_WorkState("Resuming service...", EventLogEntryType.Information)
        MyBase.OnContinue()
        objWorker.ResumeWork()
        objWorker_WorkState("Service resumed.", EventLogEntryType.Information)
    End Sub

    Protected Overrides Sub OnPause()
        objWorker_WorkState("Pausing service...", EventLogEntryType.Information)
        MyBase.OnPause()
        objWorker.PauseWork()
        objWorker_WorkState("Service paused.", EventLogEntryType.Information)
    End Sub

    Protected Overrides Sub OnShutdown()
        Try
            objWorker_WorkState("Shutdown in progress...", EventLogEntryType.Information)
            'Il pc è in fase di spegnimento, quindi il servizio deve essere arrestato
            'StopTimer = New System.Timers.Timer(10)
            'StopTimer.AutoReset = False
            'StopTimer.Start()
            Me.OnStop()
        Catch ex As Exception
            'Errore non tracciable ?
            objWorker_WorkState(String.Format("Shutdown error... {0}", ex.Message), EventLogEntryType.Error)
        End Try
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Inserire qui il codice necessario per avviare il proprio servizio. Il metodo deve effettuare le impostazioni necessarie per il funzionamento del servizio.
        StartService(False)
    End Sub

    Protected Overrides Sub OnStop()
        ' Inserire qui il codice delle procedure di chiusura necessarie per arrestare il proprio servizio.
        Try
            objWorker_WorkState(String.Format("{0} v.{1} - {2}", Application.ExecutablePath, My.Application.Info.Version, ProjectInstaller.EVLOG_SOURCE), EventLogEntryType.Information)
            objWorker_WorkState("Stopping service...", EventLogEntryType.Information)

            If StopTimer IsNot Nothing Then
                StopTimer.Stop()
                StopTimer.Dispose()
            End If
            StopTimer = Nothing

            'Segnala alla classe Worker di interrompere il ciclo continuo
            If objWorker IsNot Nothing Then
                'Resta in attesa fino a quando la classe Worker non segnala di aver terminato il lavoro
                objWorker_WorkState("Waiting for worker class to stop...", EventLogEntryType.Information)
                objWorker.StopWork()
                'Resta in attesa al massimo 10 secondi e poi distrugge la classe Worker
                Dim nCounter As Integer = 1
                Do
                    System.Threading.Thread.Sleep(1000)
                    System.Windows.Forms.Application.DoEvents()
                    nCounter += 1
                Loop While objWorker.IsRunning And nCounter < 10
                objWorker.Dispose()
                objWorker = Nothing
            Else
                objWorker_WorkState("Worker class was not loaded !", EventLogEntryType.Warning)
            End If
            objWorker_WorkState("Service stopped.", EventLogEntryType.Information)
        Catch ex As Exception
            objWorker_WorkState(String.Format("Error stopping service... {0}", ex.Message), EventLogEntryType.Error)
        End Try
    End Sub

    Private Function FormatMessage(ByVal Message As String) As String
        'Formatta eventuali caratteri speciali contenuti nel messaggio in modo da renderli "visibili"
        If Message.Contains(ChrW(0)) Then Message = Replace(Message, ChrW(0), "<NUL>")
        If Message.Contains(ChrW(1)) Then Message = Replace(Message, ChrW(1), "<SOH>")
        If Message.Contains(ChrW(2)) Then Message = Replace(Message, ChrW(2), "<STX>")
        If Message.Contains(ChrW(3)) Then Message = Replace(Message, ChrW(3), "<ETX>")
        If Message.Contains(ChrW(4)) Then Message = Replace(Message, ChrW(4), "<EOT>")
        If Message.Contains(ChrW(5)) Then Message = Replace(Message, ChrW(5), "<ENQ>")
        If Message.Contains(ChrW(6)) Then Message = Replace(Message, ChrW(6), "<ACK>")
        If Message.Contains(ChrW(7)) Then Message = Replace(Message, ChrW(7), "<BELL>")
        If Message.Contains(ChrW(8)) Then Message = Replace(Message, ChrW(8), "<BS>")
        If Message.Contains(ChrW(9)) Then Message = Replace(Message, ChrW(9), "<TAB>")
        'If Message.Contains(ChrW(10)) Then Message = Replace(Message, ChrW(10), "<LF>")
        If Message.Contains(ChrW(11)) Then Message = Replace(Message, ChrW(11), "<VT>")
        If Message.Contains(ChrW(12)) Then Message = Replace(Message, ChrW(12), "<FF>")
        'If Message.Contains(ChrW(13)) Then Message = Replace(Message, ChrW(13), "<CR>")
        If Message.Contains(ChrW(14)) Then Message = Replace(Message, ChrW(14), "<SO>")
        If Message.Contains(ChrW(15)) Then Message = Replace(Message, ChrW(15), "<SI>")
        If Message.Contains(ChrW(16)) Then Message = Replace(Message, ChrW(16), "<DLE>")
        If Message.Contains(ChrW(17)) Then Message = Replace(Message, ChrW(17), "<DC1>")
        If Message.Contains(ChrW(18)) Then Message = Replace(Message, ChrW(18), "<DC2>")
        If Message.Contains(ChrW(19)) Then Message = Replace(Message, ChrW(19), "<DC3>")
        If Message.Contains(ChrW(20)) Then Message = Replace(Message, ChrW(20), "<DC4>")
        If Message.Contains(ChrW(21)) Then Message = Replace(Message, ChrW(21), "<NAK>")
        If Message.Contains(ChrW(22)) Then Message = Replace(Message, ChrW(22), "<SYN>")
        If Message.Contains(ChrW(23)) Then Message = Replace(Message, ChrW(23), "<ETB>")
        If Message.Contains(ChrW(24)) Then Message = Replace(Message, ChrW(24), "<CAN>")
        If Message.Contains(ChrW(25)) Then Message = Replace(Message, ChrW(25), "<EM>")
        If Message.Contains(ChrW(26)) Then Message = Replace(Message, ChrW(26), "<SUB>")
        If Message.Contains(ChrW(27)) Then Message = Replace(Message, ChrW(27), "<ESC>")
        If Message.Contains(ChrW(28)) Then Message = Replace(Message, ChrW(28), "<FS>")
        If Message.Contains(ChrW(29)) Then Message = Replace(Message, ChrW(29), "<GS>")
        If Message.Contains(ChrW(30)) Then Message = Replace(Message, ChrW(30), "<RS>")
        If Message.Contains(ChrW(31)) Then Message = Replace(Message, ChrW(31), "<US>")
        Return Message
    End Function

    Function GetMyDir() As String
        Return Application.StartupPath
    End Function

    Private Sub StopTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles StopTimer.Elapsed
        Try
            StopTimer.Stop()
            Me.Stop()
        Catch ex As Exception
            objWorker_WorkState(String.Format("Error stopping service automatically... {0}", ex.Message), EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub objWorker_WorkFailed(ByVal sender As Object, ByVal e As System.EventArgs) Handles objWorker.WorkFailed
        Try
            objWorker_WorkState("Service failed !", EventLogEntryType.Error)
            'La classe worker ha segnalato un errore in fase di avvio, quindi il servizio deve essere arrestato
            StopTimer = New System.Timers.Timer(1000)
            StopTimer.AutoReset = False
            StopTimer.Start()
        Catch ex As Exception
            objWorker_WorkState(String.Format("Error tracing service... {0}", ex.Message), EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub objWorker_WorkState(ByVal Message As String, ByVal State As System.Diagnostics.EventLogEntryType) Handles objWorker.WorkState
        Try
            If objWorker Is Nothing OrElse objWorker.WriteOnEventLog OrElse State <> EventLogEntryType.Information Then
                EventLog1.WriteEntry(FormatMessage(Message), State)
            End If
        Catch ex As Exception
            WriteEventLog(Me.ServiceName, ex.Message & ": " & FormatMessage(Message), EventLogEntryType.Warning)
        End Try
    End Sub

    Public Shared Sub WriteEventLog(ByVal Source As String, ByVal Message As String, ByVal State As System.Diagnostics.EventLogEntryType)
        Try
            System.Diagnostics.EventLog.WriteEntry(Source, Message, State)
        Catch ex2 As Exception
            My.Application.Log.WriteException(ex2, TraceEventType.Warning, Message)
        End Try
    End Sub

    Public Sub New()
        ' Chiamata richiesta da Progettazione Windows Form.
        InitializeComponent()

        ' Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent().
        Me.ServiceName = ProjectInstaller.EVLOG_SOURCE
        Me.AutoLog = False
        Me.CanShutdown = True
        Me.CanPauseAndContinue = True

        'Set source and log section for the EventLog object
        EventLog1.Source = ProjectInstaller.EVLOG_SOURCE
        EventLog1.Log = ProjectInstaller.EVLOG_LOG
        'EventLog1.Log = EventLog.LogNameFromSourceName(EventLog1.Source, ".")
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
