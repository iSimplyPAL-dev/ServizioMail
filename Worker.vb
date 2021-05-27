Imports Teg.TegService
Imports Teg.TegService.Globals

Namespace TegService
    '''---------------------------------------------------------------------------------------------------
    '''Autore.......: Antonello Lo Bianco
    '''Data ........: 02/05/2014
    '''Descrizione..: Servizio Windows per importazione dati da WorkingLab.
    '''Progetto ....: ReminderServiceMail
    '''Revisioni....: Indicare qui di seguito le modifiche apportate
    '''---------------------------------+-------------------------+---------------------------------------
    '''Nome e Cognome                   Data                      Descrizione revisione
    '''---------------------------------+-------------------------+---------------------------------------
    '''Antonello Lo Bianco                  02/05/2014                Creazione classe.
    '''---------------------------------+-------------------------+---------------------------------------
    Public Class Worker
        Implements IDisposable

        Public Event WorkFailed(ByVal sender As Object, ByVal e As System.EventArgs)
        Public Event WorkState(ByVal Message As String, ByVal State As System.Diagnostics.EventLogEntryType)

        Private moWorkThread As System.Threading.Thread = Nothing
        Private msInstanceName As String
        Private mbIsRunning As Boolean
        Private WithEvents TimerLavoro As System.Timers.Timer
        Private moServiceWorker As ServiceWorker

        Private mbIsPaused As Boolean
        Private LogQueue As Generic.Queue(Of TegServiceTraceEventArgs)
        Private LastTraceMessage As TegServiceTraceEventArgs
        Private SameMessageCounter As Integer
        Private StopArchiving As Boolean

        Private TimerInterval As Integer = 1000
        Private TraceLimit As Integer = 100
        Friend WriteOnEventLog As Boolean = True

        Public ReadOnly Property InstanceName() As String
            Get
                Return msInstanceName
            End Get
        End Property

        Public ReadOnly Property IsRunning() As Boolean
            Get
                Return mbIsRunning
            End Get
        End Property

        Public Sub DoWork()
            Dim sMess As String = ""

            'Attende un attimo prima di avviare il lavoro
            System.Threading.Thread.Sleep(100)
            moWorkThread = System.Threading.Thread.CurrentThread

            mbIsRunning = True

            'Segnala l'inizio attività
            TraceLog("Loading...")

            'Crea l'istanza della classe applicativa
            moServiceWorker = New ServiceWorker
            If moServiceWorker IsNot Nothing Then
                Try
                    'Aggancia i gestori di evento per gli eventi della classe
                    AddHandler moServiceWorker.Trace, AddressOf Me.ArchiveMessage
                    AddHandler moServiceWorker.StopService, AddressOf Me.HandleStopRequest
                    AddHandler moServiceWorker.StateChanged, AddressOf Me.HandleStateChanges
                    'Inizializza la classe, passando i parametri ricevuti in fase di installazione/configurazione del servizio
                    TraceLog("Initializing...")
                    If moServiceWorker.Initialize("") Then
                        'Avvia un timer che permette di eseguire le attività a tempo
                        TimerLavoro = New System.Timers.Timer(TimerInterval)
                        TimerLavoro.AutoReset = False
                        TimerLavoro.Start()
                        TraceLog("Work Started.")
                    Else
                        TraceError("Error during initialization.")
                        mbIsRunning = False
                        RaiseEvent WorkFailed(Me, New EventArgs)
                    End If
                Catch ex As Exception
                    TraceError(String.Format("ERROR: {0}", ex.ToString))
                End Try
            Else
                TraceError("Error loading application class.")
                mbIsRunning = False
                RaiseEvent WorkFailed(Me, New EventArgs)
            End If
        End Sub

        Public Sub GetWorkStatus()
            Try
                'Segnala alla classe di trasmettere lo stato delle attività
                If moServiceWorker IsNot Nothing Then moServiceWorker.JobStatus()
            Catch ex As Exception
                TraceError(String.Format("ERROR: {0}", ex.ToString))
            End Try
        End Sub

        Public Sub PauseWork()
            mbIsPaused = True
            'Ferma il timer del ciclo di lavoro
            If TimerLavoro IsNot Nothing Then TimerLavoro.Stop()
            Try
                'Segnala alla classe di sospendere le attività
                If moServiceWorker IsNot Nothing Then moServiceWorker.PauseJob()
            Catch ex As Exception
                TraceError(String.Format("ERROR: {0}", ex.ToString))
            End Try
        End Sub

        Public Sub ResumeWork()
            Try
                'Segnala alla classe di riprendere le attività
                If moServiceWorker IsNot Nothing Then moServiceWorker.ResumeJob()
            Catch ex As Exception
                TraceError(String.Format("ERROR: {0}", ex.ToString))
            End Try
            mbIsPaused = False
            'Riattiva il timer del ciclo di lavoro
            If TimerLavoro IsNot Nothing Then TimerLavoro.Start()
        End Sub

        Public Sub StopWork()
            'Ferma il timer del ciclo di lavoro
            If TimerLavoro IsNot Nothing Then TimerLavoro.Stop()

            If moServiceWorker IsNot Nothing Then
                'Segnala alla classe di terminare le attività
                Try
                    TraceLog("Stopping...")
                    moServiceWorker.Terminate()
                    TraceLog("Stopped.")
                Catch ex As Exception
                    TraceError(String.Format("ERROR: {0}", ex.ToString))
                End Try
                'Rimuove i gestori di evento
                RemoveHandler moServiceWorker.Trace, AddressOf Me.ArchiveMessage
                RemoveHandler moServiceWorker.StopService, AddressOf Me.HandleStopRequest
                RemoveHandler moServiceWorker.StateChanged, AddressOf Me.HandleStateChanges
            End If

            'Salva su file la coda messaggi
            SaveMessageQueue()

            mbIsRunning = False
            If Not moWorkThread Is Nothing Then
                TraceLog("Closing Working Thread...")
                If Not moWorkThread.Join(3000) Then
                    TraceLog("Working Thread Aborted.")
                    moWorkThread.Abort()
                Else
                    TraceLog("Working Thread Closed.")
                End If
            Else
                TraceLog("Work Finished.")
            End If
        End Sub

        Private Sub HandleStateChanges(ByVal sender As Object, ByVal e As TegServiceStateChangedEventArgs)
            TraceLog(String.Format("Application state changed: {0} to {1}", e.PreviuosState, e.CurrentState), EventLogEntryType.Information)
        End Sub

        Private Sub HandleStopRequest(ByVal sender As Object, ByVal e As EventArgs)
            PauseWork()
            TraceLog("Application class request to stop the service.", EventLogEntryType.Error)
            RaiseEvent WorkFailed(sender, e)
        End Sub

        Public Function GetMessageQueue() As String
            Dim sRetVal As String = ""
            Try
                'Verifica che la coda messaggi sia stata creata e che contenga dei messaggi
                If LogQueue IsNot Nothing AndAlso LogQueue.Count > 0 Then
                    'Legge e rimuove il primo messaggio presente nella coda messaggi
                    Using e As TegServiceTraceEventArgs = LogQueue.Dequeue()
                        sRetVal = e.ToString
                        If e.Source = "" Then
                            sRetVal &= ProjectInstaller.EVLOG_SOURCE
                        End If
                    End Using
                End If
            Catch ex As Exception
                TraceError(String.Format("ERROR: {0}", ex.ToString))
                sRetVal = ""
            End Try
            If sRetVal <> "" Then
                sRetVal = "<TRACE>" & sRetVal & "</TRACE>"
            End If

            Return sRetVal
        End Function

        Private Sub ArchiveMessage(ByVal sender As Object, ByVal e As TegServiceTraceEventArgs)
            'Verifica se registrare il messaggio ricevuto nella coda messaggi oppure no
            If Not StopArchiving Then
                If LogQueue Is Nothing Then
                    'Crea la coda messaggi 
                    LogQueue = New Generic.Queue(Of TegServiceTraceEventArgs)
                End If

                'Verifica che la coda messaggi non sia piena
                If LogQueue.Count > 0 AndAlso LogQueue.Count >= TraceLimit Then
                    'La coda messaggi ha raggiunto il limite di messaggi consentiti, quindi rimuove il messaggio più vecchio
                    LogQueue.Dequeue()
                End If

                'Se l'ultimo messaggio ricevuto non è valorizzato, allora può aggiungere il messaggio ricevuto nella coda messaggi
                If LastTraceMessage Is Nothing Then
                    LogQueue.Enqueue(e)
                    RaiseEvent WorkState(e.Message, e.Type)
                    SameMessageCounter = 0
                ElseIf LastTraceMessage.Message <> e.Message Then
                    'Il messaggio ricevuto è diverso dall'ultimo ricevuto, quindi registra prima il messaggio salvato e poi quello ricevuto
                    'In questo modo si evita di registrare nella coda i messaggi identici ricevuti in sequenza, ma solo il primo e l'ultimo della serie
                    If SameMessageCounter > 0 Then
                        LogQueue.Enqueue(LastTraceMessage)
                        RaiseEvent WorkState(LastTraceMessage.Message, LastTraceMessage.Type)
                        'Verifica che la coda messaggi non sia piena
                        If LogQueue.Count >= TraceLimit Then
                            'La coda messaggi ha raggiunto il limite di messaggi consentiti, quindi rimuove il messaggio più vecchio
                            LogQueue.Dequeue()
                        End If
                        SameMessageCounter = 0
                    End If
                    LogQueue.Enqueue(e)
                    RaiseEvent WorkState(e.Message, e.Type)
                Else
                    Try
                        SameMessageCounter += 1
                    Catch ex As Exception
                        'E' arrivato a fondo scala, registra il messaggio e ricomincia il conteggio
                        LogQueue.Enqueue(LastTraceMessage)
                        RaiseEvent WorkState(LastTraceMessage.Message, LastTraceMessage.Type)
                        SameMessageCounter = 0
                    End Try
                End If
                'Salva il messaggio ricevuto come ultimo
                LastTraceMessage = e
            End If
        End Sub

        Private Sub SaveMessageQueue()
            'Segnala di sospendere l'archiviazione dei messaggi nella coda messaggi
            StopArchiving = True
            'Verifica che la coda messaggi esista e contenga dei messaggi da salvare
            If LogQueue IsNot Nothing AndAlso LogQueue.Count > 1 Then
                Dim oText As New System.Text.StringBuilder()
                Dim nCounter As Integer = 0
                TraceLog(String.Format("Saving message queue log content on local text file... {0} messages to save...", LogQueue.Count))
                'Scarica la coda messaggi
                Do While LogQueue.Count > 0
                    nCounter += 1
                    Using e As TegServiceTraceEventArgs = LogQueue.Dequeue()
                        oText.Append(e.Type.ToString)
                        oText.Append(vbTab)
                        oText.Append(e.Time.ToString)
                        oText.Append(vbTab)
                        oText.Append(e.Message)
                        oText.Append(vbTab)
                        oText.Append(e.State.ToString)
                        oText.Append(vbTab)
                        If e.Source = "" Then
                            oText.Append(ProjectInstaller.EVLOG_SOURCE)
                        Else
                            oText.Append(e.Source)
                        End If
                        oText.AppendLine()
                    End Using
                Loop
                'Aggiunge un messaggio per indicare l'avvenuto salvataggio della coda e il numero di messaggi salvati
                oText.Append(EventLogEntryType.Information.ToString)
                oText.Append(vbTab)
                oText.Append(Now.ToString)
                oText.Append(vbTab)
                oText.Append(String.Format("Message queue log content saved on local text file: {0} messages saved.", nCounter))
                oText.Append(vbTab)
                oText.AppendLine(ServiceWorker.TegServiceStateType.Unknown.ToString)
                Try
                    'Se esiste già il file di log, ne fa una copia di backup prima di sovrascriverlo
                    Dim sFile As String = String.Concat(Application.StartupPath, "\", ProjectInstaller.EVLOG_SOURCE, ".log")
                    If My.Computer.FileSystem.FileExists(sFile) Then
                        My.Computer.FileSystem.CopyFile(sFile, sFile.Replace(".log", ".old"), True)
                    End If
                    'Salva la coda messaggi in un file di testo nella stessa cartella dove si trova il programma
                    My.Computer.FileSystem.WriteAllText(sFile, oText.ToString, False)
                    TraceLog(String.Format("Message queue log content saved on local text file: {0} messages saved in {1}", nCounter, sFile))
                Catch ex As Exception
                    TraceError(String.Format("ERROR: {0} ", ex.ToString))
                End Try
            End If
        End Sub

        Friend Sub TraceError(ByVal Message As String)
            TraceLog(Message, EventLogEntryType.Error)
        End Sub

        Friend Sub TraceLog(ByVal Message As String)
            TraceLog(Message, EventLogEntryType.Information)
        End Sub
        Friend Sub TraceLog(ByVal Message As String, ByVal LogType As EventLogEntryType)
            If moServiceWorker IsNot Nothing Then
                ArchiveMessage(Me, New TegServiceTraceEventArgs(Message, LogType, moServiceWorker.State, ProjectInstaller.EVLOG_SOURCE))
            Else
                ArchiveMessage(Me, New TegServiceTraceEventArgs(Message, LogType, ServiceWorker.TegServiceStateType.Unknown, ProjectInstaller.EVLOG_SOURCE))
            End If
            'If WriteOnEventLog Then
            '  'RaiseEvent WorkState(Message, LogType)
            'End If
        End Sub

        Friend Sub TraceWork(ByVal Message As String)
            TraceWork(Message, False)
        End Sub
        Friend Sub TraceWork(ByVal Message As String, ByVal Warning As Boolean)
            If Warning Then
                TraceLog(Message, EventLogEntryType.Warning)
            Else
                TraceLog(Message, EventLogEntryType.Information)
            End If
        End Sub

        Private Sub TimerLavoro_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles TimerLavoro.Elapsed
            Try
                'Arresto il timer, per evitare concatenazioni indesiderate di eventi
                TimerLavoro.Stop()
                'Richiama il metodo della classe che esegue le attività temporizzate
                moServiceWorker.DoSomething()
                'Riabilito il timer solo se la chiamata del metodo non ha generato errori non gestiti e il servizio non è stato messo in pausa
                If Not mbIsPaused AndAlso TimerLavoro IsNot Nothing Then
                    'Verifico che la temporizzazione sia quella prevista, in caso contrario la reimposta
                    If TimerLavoro.Interval <> TimerInterval Then TimerLavoro.Interval = TimerInterval
                    'Riattiva il timer se necessario
                    If TimerInterval > 0 Then TimerLavoro.Start()
                End If
            Catch ex As Exception
                'E' stato rilevato un errore non gestito, quindi segnala l'anomalia e prova a fermare il servizio
                TraceError(String.Format("ERROR: {0} ", ex.ToString))
                mbIsRunning = False
                RaiseEvent WorkFailed(Me, New EventArgs)
            End Try
        End Sub

        Public Sub New()
            MyBase.New()
            Dim dnum As Double = System.Threading.Thread.CurrentThread.GetHashCode()
            msInstanceName = String.Concat(My.Application.Info.AssemblyName, "-", Convert.ToInt32(dnum))
        End Sub

        Private disposedValue As Boolean = False    ' Per rilevare chiamate ridondanti

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: liberare altro stato (oggetti gestiti).
                    If TimerLavoro IsNot Nothing Then
                        TimerLavoro.Stop()
                        TimerLavoro.Dispose()
                    End If
                    If moServiceWorker IsNot Nothing Then moServiceWorker.Dispose()
                End If

                ' TODO: liberare lo stato personale (oggetti non gestiti).
                ' TODO: impostare campi di grandi dimensioni su null.
                TimerLavoro = Nothing
                moServiceWorker = Nothing
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' Questo codice è aggiunto da Visual Basic per implementare in modo corretto il modello Disposable.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Non modificare questo codice. Inserire il codice di pulitura in Dispose(ByVal disposing As Boolean).
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace