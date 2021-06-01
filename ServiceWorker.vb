Imports Microsoft.VisualBasic
Imports System.Data
Imports log4net
Imports System.Data.SqlClient
Imports System.Web.Mail

Namespace TegService
    '''---------------------------------------------------------------------------------------------------
    '''Autore.......: Antonello Lo Bianco
    '''Data ........: 02/05/2014
    '''Descrizione..: Creazione della classe applicativa.
    '''Progetto ....: ReminderServiceMail
    '''Revisioni....: Indicare qui di seguito le modifiche apportate
    '''---------------------------------+-------------------------+---------------------------------------
    '''Nome e Cognome                   Data                      Descrizione revisione
    '''---------------------------------+-------------------------+---------------------------------------
    '''Antonello Lo Bianco                  02/05/2014                Creazione classe.
    '''---------------------------------+-------------------------+---------------------------------------
    Public Class ServiceWorker
        Implements IDisposable
        Private Shared ReadOnly Log As ILog = LogManager.GetLogger(GetType(ServiceWorker))


        Private ListaAttachment As Generic.List(Of String)
        Private LastWorkTime As Date = Nothing
        Private LastCheckDate As Date = Nothing

        Public Enum TegServiceStateType
            Unknown
            Waiting
            Running
            Paused
            Stopped
            ApplicationError
        End Enum

        Public Event StateChanged(ByVal sender As Object, ByVal e As TegServiceStateChangedEventArgs)
        Public Event Trace(ByVal sender As Object, ByVal e As TegServiceTraceEventArgs)
        Public Event StopService(ByVal sender As Object, ByVal e As System.EventArgs)

        Private WorkState As TegServiceStateType
        Public Property State() As TegServiceStateType
            Get
                Return WorkState
            End Get
            Set(ByVal value As TegServiceStateType)
                If WorkState <> value Then
                    Dim tempState As TegServiceStateType = WorkState
                    WorkState = value
                    RaiseEvent StateChanged(Me, New TegServiceStateChangedEventArgs(value, tempState))
                End If
            End Set
        End Property

        ''' <summary>
        ''' Funzione che controlla se ci sono mail da inviare e nel caso le invia.
        ''' </summary>
        ''' <revisionHistory>
        ''' <revision date="12/04/2019">
        ''' Modifiche da revisione manuale
        ''' </revision>
        ''' </revisionHistory>
        Private Sub CheckForEMailToSend()
            Dim sSQL As String = ""
            Dim sErr As String = String.Empty
            Dim ListRecipient As New List(Of String)
            Dim ListBCC As New List(Of String)
            Dim ListAttachments = New List(Of MailAttachment)
            Dim myMail As New BaseMail
            Dim LottoPrec As Integer = 0
            Dim RecipientPrec As Integer = 0
            Try
                sSQL = "exec prc_GetMailToSend"
                Dim myDataView As DataView = DB.CaricaDV(sSQL)
                If myDataView.Count > 0 Then
                    Log.Debug("CheckForEmailToSend Start at:" + DateTime.Now.ToString())
                End If
                For Each myRow As DataRowView In myDataView
                    If LottoPrec <> CInt(myRow("lotto")) Then
                        myMail = New BaseMail
                        'dati mail
                        myMail.ID = CInt(myRow("lotto"))
                        myMail.Sender = myRow("SENDER").ToString()
                        myMail.SenderName = myRow("SENDERNAME").ToString()
                        myMail.SSL = CInt(myRow("SSL"))
                        myMail.Server = myRow("SERVER").ToString
                        myMail.ServerPort = myRow("SERVERPORT").ToString()
                        myMail.Password = myRow("PASSWORD").ToString()
                        myMail.WarningRecipient = myRow("WARNINGRECIPIENT").ToString()
                        myMail.WarningSubject = myRow("WARNINGSUBJECT").ToString()
                        myMail.WarningMessage = myRow("WARNINGMESSAGE").ToString()
                        myMail.SendErrorMessage = My.Settings.MailSendErrorMessage
                        myMail.Subject = myRow("oggetto_mail").ToString
                        myMail.Message = myRow("testo_mail").ToString()
                        'elenco destinatari copia nascosta
                        ListBCC = New List(Of String)
                        ListBCC.Add(myRow("email_administrative").ToString)
                    End If
                    If RecipientPrec <> CInt(myRow("iddest")) Then
                        If ListRecipient.Count > 0 Then
                            'Log.Debug("invio per:ID =" + myRow("lotto").ToString() + ",Sender =" + myRow("SENDER").ToString() + ",SenderName =" + myRow("SENDERNAME").ToString() + ",SSL =" + myRow("SSL").ToString() + ",Server =" + myRow("SERVER").ToString + ",ServerPort =" + myRow("SERVERPORT").ToString() + ",Password =" + myRow("PASSWORD").ToString() + ",WarningRecipient =" + myRow("WARNINGRECIPIENT").ToString() + ",WarningSubject =" + myRow("WARNINGSUBJECT").ToString() + ",WarningMessage =" + myRow("WARNINGMESSAGE").ToString() + ",SendErrorMessage =" + My.Settings.MailSendErrorMessage + ",Subject =" + myRow("oggetto_mail").ToString + ",Message =" + myRow("testo_mail").ToString() + ",ListBCC =" + myRow("email_administrative").ToString())
                            'ho dei destinatari precedenti
                            CreateMail(myMail, ListRecipient, ListBCC, ListAttachments, CInt(myRow("iddest")), sErr)
                        End If
                        'elenco destinatari
                        ListRecipient = New List(Of String)
                        ListAttachments = New List(Of MailAttachment)
                        ListRecipient.Add(myRow("email_dest").ToString)
                    End If
                    'elenco allegati
                    ListAttachments.Add(New MailAttachment(myRow("attachment_pathname")))
                    LottoPrec = CInt(myRow("lotto"))
                    RecipientPrec = CInt(myRow("iddest"))
                Next
                'invio l'ultimo soggetto
                CreateMail(myMail, ListRecipient, ListBCC, ListAttachments, RecipientPrec, sErr)
                If myDataView.Count > 0 Then
                    Log.Debug("CheckForEmailToSend End at:" + DateTime.Now.ToString)
                End If
            Catch ex As Exception
                Log.Debug("ReminderServiceMail.ServiceWorker.ChekForEMailToSend.errore::", ex)
            End Try
        End Sub
        'Private Sub CheckForEmailToSend()
        '    Log.Debug("CheckForEmailToSend Start at:" + DateTime.Now.ToString("dd/MM/yyy"))
        '    'Verifica la presenza di ore di ferie pianificate nella settimana a seguire ed invia le email ai diretti interessati e relativi PM
        '    Dim ret As Boolean = False
        '    Dim sqlQuery As String = ""
        '    Dim PathFileName As String = String.Empty
        '    Dim sqlQueryUpdateLotto As String = String.Empty
        '    Dim SendErrorEmail = String.Empty
        '    Dim CanSend As Boolean = False
        '    Dim FilesToAttachment = String.Empty
        '    Dim myArray(-1) As String

        '    DB.ApriConnessione()
        '    If Not DB.HasError Then
        '        Try

        '            sqlQuery = "SELECT * FROM INVIOMAIL_DEST WHERE STATO='W'"

        '            Dim dv As DataView = DB.CaricaDV(sqlQuery)

        '            If dv IsNot Nothing AndAlso dv.Count > 0 Then



        '                Dim mailTo As String = ""
        '                Dim mail As New DriverEmailConnector.ManageEmail()
        '                mail.MailServer = My.Settings.mailServer
        '                mail.ServerPort = My.Settings.mailServerPort
        '                mail.Username = My.Settings.mailUser
        '                mail.Password = My.Settings.mailPassword
        '                mail.EnableSSL = My.Settings.mailSSL

        '                For i As Integer = 0 To dv.Count - 1
        '                    CanSend = True

        '                    Array.Clear(myArray, 0, myArray.Length)

        '                    ListaAttachment.Clear()
        '                    ListaAttachment = Nothing
        '                    ListaAttachment = New List(Of String)
        '                    'Ricava l'elenco degli indirizzi di email al quale inviare la notifica
        '                    sqlQuery = "SELECT * FROM INVIOMAIL_LOTTO WHERE LOTTO =" & dv(i)("fkLOTTO")
        '                    Dim dve As DataView = DB.CaricaDV(sqlQuery)
        '                    For j = 0 To dve.Count - 1

        '                        'Se una delle voci è vuota non manda l'email
        '                        If Globals.IsNullOrEmpty(dve(j)("EMAIL_FROM")) Or Globals.IsNullOrEmpty(dve(j)("OGGETTO_MAIL")) Or Globals.IsNullOrEmpty(dve(j)("TESTO_MAIL")) Then
        '                            CanSend = False
        '                            Log.Error("Tabella INVIOMAIL_LOTTO Fields EMAIL_FROM o OGGETTO_MAIL o TESTO_MAIL vuoti non invio nessuna email")
        '                        Else

        '                            mail.Subject = dve(j)("OGGETTO_MAIL")
        '                            mail.Message = dve(j)("TESTO_MAIL")
        '                            mail.SenderEmail = dve(j)("EMAIL_FROM")
        '                            mail.SenderName = My.Settings.mailSenderName 'dve(j)("EMAIL_FROM")

        '                            SendErrorEmail = dve(j)("EMAIL_ADMINISTRATIVE")
        '                        End If
        '                    Next

        '                    If dve IsNot Nothing Then dve.Dispose()
        '                    dve = Nothing

        '                    If CanSend Then

        '                        'Ha trovato delle email da inviare 
        '                        'aggiorno il Data Base Stato P in invio
        '                        Log.Debug("Aggiornamento Tabella Stato P")

        '                        sqlQueryUpdateLotto = "UPDATE INVIOMAIL_DEST SET DATA_INIZIO_INVIO = GETDATE(), STATO ='P' " & vbCrLf &
        '                                              "WHERE ID = {0}"

        '                        sqlQueryUpdateLotto = String.Format(sqlQueryUpdateLotto, dv(i)("ID"))

        '                        If Not DB.EseguiDML(sqlQueryUpdateLotto) Then
        '                            If DB.HasError Then
        '                                Log.Error("Errore durnate Update Tabella INVIOMAIL_DEST STATO P: " & DB.exc.Message)
        '                                Globals.sErrMsg = "Errore durnate Update Tabella INVIOMAIL_DEST STATO P: " & DB.exc.Message
        '                                Globals.ListaErrori.Add(Globals.sErrMsg)
        '                                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg & vbCrLf & sqlQueryUpdateLotto, EventLogEntryType.Error, State))
        '                                Exit For
        '                            Else

        '                                Globals.sErrMsg = String.Format("ID INVIOMAIL_DEST {0} NON aggiornato su DataBase. STATO P", dv(i)("fkLOTTO"))
        '                                Log.Error(Globals.sErrMsg)
        '                                Globals.ListaErrori.Add(Globals.sErrMsg)
        '                                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg, EventLogEntryType.Warning, State))
        '                            End If
        '                        Else
        '                            Log.Debug("Aggiornamento Tabella Stato P andato a buon fine")
        '                        End If


        '                        mailTo = ""
        '                        mailTo &= dv(i)("EMAIL_DEST") & ";"

        '                        Log.Debug("mailTo " + mailTo)

        '                        sqlQuery = "SELECT * FROM INVIOMAIL_ATTACHMENTS WHERE fkIDMAIL_DEST =" & dv(i)("ID")

        '                        Dim dvefa As DataView = DB.CaricaDV(sqlQuery)
        '                        If dvefa IsNot Nothing AndAlso dvefa.Count > 0 Then
        '                            For r = 0 To dvefa.Count - 1
        '                                PathFileName = dvefa(r)("ATTACHMENT_PATHNAME")
        '                                ListaAttachment.Add(PathFileName)

        '                            Next

        '                            myArray = ListaAttachment.ToArray()


        '                        End If
        '                        If ListaAttachment.Count > 0 Then
        '                            If myArray.Length > 0 Then
        '                                FilesToAttachment = String.Join(";", myArray)
        '                            End If

        '                        End If

        '                        sqlQueryUpdateLotto = String.Empty
        '                        If mailTo.Length > 0 Then
        '                            'Destinatari della mail
        '                            mail.Recipient = mailTo
        '                            'Invio della mail
        '                            If mail.MailServer.Length > 0 Then
        '                                If ListaAttachment.Count > 0 And myArray.Length > 0 Then
        '                                    ret = mail.SendMailAttachment(FilesToAttachment)
        '                                    If ret = True Then

        '                                        sqlQueryUpdateLotto = "UPDATE INVIOMAIL_DEST SET DATA_ESITO = GETDATE(), STATO ='Y' " & vbCrLf &
        '                                              "WHERE ID = {0}"
        '                                        Log.Debug("Aggiornamento Tabella Stato Y")
        '                                        sqlQueryUpdateLotto = String.Format(sqlQueryUpdateLotto, dv(i)("ID"))

        '                                        If Not DB.EseguiDML(sqlQueryUpdateLotto) Then
        '                                            If DB.HasError Then
        '                                                Globals.sErrMsg = "Errore durnate Update Tabella INVIOMAIL_DEST: STATO Y" & DB.exc.Message
        '                                                Log.Error(Globals.sErrMsg)
        '                                                Globals.ListaErrori.Add(Globals.sErrMsg)
        '                                                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg & vbCrLf & sqlQueryUpdateLotto, EventLogEntryType.Error, State))
        '                                                Exit For
        '                                            Else
        '                                                Globals.sErrMsg = String.Format("ID INVIOMAIL_DEST {0} NON aggiornato su DataBase. STATO Y", dv(i)(""))
        '                                                Log.Error(Globals.sErrMsg)
        '                                                Globals.ListaErrori.Add(Globals.sErrMsg)
        '                                                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg, EventLogEntryType.Warning, State))
        '                                            End If
        '                                        Else
        '                                            Log.Debug("Aggiornamento Tabella Stato Y andato a buon fine")
        '                                        End If

        '                                    Else
        '                                        Log.Debug("SendMailAttachment non andata a buon fine")

        '                                        sqlQueryUpdateLotto = "UPDATE INVIOMAIL_DEST SET DATA_ESITO = GETDATE(), STATO ='N', DETTAGLIOESITO='" + My.Settings.MailSendErrorMessage + "' " & vbCrLf &
        '                                              "WHERE ID = {0}"

        '                                        sqlQueryUpdateLotto = String.Format(sqlQueryUpdateLotto, dv(i)("ID"))

        '                                        If Not DB.EseguiDML(sqlQueryUpdateLotto) Then
        '                                            If DB.HasError Then
        '                                                Globals.sErrMsg = "Errore durnate Update Tabella INVIOMAIL_DEST: STATO N" & DB.exc.Message
        '                                                Log.Error(Globals.sErrMsg)
        '                                                Globals.ListaErrori.Add(Globals.sErrMsg)
        '                                                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg & vbCrLf & sqlQueryUpdateLotto, EventLogEntryType.Error, State))
        '                                                Exit For
        '                                            Else
        '                                                Globals.sErrMsg = String.Format("ID INVIOMAIL_DEST {0} NON aggiornato su DataBase. STATO N", dv(i)("ID"))
        '                                                Log.Error(Globals.sErrMsg)
        '                                                Globals.ListaErrori.Add(Globals.sErrMsg)
        '                                                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg, EventLogEntryType.Warning, State))
        '                                            End If
        '                                        End If

        '                                        Dim mailError As New DriverEmailConnector.ManageEmail()

        '                                        mailError.MailServer = My.Settings.mailServer
        '                                        mailError.ServerPort = My.Settings.mailServerPort
        '                                        mailError.Username = My.Settings.mailUser
        '                                        mailError.Password = My.Settings.mailPassword
        '                                        mailError.EnableSSL = My.Settings.mailSSL
        '                                        mailError.SenderEmail = My.Settings.mailSender
        '                                        mailError.SenderName = My.Settings.mailSenderName
        '                                        mailError.Subject = My.Settings.MailWarningSubject
        '                                        mailError.Recipient = SendErrorEmail
        '                                        mailError.Message = My.Settings.MailWarningMessage + mailTo
        '                                        If mailError.MailServer.Length > 0 Then
        '                                            ret = mailError.SendMail()
        '                                            If ret = False Then
        '                                                Log.Error("Email di Warning Mail non inviate non andata a buon fine")
        '                                            End If
        '                                        End If


        '                                    End If
        '                                Else
        '                                    Log.Error("Lista file da inviare non valorizzata")
        '                                End If
        '                            Else
        '                                Log.Error("MailServer non definito")
        '                            End If
        '                        Else
        '                            Log.Error("MailTo non definito")
        '                        End If
        '                    End If
        '                Next
        '            End If
        '            If dv IsNot Nothing Then dv.Dispose()
        '            dv = Nothing

        '        Catch ex As Exception
        '            RaiseEvent Trace(Me, New TegServiceTraceEventArgs(ex.Message, EventLogEntryType.Warning, State))
        '            Log.Error(ex.Message)
        '        Finally
        '            If DB.HasError Then
        '                'Segnala la condizione di errore
        '                State = TegServiceStateType.ApplicationError
        '                RaiseEvent Trace(Me, New TegServiceTraceEventArgs(DB.exc.Message & vbCrLf & sqlQuery, EventLogEntryType.Error, State))
        '                Log.Error(DB.exc.Message)
        '            End If
        '            'Nel caso sia rimasta aperta, chiude la connessione al database
        '            DB.ChiudiConnessione()
        '        End Try
        '    End If
        '    Log.Debug("CheckForEmailToSend End at:" + DateTime.Now.ToString("dd/MM/yyy"))
        'End Sub
        Public Sub CreateMail(myMail As BaseMail, ListRecipient As List(Of String), ListRecipientBCC As List(Of String), ListAttachment As List(Of MailAttachment), IdRecipient As Integer, ByRef sErr As String)
            Dim fncEmail As New EmailService()
            Dim sSQL As String = ""
            sErr = String.Empty
            Try
                If myMail.ID > 0 Then
                    fncEmail.SendEmail(myMail, ListRecipient, New List(Of String), ListRecipientBCC, ListAttachment)
                If sErr = "" Then
                    sErr = "OK"
                End If
                'Log.Debug("OPENgovSPORTELLO.BLL.Messages.SendMail::errore in createmal::" + sErr)
                sSQL = "exec prc_SetMailResult {0},{1}"
                sSQL = String.Format(sSQL, IdRecipient, sErr)
                    If Not DB.EseguiDML(sSQL) Then
                        If DB.HasError Then
                            Globals.sErrMsg = "Errore durante Update Tabella INVIOMAIL_DEST: STATO N" & DB.exc.Message
                            Log.Error(Globals.sErrMsg)
                            Globals.ListaErrori.Add(Globals.sErrMsg)
                            RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg & vbCrLf & sSQL, EventLogEntryType.Error, State))
                        Else
                            Globals.sErrMsg = String.Format("ID INVIOMAIL_DEST {0} NON aggiornato su DataBase. STATO N", myMail.ID)
                            Log.Error(Globals.sErrMsg)
                            Globals.ListaErrori.Add(Globals.sErrMsg)
                            RaiseEvent Trace(Me, New TegServiceTraceEventArgs(Globals.sErrMsg, EventLogEntryType.Warning, State))
                        End If
                    End If
                End If
            Catch ex As Exception
                Log.Debug("ReminderServiceMail.ServiceWorker.CreateMail.errore::", ex)
                sErr = ex.Message
            End Try
        End Sub

        Public Sub DoSomething()
            'Procede con l'esecuzione del codice solo se sia passato un minuto dall'ultima esecuzione
            If LastWorkTime.AddSeconds(Val(Globals.NVLE(My.Settings.CheckDelayTime, "60"))) < Now Then
                LastWorkTime = Now
                State = TegServiceStateType.Running
                If Globals.ListaErrori IsNot Nothing Then Globals.ListaErrori.Clear()
                If ListaAttachment Is Nothing Then
                    ListaAttachment = New List(Of String)
                Else
                    ListaAttachment.Clear()
                End If

                CheckForEMailToSend()

                LastCheckDate = Now
            ElseIf Debugger.IsAttached Then
                'ElseIf 1 = 0 Then
                'Exit point x DEBUG del servizio
                State = TegServiceStateType.ApplicationError
                RaiseEvent Trace(Me, New TegServiceTraceEventArgs("DEBUG: Internal code signal to stop the job.", EventLogEntryType.Warning, State))
                RaiseEvent StopService(Me, New EventArgs)
            End If
        End Sub

        Public Function Initialize(ByVal CommandLine As String, Optional ByVal Debug As Boolean = False) As Boolean
            State = TegServiceStateType.Waiting
            If CommandLine.Length > 0 Then
                RaiseEvent Trace(Me, New TegServiceTraceEventArgs("job initialized with params: " & CommandLine, EventLogEntryType.Information, State))
            Else
                RaiseEvent Trace(Me, New TegServiceTraceEventArgs("job initialized.", EventLogEntryType.Information, State))
            End If

            LastWorkTime = Today
            LastCheckDate = Today.AddDays(-1)
            Globals.ListaErrori = New List(Of String)

            Return True
        End Function

        Public Sub JobStatus()
            RaiseEvent Trace(Me, New TegServiceTraceEventArgs("job status log.", EventLogEntryType.Information, State))
        End Sub

        Public Sub PauseJob()
            State = TegServiceStateType.Paused
            RaiseEvent Trace(Me, New TegServiceTraceEventArgs("job paused.", EventLogEntryType.Information, State))
        End Sub

        Public Sub ResumeJob()
            State = TegServiceStateType.Running
            RaiseEvent Trace(Me, New TegServiceTraceEventArgs("job resumed.", EventLogEntryType.Information, State))
        End Sub

        Public Sub Terminate()
            State = TegServiceStateType.Stopped
            RaiseEvent Trace(Me, New TegServiceTraceEventArgs("job terminated.", EventLogEntryType.Information, State))
        End Sub

        Private disposedValue As Boolean = False    ' Per rilevare chiamate ridondanti

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: liberare altro stato (oggetti gestiti).
                End If

                ' TODO: liberare lo stato personale (oggetti non gestiti).
                ' TODO: impostare campi di grandi dimensioni su null.
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

    Public Class TegServiceStateChangedEventArgs
        Inherits System.EventArgs
        Implements IDisposable

        Private state1 As ServiceWorker.TegServiceStateType
        Public ReadOnly Property CurrentState() As ServiceWorker.TegServiceStateType
            Get
                Return state1
            End Get
        End Property

        Private state2 As ServiceWorker.TegServiceStateType
        Public ReadOnly Property PreviuosState() As ServiceWorker.TegServiceStateType
            Get
                Return state2
            End Get
        End Property

        Public Sub New(ByVal NewState As ServiceWorker.TegServiceStateType, ByVal OldState As ServiceWorker.TegServiceStateType)
            MyBase.New()
            state1 = NewState
            state2 = OldState
        End Sub


        Private disposedValue As Boolean = False    ' Per rilevare chiamate ridondanti

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: liberare altro stato (oggetti gestiti).
                End If

                ' TODO: liberare lo stato personale (oggetti non gestiti).
                ' TODO: impostare campi di grandi dimensioni su null.
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
    Public Class TegServiceTraceEventArgs
        Inherits System.EventArgs
        Implements IDisposable

        Private msgid As Long
        Public ReadOnly Property ID() As Long
            Get
                Return msgid
            End Get
        End Property

        Private msglv As Integer
        Public ReadOnly Property Level() As Integer
            Get
                Return msglv
            End Get
        End Property

        Private msg As String
        Public ReadOnly Property Message() As String
            Get
                Return msg
            End Get
        End Property

        Private msgsource As String
        Public ReadOnly Property Source() As String
            Get
                Return msgsource
            End Get
        End Property

        Private st As ServiceWorker.TegServiceStateType
        Public ReadOnly Property State() As ServiceWorker.TegServiceStateType
            Get
                Return st
            End Get
        End Property

        Private msgtime As Date
        Public ReadOnly Property Time() As Date
            Get
                Return msgtime
            End Get
        End Property

        Private msgtype As EventLogEntryType
        Public ReadOnly Property Type() As EventLogEntryType
            Get
                Return msgtype
            End Get
        End Property

        Public Overrides Function ToString() As String
            Dim sRetVal As String = Type.ToString & vbTab & Level.ToString & vbTab & ID.ToString & vbTab & Time.ToString & vbTab & Message & vbTab & State.ToString & vbTab & Source
            Return sRetVal
        End Function

        Public Sub New(ByVal LogMessage As String, ByVal LogType As EventLogEntryType, ByVal JobState As ServiceWorker.TegServiceStateType)
            Me.New(LogMessage, LogType, 1, JobState, "")
        End Sub
        Public Sub New(ByVal LogMessage As String, ByVal LogType As EventLogEntryType, ByVal JobState As ServiceWorker.TegServiceStateType, ByVal LogSource As String)
            Me.New(LogMessage, LogType, 1, JobState, LogSource)
        End Sub
        Public Sub New(ByVal LogMessage As String, ByVal LogType As EventLogEntryType, ByVal LogLevel As Integer, ByVal JobState As ServiceWorker.TegServiceStateType, ByVal LogSource As String)
            MyBase.New()
            msgid = Globals.GetTraceMessageID
            msg = LogMessage
            msgsource = LogSource
            msgtype = LogType
            msglv = LogLevel
            msgtime = Now
            st = JobState
        End Sub

        Private disposedValue As Boolean = False    ' Per rilevare chiamate ridondanti

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: liberare altro stato (oggetti gestiti).
                End If

                ' TODO: liberare lo stato personale (oggetti non gestiti).
                ' TODO: impostare campi di grandi dimensioni su null.
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
    Public Class Globals
        Public Shared logger As log4net.ILog

        'Public Shared oCompany As SAPbobsCOM.Company = Nothing
        Public Shared sErrMsg As String = ""
        Public Shared lErrCode As Integer = 0
        Public Shared lRetCode As Integer = 0

        Public Shared ListaErrori As Generic.List(Of String)
        Public Shared ListaEMailNonInviate As Generic.List(Of String)
        Public Shared Mailer As DriverEmailConnector.ManageEmail
        Private Shared TraceMessageIDCounter As Long = Int(Date.Now.ToOADate * 10 ^ 10)
        Public Shared Function GetTraceMessageID() As Long
            TraceMessageIDCounter += 1
            Return TraceMessageIDCounter
        End Function

        Public Shared Function IsNull(ByVal ValueToCheck As Object) As Boolean
            If ValueToCheck Is Nothing OrElse IsDBNull(ValueToCheck) Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Shared Function IsNullOrEmpty(ByVal ValueToCheck As Object) As Boolean
            If ValueToCheck Is Nothing OrElse IsDBNull(ValueToCheck) OrElse ValueToCheck.ToString = "" OrElse ValueToCheck.ToString.Trim.Length = 0 Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Shared Function NVL(ByVal ValueToCheck As Object, ByVal ValueIfNull As Object) As Object
            If IsNull(ValueToCheck) Then
                Return ValueIfNull
            Else
                Return ValueToCheck
            End If
        End Function
        Public Shared Function NVLE(ByVal ValueToCheck As Object, ByVal ValueIfNullOrEmpty As Object) As Object
            If IsNullOrEmpty(ValueToCheck) Then
                Return ValueIfNullOrEmpty
            Else
                Return ValueToCheck
            End If
        End Function

        Public Shared Function SubString(ByVal Source As String, ByVal StartDelimiter As String, ByVal EndDelimiter As String, Optional ByVal DelimitersIncluded As Boolean = False) As String
            Dim Ret As String = ""
            Dim StartPos As Integer = Source.IndexOf(StartDelimiter, 0, StringComparison.OrdinalIgnoreCase)
            Dim EndPos As Integer = -1
            If StartPos + StartDelimiter.Length < Source.Length Then
                EndPos = Source.IndexOf(EndDelimiter, StartPos + StartDelimiter.Length, StringComparison.OrdinalIgnoreCase)
            End If
            Select Case True
                Case StartDelimiter = "" AndAlso EndDelimiter = ""
                    Ret = Source
                Case StartDelimiter = "" AndAlso EndDelimiter <> "" AndAlso EndPos > -1
                    Ret = Source.Substring(0, EndPos)
                Case StartDelimiter <> "" AndAlso EndDelimiter = "" AndAlso StartPos > -1
                    Ret = Source.Substring(StartPos + StartDelimiter.Length)
                Case StartDelimiter <> "" AndAlso EndDelimiter <> "" AndAlso EndPos > -1 AndAlso StartPos > -1
                    Ret = Source.Substring(StartPos + StartDelimiter.Length, EndPos - (StartPos + StartDelimiter.Length))
            End Select

            If DelimitersIncluded Then
                If StartPos > -1 Then Ret = StartDelimiter & Ret
                If EndPos > -1 Then Ret &= EndDelimiter
            End If
            Return Ret
        End Function
    End Class
    Public Class RowInfo
        Public codiceFiscale As String = ""
        Public matricola As Integer = 0
        Public cognome As String = ""
        Public nome As String = ""
        Public dataInizio As Date = Date.MinValue
        Public dataFine As Date = Date.MinValue
        Public tipo As String = ""

        Public Sub Clear()
            codiceFiscale = ""
            matricola = 0
            cognome = ""
            nome = ""
            dataInizio = Date.MinValue
            dataFine = Date.MinValue
            tipo = ""
        End Sub
    End Class

    Public Class BaseMail
        Public Sub New()
            MyBase.New
            Me.Reset()
        End Sub
#Region "Public properties"
        Dim _ID As Integer
        Dim _Sender As String
        Dim _SenderName As String
        Dim _SSL As String
        Dim _Server As String
        Dim _ServerPort As String
        Dim _Password As String
        Dim _WarningRecipient As String
        Dim _WarningSubject As String
        Dim _WarningMessage As String
        Dim _SendErrorMessage As String
        Dim _Subject As String
        Dim _Message As String

        Public Property ID As Integer
            Get
                Return _ID
            End Get
            Set(ByVal Value As Integer)
                _ID = Value
            End Set
        End Property
        Public Property Sender As String
            Get
                Return _Sender
            End Get
            Set(ByVal Value As String)
                _Sender = Value
            End Set
        End Property

        Public Property SenderName As String
            Get
                Return _SenderName
            End Get
            Set(ByVal Value As String)
                _SenderName = Value
            End Set
        End Property
        Public Property SSL As Integer
            Get
                Return _SSL
            End Get
            Set(ByVal Value As Integer)
                _SSL = Value
            End Set
        End Property
        Public Property Server As String
            Get
                Return _Server
            End Get
            Set(ByVal Value As String)
                _Server = Value
            End Set
        End Property
        Public Property ServerPort As String
            Get
                Return _ServerPort
            End Get
            Set(ByVal Value As String)
                _ServerPort = Value
            End Set
        End Property
        Public Property Password As String
            Get
                Return _Password
            End Get
            Set(ByVal Value As String)
                _Password = Value
            End Set
        End Property
        Public Property WarningRecipient As String
            Get
                Return _WarningRecipient
            End Get
            Set(ByVal Value As String)
                _WarningRecipient = Value
            End Set
        End Property
        Public Property WarningSubject As String
            Get
                Return _WarningSubject
            End Get
            Set(ByVal Value As String)
                _WarningSubject = Value
            End Set
        End Property
        Public Property WarningMessage As String
            Get
                Return _WarningMessage
            End Get
            Set(ByVal Value As String)
                _WarningMessage = Value
            End Set
        End Property
        Public Property SendErrorMessage As String
            Get
                Return _SendErrorMessage
            End Get
            Set(ByVal Value As String)
                _SendErrorMessage = Value
            End Set
        End Property
        Public Property Subject As String
            Get
                Return _Subject
            End Get
            Set(ByVal Value As String)
                _Subject = Value
            End Set
        End Property
        Public Property Message As String
            Get
                Return _Message
            End Get
            Set(ByVal Value As String)
                _Message = Value
            End Set
        End Property
#End Region
#Region "DbObject methods"
        Public Sub Reset()
            _ID = -1
            _Sender = My.Settings.mailServer
            _SenderName = My.Settings.mailSenderName
            _SSL = My.Settings.mailSSL
            _Server = My.Settings.mailServer
            _ServerPort = My.Settings.mailServerPort
            _Password = My.Settings.mailPassword
            _WarningRecipient = My.Settings.MailWarningRecipient
            _WarningSubject = My.Settings.MailWarningSubject
            _WarningMessage = My.Settings.MailWarningMessage
            _SendErrorMessage = My.Settings.MailSendErrorMessage
            _Subject = String.Empty
            _Message = String.Empty
        End Sub

        Public Shared Function implicitOperator(ByVal v As BaseMail) As List(Of Object)
            Throw New NotImplementedException
        End Function
#End Region
    End Class
    '___________________________CODICE INVIO MAIL__________________________________
    Public Class EmailService
        Private Shared ReadOnly Log As ILog = LogManager.GetLogger(GetType(EmailService))

        Public Sub SendEmail(myMail As BaseMail, ListTO As List(Of String), ListCC As List(Of String), ListBCC As List(Of String), ListAttachment As List(Of MailAttachment))
            Dim mail As New MailMessage()
            Dim mailTo As String = String.Empty
            Dim mailCc As String = String.Empty
            Dim mailToBCC As String = String.Empty
            Try
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", myMail.Server)
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", myMail.ServerPort)
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", "2") 'Send the message Using the network (SMTP over the network)
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'YES
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", myMail.Sender)
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", myMail.Password)
                mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", myMail.SSL)

                mail.From = myMail.Sender
                For Each myRecipient As String In ListTO
                    If mailTo <> String.Empty Then
                        mailTo += ";"
                    End If
                    mailTo += myRecipient
                Next
                mail.To = mailTo
                For Each myRecipient As String In ListCC
                    If mailCc <> String.Empty Then
                        mailCc += ";"
                    End If
                    mailCc += myRecipient
                Next
                mail.Cc = mailCc
                If myMail.SSL = 0 Then
                    For Each myRecipient As String In ListBCC
                        If mailToBCC <> String.Empty Then
                            mailToBCC += ";"
                        End If
                        mailToBCC += myRecipient
                    Next
                End If
                mail.Bcc = mailToBCC
                mail.Subject = myMail.Subject
                mail.Body = myMail.message
                For Each myAttach As MailAttachment In ListAttachment
                    mail.Attachments.Add(myAttach)
                Next

                Try
                    SmtpMail.Send(mail)
                Catch mailEx As Exception
                    Log.Debug("OPENgovSPORTELLO.EmailService.SendEmail.Send.errore::", mailEx)
                    mail = New MailMessage()
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpserver", myMail.Server)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpserverport", myMail.ServerPort)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/sendusing", "2") 'Send the message using the network (SMTP over the network)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'YES
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/sendusername", myMail.Sender)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/sendpassword", myMail.Password)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpusessl", myMail.SSL)
                    mail.From = myMail.SenderName
                    mail.To = myMail.WarningRecipient
                    mail.Subject = myMail.WarningSubject
                    mail.Body = (myMail.WarningMessage + " Errore rilevato:" + mailEx.Message + "\nMail inviata a:" + mailTo + "\nMail:" + myMail.Subject + "\n" + myMail.message)
                    SmtpMail.Send(mail)
                    Throw New Exception("Send." + (myMail.WarningMessage + " Errore rilevato:" + mailEx.Message + "\nMail inviata a:" + mailTo + "\nMail:" + myMail.Subject + "\n" + myMail.message) + ".errore." + mailEx.Message)
                End Try
            Catch ex As Exception
                Log.Debug("OPENgovSPORTELLO.EmailService.SendEmail.errore::", ex)
                Try
                    mail = New MailMessage()
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpserver", myMail.Sender)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpserverport", myMail.ServerPort)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/sendusing", "2") 'Send the message using the network (SMTP over the network)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'YES
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/sendusername", myMail.SenderName)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/sendpassword", myMail.Password)
                    mail.Fields.Add("http:'schemas.microsoft.com/cdo/configuration/smtpusessl", myMail.SSL)
                    mail.From = myMail.SenderName
                    mail.To = myMail.WarningRecipient
                    mail.Subject = myMail.WarningSubject
                    mail.Body = (myMail.WarningMessage + " " + mailTo)
                    SmtpMail.Send(mail)
                    Throw New Exception("SendEmail." + (myMail.WarningMessage + " " + mailTo) + ".errore." + ex.Message)
                Catch Err As Exception
                    Log.Debug("OPENgovSPORTELLO.EmailService.SendEmailWarning.errore::", Err)
                    Throw New Exception("SendEmailWarning." + (myMail.WarningMessage + " " + mailTo) + ".errore." + Err.Message)
                End Try
            End Try
        End Sub
    End Class
    '___________________________FINE CODICE INVIO MAIL__________________________________
End Namespace