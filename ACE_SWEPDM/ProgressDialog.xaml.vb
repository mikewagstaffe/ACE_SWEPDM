Imports System.Windows
Imports System.ComponentModel
Imports System.Windows.Threading
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Controls

Namespace AddinDialogs
    Partial Public Class ProgressDialog
        Inherits Window
#Region "fields"

        ''' <summary>
        ''' The background worker which handles asynchronous invocation
        ''' of the worker method.
        ''' </summary>
        Private ReadOnly worker As BackgroundWorker

        ''' <summary>
        ''' The timer to be used for automatic progress bar updated.
        ''' </summary>
        Private ReadOnly progressTimer As DispatcherTimer

        ''' <summary>
        ''' The UI culture of the thread that invokes the dialog.
        ''' </summary>
        Private uiCulture As CultureInfo

        ''' <summary>
        ''' If set, the interval in which the progress bar
        ''' gets incremented automatically.
        ''' </summary>
        Private m_autoIncrementInterval As System.Nullable(Of Integer) = Nothing

        ''' <summary>
        ''' Whether background processing was cancelled by the user.
        ''' </summary>
        Private m_cancelled As Boolean = False

        ''' <summary>
        ''' Defines the size of a single increment of the progress bar.
        ''' Defaults to 5.
        ''' </summary>
        Private m_progressBarIncrement As Integer = 5

        ''' <summary>
        ''' Provides an exception that occurred during the asynchronous
        ''' operation on the worker thread. Defaults to null, which
        ''' indicates that no exception occurred at all.
        ''' </summary>
        Private m_error As Exception = Nothing

        ''' <summary>
        ''' The result, if assigned to the <see cref="DoWorkEventArgs.Result"/>
        ''' property by the worker method.
        ''' </summary>
        Private m_result As Object = Nothing

        ''' <summary>
        ''' The 
        ''' </summary>
        Private workerCallback As DoWorkEventHandler

        ''' <summary>
        ''' Mark This to enable runniong as a non threaded dialog
        ''' </summary>
        ''' <remarks></remarks>
        Private m_NonThreaded As Boolean = False
#End Region


#Region "properties"

        ''' <summary>
        ''' Gets or sets the dialog text.
        ''' </summary>
        Public Property DialogText() As String
            Get
                Return txtDialogMessage.Text
            End Get
            Set(value As String)
                txtDialogMessage.Text = value
            End Set
        End Property


        ''' <summary>
        ''' Whether to enable cancelling the process. This basically
        ''' shows or hides the Cancel button. Defaults to false.
        ''' </summary>
        Public Property IsCancellingEnabled() As Boolean
            Get
                Return btnCancel.IsVisible
            End Get
            Set(value As Boolean)
                btnCancel.Visibility = If(value, Visibility.Visible, Visibility.Collapsed)
            End Set
        End Property


        ''' <summary>
        ''' Whether the process was cancelled by the user.
        ''' </summary>
        Public ReadOnly Property Cancelled() As Boolean
            Get
                Return m_cancelled
            End Get
        End Property

        ''' <summary>
        ''' If set, the interval in which the progress bar
        ''' gets incremented automatically.
        ''' </summary>
        ''' <exception cref="ArgumentOutOfRangeException">If the interval
        ''' is lower than 100 ms.</exception>
        Public Property AutoIncrementInterval() As System.Nullable(Of Integer)
            Get
                Return m_autoIncrementInterval
            End Get
            Set(value As System.Nullable(Of Integer))
                If value.HasValue AndAlso value < 100 Then
                    Throw New ArgumentOutOfRangeException("value")
                End If
                m_autoIncrementInterval = value
            End Set
        End Property

        ''' <summary>
        ''' Defines the size of a single increment of the progress bar.
        ''' The default value is 5, with a progress bar range of 0 - 100.
        ''' </summary>
        Public Property ProgressBarIncrement() As Integer
            Get
                Return m_progressBarIncrement
            End Get
            Set(value As Integer)
                m_progressBarIncrement = value
            End Set
        End Property

        ''' <summary>
        ''' Provides an exception that occurred during the asynchronous
        ''' operation on the worker thread. Defaults to null, which
        ''' indicates that no exception occurred at all.
        ''' </summary>
        Public ReadOnly Property [Error]() As Exception
            Get
                Return m_error
            End Get
        End Property

        ''' <summary>
        ''' The result, if assigned to the <see cref="DoWorkEventArgs.Result"/>
        ''' property by the worker method. Defaults to null.
        ''' </summary>
        Public ReadOnly Property Result() As Object
            Get
                Return m_result
            End Get
        End Property


        ''' <summary>
        ''' Shows or hides the progressbar control. Defaults to
        ''' true.
        ''' </summary>
        Public Property ShowProgressBar() As Boolean
            Get
                Return progressBar.Visibility = Visibility.Visible
            End Get
            Set(value As Boolean)
                progressBar.Visibility = If(value, Visibility.Visible, Visibility.Hidden)
            End Set
        End Property

#End Region


        ''' <summary>
        ''' Inits the dialog with a given dialog text.
        ''' </summary>
        Public Sub New(dialogText__1 As String)
            Me.New()
            DialogText = dialogText__1
        End Sub


        ''' <summary>
        ''' Inits the dialog without displaying it.
        ''' </summary>
        Public Sub New()
            InitializeComponent()

            'init the timer
            progressTimer = New DispatcherTimer(DispatcherPriority.SystemIdle, Dispatcher)
            AddHandler progressTimer.Tick, AddressOf OnProgressTimer_Tick

            'init background worker
            worker = New BackgroundWorker()
            worker.WorkerReportsProgress = True
            worker.WorkerSupportsCancellation = True

            AddHandler worker.DoWork, AddressOf worker_DoWork
            AddHandler worker.ProgressChanged, AddressOf worker_ProgressChanged
            AddHandler worker.RunWorkerCompleted, AddressOf worker_RunWorkerCompleted
        End Sub


#Region "run worker thread"

        ''' <summary>
        ''' Launches a worker thread which is intendet to perform
        ''' work while progress is indicated.
        ''' </summary>
        ''' <param name="workHandler">A callback method which is
        ''' being invoked on a background thread in order to perform
        ''' the work to be performed.</param>
        Public Function RunWorkerThread(workHandler As DoWorkEventHandler) As Boolean
            Return RunWorkerThread(Nothing, workHandler)
        End Function


        ''' <summary>
        ''' Launches a worker thread which is intended to perform
        ''' work while progress is indicated, and displays the dialog
        ''' modally in order to block the calling thread.
        ''' </summary>
        ''' <param name="argument">A custom object which will be
        ''' submitted in the <see cref="DoWorkEventArgs.Argument"/>
        ''' property <paramref name="workHandler"/> callback method.</param>
        ''' <param name="workHandler">A callback method which is
        ''' being invoked on a background thread in order to perform
        ''' the work to be performed.</param>
        Public Function RunWorkerThread(argument As Object, workHandler As DoWorkEventHandler) As Boolean
            If m_autoIncrementInterval.HasValue Then
                'run timer to increment progress bar
                progressTimer.Interval = TimeSpan.FromMilliseconds(m_autoIncrementInterval.Value)
                progressTimer.Start()
            End If

            'store the UI culture
            uiCulture = CultureInfo.CurrentUICulture

            'store reference to callback handler and launch worker thread
            workerCallback = workHandler
            worker.RunWorkerAsync(argument)

            'display modal dialog (blocks caller)
            Return If(ShowDialog(), False)
        End Function

#End Region


#Region "event handlers"

        ''' <summary>
        ''' Worker method that gets called from a worker thread.
        ''' Synchronously calls event listeners that may handle
        ''' the work load.
        ''' </summary>
        Private Sub worker_DoWork(sender As Object, e As DoWorkEventArgs)
            Try
                'make sure the UI culture is properly set on the worker thread
                Thread.CurrentThread.CurrentUICulture = uiCulture

                'invoke the callback method with the designated argument
                workerCallback(sender, e)
            Catch generatedExceptionName As Exception
                'disable cancelling and rethrow the exception
                Dispatcher.BeginInvoke(DispatcherPriority.Normal, DirectCast(Sub() btnCancel.SetValue(Button.IsEnabledProperty, False), SendOrPostCallback), Nothing)

                Throw
            End Try
        End Sub


        ''' <summary>
        ''' Cancels the background worker's progress.
        ''' </summary>
        Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
            btnCancel.IsEnabled = False
            worker.CancelAsync()
            m_cancelled = True
        End Sub


        ''' <summary>
        ''' Visually indicates the progress of the background operation by
        ''' updating the dialog's progress bar.
        ''' </summary>
        Private Sub worker_ProgressChanged(sender As Object, e As ProgressChangedEventArgs)
            If Not Dispatcher.CheckAccess() Then
                'run on UI thread
                Dim handler As ProgressChangedEventHandler = AddressOf worker_ProgressChanged
                Dispatcher.Invoke(DispatcherPriority.SystemIdle, handler, New Object() {sender, e}, Nothing)
                Return
            End If

            If e.ProgressPercentage <> Integer.MinValue Then
                progressBar.Value = e.ProgressPercentage
            End If

            lblStatus.Content = e.UserState
        End Sub


        ''' <summary>
        ''' Updates the user interface once an operation has been completed and
        ''' sets the dialog's <see cref="Window.DialogResult"/> depending on the value
        ''' of the <see cref="AsyncCompletedEventArgs.Cancelled"/> property.
        ''' </summary>
        Private Sub worker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
            If Not Dispatcher.CheckAccess() Then
                'run on UI thread
                Dim handler As RunWorkerCompletedEventHandler = AddressOf worker_RunWorkerCompleted
                Dispatcher.Invoke(DispatcherPriority.SystemIdle, handler, New Object() {sender, e}, Nothing)
                Return
            End If

            If e.[Error] IsNot Nothing Then
                m_error = e.[Error]
            ElseIf Not e.Cancelled Then
                'assign result if there was neither exception nor cancel
                m_result = e.Result
            End If

            'update UI in case closing the dialog takes a moment
            progressTimer.[Stop]()
            progressBar.Value = progressBar.Maximum
            btnCancel.IsEnabled = False

            'set the dialog result, which closes the dialog
            DialogResult = m_error Is Nothing AndAlso Not e.Cancelled
        End Sub


        ''' <summary>
        ''' Periodically increments the value of the progress bar.
        ''' </summary>
        Private Sub OnProgressTimer_Tick(sender As Object, e As EventArgs)
            Dim threshold As Integer = 100 + m_progressBarIncrement
            progressBar.Value = ((progressBar.Value + m_progressBarIncrement) Mod threshold)
        End Sub

#End Region


#Region "update progress bar / status label"

        ''' <summary>
        ''' Directly updates the value of the underlying
        ''' progress bar. This method can be invoked from a worker thread.
        ''' </summary>
        ''' <param name="progress"></param>
        ''' <exception cref="ArgumentOutOfRangeException">If the
        ''' value is not between 0 and 100.</exception>
        Public Sub UpdateProgress(progress As Integer)
            If Not Dispatcher.CheckAccess() Then
                'switch to UI thread
                Dispatcher.BeginInvoke(DispatcherPriority.Background, DirectCast(Sub() UpdateProgress(progress), SendOrPostCallback), Nothing)
                Return
            End If


            'validate range
            If progress < progressBar.Minimum OrElse progress > progressBar.Maximum Then
                Dim msg As String = "Only values between {0} and {1} can be assigned to the progress bar."
                msg = [String].Format(msg, progressBar.Minimum, progressBar.Maximum)
                Throw New ArgumentOutOfRangeException("progress", progress, msg)
            End If

            'set the progress bar's value
            progressBar.SetValue(Controls.Primitives.RangeBase.ValueProperty, progress)
        End Sub


        ''' <summary>
        ''' Sets the content of the status label to a given value. This method
        ''' can be invoked from a worker thread.
        ''' </summary>
        ''' <param name="status">The status to be displayed.</param>
        Public Sub UpdateStatus(status As Object)
            Dispatcher.BeginInvoke(DispatcherPriority.Background, DirectCast(Sub() lblStatus.SetValue(ContentProperty, status), SendOrPostCallback), Nothing)
        End Sub

#End Region


#Region "invoke methods on UI thread"

        ''' <summary>
        ''' Asynchronously invokes a given method on the thread
        ''' of the dialog's dispatcher.
        ''' </summary>
        ''' <param name="method">The method to be invoked.</param>
        ''' <param name="priority">The priority of the operation.</param>
        ''' <returns>The result of the
        ''' method.</returns>
        Public Function BeginInvoke(method As [Delegate], priority As DispatcherPriority) As DispatcherOperation
            Return Dispatcher.BeginInvoke(priority, method)
        End Function


        ''' <summary>
        ''' Synchronously invokes a given method on the thread
        ''' of the dialog's dispatcher.
        ''' </summary>
        ''' <param name="method">The method to be invoked.</param>
        ''' <param name="priority">The priority of the operation.</param>
        ''' <returns>The result of the
        ''' method.</returns>
        Public Function Invoke(method As [Delegate], priority As DispatcherPriority) As Object
            Return Dispatcher.Invoke(priority, method)
        End Function

#End Region
    End Class

End Namespace
