    Public NotInheritable Class LongTimer : Implements IDisposable
        Private Const MAXTIME As Long = UInteger.MaxValue - 1
        Private Const EPSILON As Long = 15
        Private Const RATIO As Double = 0.8
        Private _timer As Timer
        Private _callback As TimerCallback
        Private ReadOnly _innerCallback As TimerCallback = AddressOf InnerCallback
        Private _state As Object
        Private _needDispose As Boolean

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private x As Date
        Private isSimple As Boolean

        Public ReadOnly Property [Date]() As Date
            <DebuggerStepThrough()> _
            Get
                Return x
            End Get
        End Property

        Public Sub New(ByVal callback As TimerCallback, ByVal [date] As Date, Optional ByVal state As Object = Nothing)
            x = [date]
            _callback = callback
            _state = state
            InnerCallback(_state)
        End Sub

        Public Sub New(ByVal callback As TimerCallback, ByVal interval As Long, Optional ByVal state As Object = Nothing)
            x = Date.Now.Add(TimeSpan.FromTicks(interval * 10000L))
            _callback = callback
            _state = state
            InnerCallback(_state)
        End Sub

        Private Sub InnerCallback(ByVal state As Object)
            'Debug.Print("InnerCallback")
            If isSimple Then
Finish:
                _needDispose = True
                Try
                    _callback(If(_state, Me))
                Finally
                    If _needDispose Then Dispose()
                End Try
            Else
                Dim dueTime = (x - Date.Now).Ticks \ 10000L
                If dueTime <= EPSILON Then
                    GoTo Finish
                ElseIf dueTime <= 3600 * 1000 Then
                    isSimple = True
                    Dim ok As Boolean
                    Try
                        If _timer Is Nothing Then
                            _timer = New Timer(_innerCallback, Me, dueTime, Timeout.Infinite)
                            ok = True
                        Else
                            ok = _timer.Change(dueTime, Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok Then
                        If _timer IsNot Nothing Then
                            Try : _timer.Dispose() : Catch : End Try
                        End If
                        _timer = New Timer(_innerCallback, Me, dueTime, Timeout.Infinite)
                    End If
                ElseIf dueTime > MAXTIME Then
                    Dim ok As Boolean
                    Try
                        If _timer Is Nothing Then
                            _timer = New Timer(_innerCallback, Me, CLng(MAXTIME * RATIO), Timeout.Infinite)
                            ok = True
                        Else
                            ok = _timer.Change(CLng(MAXTIME * RATIO), Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok Then
                        If _timer IsNot Nothing Then
                            Try : _timer.Dispose() : Catch : End Try
                        End If
                        _timer = New Timer(_innerCallback, Me, CLng(MAXTIME * RATIO), Timeout.Infinite)
                    End If
                Else
                    Dim ok As Boolean
                    Try
                        If _timer Is Nothing Then
                            _timer = New Timer(_innerCallback, Me, CLng(dueTime * RATIO), Timeout.Infinite)
                            ok = True
                        Else
                            ok = _timer.Change(CLng(dueTime * RATIO), Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok Then
                        If _timer IsNot Nothing Then
                            Try : _timer.Dispose() : Catch : End Try
                        End If
                        _timer = New Timer(_innerCallback, Me, CLng(dueTime * RATIO), Timeout.Infinite)
                    End If
                End If
            End If
        End Sub

        Public Sub Change(ByVal [date] As Date)
            x = [date]
            isSimple = False
            If Disposed Then
                disposedValue = False
                GC.ReRegisterForFinalize(Me)
            End If
            _needDispose = False
            InnerCallback(_state)
        End Sub

        Public Sub Change(ByVal interval As Long)
            x = Date.Now.AddTicks(interval * 10000L)
            isSimple = False
            If Disposed Then
                disposedValue = False
                GC.ReRegisterForFinalize(Me)
            End If
            _needDispose = False
            InnerCallback(_state)
        End Sub

        Public Sub Close()
            Dispose()
            _callback = Nothing
            _state = Nothing
        End Sub

#Region " IDisposable Support "
        Public ReadOnly Property Disposed() As Boolean
            <DebuggerStepThrough()> _
            Get
                Return disposedValue
            End Get
        End Property

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private disposedValue As Boolean

        Protected Sub Dispose(ByVal disposing As Boolean)
            If Not disposedValue Then
                If _timer IsNot Nothing Then _timer.Dispose() : _timer = Nothing
                If Not disposing Then
                    _callback = Nothing
                    _state = Nothing
                End If
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub
#End Region

    End Class
