    Public NotInheritable Class LongTimer : Implements IDisposable
        Private Const MAXTIME = UInteger.MaxValue - 1L
        Private Const EPSILON = 15L
        Private Const RATIO = 0.8
        Private _timer As Timer
        Private _callback As TimerCallback
        Private ReadOnly _innerCallback As TimerCallback = AddressOf InnerCallback
        Private _state As Object
        Private _needDispose As Boolean

        <DebuggerBrowsable(DebuggerBrowsableState.Never)>
        Private _x As Date
        Private _isSimple As Boolean

        Public ReadOnly Property [Date]() As Date
            <DebuggerStepThrough()>
            Get
                Return _x
            End Get
        End Property

        Public Sub New(callback As TimerCallback, [date] As Date, Optional state As Object = Nothing)
            _x = [date]
            If _x.Kind <> DateTimeKind.Utc Then _x = _x.ToUniversalTime
            _callback = callback
            _state = state
            InnerCallback(_state)
        End Sub

        Public Sub New(callback As TimerCallback, interval As Long, Optional state As Object = Nothing)
            _x = Date.UtcNow.Add(TimeSpan.FromTicks(interval * 10000L))
            _callback = callback
            _state = state
            InnerCallback(_state)
        End Sub

        Private Sub InnerCallback(state As Object)
            'Debug.Print("InnerCallback")
            If _isSimple
Finish:
                _needDispose = True
                Try
                    _callback(If(_state, Me))
                Finally
                    If _needDispose Then Dispose()
                End Try
            Else
                Dim dueTime = (_x - Date.UtcNow).Ticks \ 10000L
                If dueTime <= EPSILON
                    GoTo Finish
                ElseIf dueTime <= 3600 * 1000
                    _isSimple = True
                    Dim ok As Boolean
                    Try
                        If _timer Is Nothing
                            _timer = New Timer(_innerCallback, Me, dueTime, Timeout.Infinite)
                            ok = True
                        Else
                            ok = _timer.Change(dueTime, Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok
                        If _timer IsNot Nothing
                            Try : _timer.Dispose() : Catch : End Try
                        End If
                        _timer = New Timer(_innerCallback, Me, dueTime, Timeout.Infinite)
                    End If
                ElseIf dueTime > MAXTIME
                    Dim ok As Boolean
                    Try
                        If _timer Is Nothing
                            _timer = New Timer(_innerCallback, Me, CLng(MAXTIME * RATIO), Timeout.Infinite)
                            ok = True
                        Else
                            ok = _timer.Change(CLng(MAXTIME * RATIO), Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok
                        If _timer IsNot Nothing
                            Try : _timer.Dispose() : Catch : End Try
                        End If
                        _timer = New Timer(_innerCallback, Me, CLng(MAXTIME * RATIO), Timeout.Infinite)
                    End If
                Else
                    Dim ok As Boolean
                    Try
                        If _timer Is Nothing
                            _timer = New Timer(_innerCallback, Me, CLng(dueTime * RATIO), Timeout.Infinite)
                            ok = True
                        Else
                            ok = _timer.Change(CLng(dueTime * RATIO), Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok
                        If _timer IsNot Nothing
                            Try : _timer.Dispose() : Catch : End Try
                        End If
                        _timer = New Timer(_innerCallback, Me, CLng(dueTime * RATIO), Timeout.Infinite)
                    End If
                End If
            End If
        End Sub

        Public Sub Change([date] As Date)
            _x = [date]
            If _x.Kind <> DateTimeKind.Utc Then _x = _x.ToUniversalTime
            _isSimple = False
            If Disposed
                _disposed = False
                GC.ReRegisterForFinalize(Me)
            End If
            _needDispose = False
            InnerCallback(_state)
        End Sub

        Public Sub Change(interval As Long)
            _x = Date.UtcNow.AddTicks(interval * 10000L)
            _isSimple = False
            If Disposed
                _disposed = False
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
            <DebuggerStepThrough()>
            Get
                Return _disposed
            End Get
        End Property

        <DebuggerBrowsable(DebuggerBrowsableState.Never)>
        Private _disposed As Boolean

        Protected Sub Dispose(disposing As Boolean)
            If Not _disposed
                If _timer IsNot Nothing Then _timer.Dispose() : _timer = Nothing
                If Not disposing
                    _callback = Nothing
                    _state = Nothing
                End If
                _disposed = True
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
