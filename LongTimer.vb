    Public NotInheritable Class LongTimer : Implements IDisposable
        Private Const MAXTIME As Int64 = UInt32.MaxValue - 1
        Private Const EPSILON As Int64 = 15
        Private Const RATIO As Double = 0.8
        Private _Timer As System.Threading.Timer
        Private _callback As System.Threading.TimerCallback
        Private _InnerCallback As System.Threading.TimerCallback = AddressOf InnerCallback
        Private _state As Object
        Private _NeedDispose As Boolean

        <DebuggerBrowsable(DebuggerBrowsableState.Never)> _
        Private X As Date
        Private IsSimple As Boolean

        Public ReadOnly Property [Date]() As Date
            <DebuggerStepThrough()> _
            Get
                Return X
            End Get
        End Property

        Public Sub New(ByVal callback As System.Threading.TimerCallback, ByVal [Date] As Date, Optional ByVal State As Object = Nothing)
            X = [Date]
            _callback = callback
            _state = State
            InnerCallback(_state)
        End Sub

        Public Sub New(ByVal callback As System.Threading.TimerCallback, ByVal Interval As Long, Optional ByVal State As Object = Nothing)
            X = Date.Now.Add(TimeSpan.FromTicks(Interval * 10000L))
            _callback = callback
            _state = State
            InnerCallback(_state)
        End Sub

        Private Sub InnerCallback(ByVal state As Object)
            'Debug.Print("InnerCallback")
            If IsSimple Then
Finish:
                _NeedDispose = True
                Try
                    _callback(If(_state Is Nothing, Me, _state))
                Finally
                    If _NeedDispose Then Dispose()
                End Try
            Else
                Dim DueTime = (X - Date.Now).Ticks \ 10000L
                If DueTime <= EPSILON Then
                    GoTo Finish
                ElseIf DueTime <= 3600 * 1000 Then
                    IsSimple = True
                    Dim ok As Boolean
                    Try
                        If _Timer Is Nothing Then
                            _Timer = New System.Threading.Timer(_InnerCallback, Me, DueTime, System.Threading.Timeout.Infinite)
                            ok = True
                        Else
                            ok = _Timer.Change(DueTime, System.Threading.Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok Then
                        If _Timer IsNot Nothing Then
                            Try : _Timer.Dispose() : Catch : End Try
                        End If
                        _Timer = New System.Threading.Timer(_InnerCallback, Me, DueTime, System.Threading.Timeout.Infinite)
                    End If
                ElseIf DueTime > MAXTIME Then
                    Dim ok As Boolean
                    Try
                        If _Timer Is Nothing Then
                            _Timer = New System.Threading.Timer(_InnerCallback, Me, CLng(MAXTIME * RATIO), System.Threading.Timeout.Infinite)
                            ok = True
                        Else
                            ok = _Timer.Change(CLng(MAXTIME * RATIO), System.Threading.Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok Then
                        If _Timer IsNot Nothing Then
                            Try : _Timer.Dispose() : Catch : End Try
                        End If
                        _Timer = New System.Threading.Timer(_InnerCallback, Me, CLng(MAXTIME * RATIO), System.Threading.Timeout.Infinite)
                    End If
                Else
                    Dim ok As Boolean
                    Try
                        If _Timer Is Nothing Then
                            _Timer = New System.Threading.Timer(_InnerCallback, Me, CLng(DueTime * RATIO), System.Threading.Timeout.Infinite)
                            ok = True
                        Else
                            ok = _Timer.Change(CLng(DueTime * RATIO), System.Threading.Timeout.Infinite)
                        End If
                    Catch
                    End Try
                    If Not ok Then
                        If _Timer IsNot Nothing Then
                            Try : _Timer.Dispose() : Catch : End Try
                        End If
                        _Timer = New System.Threading.Timer(_InnerCallback, Me, CLng(DueTime * RATIO), System.Threading.Timeout.Infinite)
                    End If
                End If
            End If
        End Sub

        Public Sub Change(ByVal [Date] As Date)
            X = [Date]
            IsSimple = False
            If Disposed Then
                disposedValue = False
                GC.ReRegisterForFinalize(Me)
            End If
            _NeedDispose = False
            InnerCallback(_state)
        End Sub

        Public Sub Change(ByVal Interval As Long)
            X = Date.Now.AddTicks(Interval * 10000L)
            IsSimple = False
            If Disposed Then
                disposedValue = False
                GC.ReRegisterForFinalize(Me)
            End If
            _NeedDispose = False
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
                If _Timer IsNot Nothing Then _Timer.Dispose() : _Timer = Nothing
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
