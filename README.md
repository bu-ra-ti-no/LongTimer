# LongTimer
vb class that implements timer with support for long time intervals.

Looks like [System.Threading.Timer] (https://msdn.microsoft.com/en-us/library/system.threading.timer.aspx)
***
**VB.NET**
```vb
Private Shared ReadOnly timerBeginCallbackDelegate As Threading.TimerCallback = AddressOf TimerBeginCallback

. . .

timerBegin = New LongTimer(timerBeginCallbackDelegate, startDate, State:=Me)
. . . 
timerBegin.Change(New Date(2016, 8, 11, 15, 0, 0))
```
***
**C\#**
```c#
private static readonly System.Threading.TimerCallback timerBeginCallbackDelegate = TimerBeginCallback;

. . .

timerBegin = new LongTimer(timerBeginCallbackDelegate, startDate, this);
. . . 
timerBegin.Change(new System.DateTime(2016, 8, 11, 15, 0, 0));
```
