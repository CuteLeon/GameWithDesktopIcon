Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading

Public Class GameForm
    Dim IconSize As Size = New Size(75, 75)
    Dim DesktopIconBound As Integer
    Dim ScreenSize As Size
    Dim DesktopIconWorker As BackgroundWorker = New BackgroundWorker With {.WorkerReportsProgress = False, .WorkerSupportsCancellation = True}

    Private Sub GameForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        AddHandler DesktopIconWorker.DoWork, AddressOf DoWork
    End Sub

    Private Sub DoWork(sender As BackgroundWorker, e As DoWorkEventArgs)
        DesktopIconBound = GetDesktopIconCount() - 1
        ScreenSize = Screen.PrimaryScreen.Bounds.Size
        Dim Index As Integer
        Dim IconPoints(DesktopIconBound) As Point
        For Index = 0 To DesktopIconBound
            IconPoints(Index) = New Point(IconSize.Width * (DesktopIconBound - Index), 0)
            SetIconPosition(Index, IconPoints(Index))
        Next
        IconPoints(0) = GetNextPoint(IconPoints(0))
        Do Until sender.CancellationPending
            SetIconPosition(0, IconPoints(0))
            For Index = DesktopIconBound To 1 Step -1
                IconPoints(Index) = IconPoints(Index - 1)
                SetIconPosition(Index, IconPoints(Index))
            Next
            IconPoints(0) = GetNextPoint(IconPoints(0))
            'Thread.Sleep(100)
        Loop
    End Sub

    Private Sub WorkButton_Click(sender As Object, e As EventArgs) Handles WorkButton.Click
        If WorkButton.Text = "开始" Then
            OpenExplorer()
            DesktopIconWorker.RunWorkerAsync()
            WorkButton.Text = "停止"
        Else
            DesktopIconWorker.CancelAsync()
            CloseExplorer()
            WorkButton.Text = "开始"
        End If
    End Sub

    Private Function GetNextPoint(ByVal IniPoint As Point) As Point
        Static ToLeft, ToUp As Boolean
        Dim X As Integer
        Dim Y As Integer
        If IniPoint.X < 0 Then
            ToLeft = False
        ElseIf IniPoint.X > ScreenSize.Width Then
            ToLeft = True
        End If
        If IniPoint.Y < 0 Then
            ToUp = False
        ElseIf IniPoint.Y > ScreenSize.Height Then
            ToUp = True
        End If
        X = IniPoint.X + IIf(ToLeft, -1, 1) * VBMath.Rnd * IconSize.Width
        Y = IniPoint.Y + IIf(ToUp, -1, 1) * VBMath.Rnd * IconSize.Height
        Return New Point(X, Y)
    End Function
End Class