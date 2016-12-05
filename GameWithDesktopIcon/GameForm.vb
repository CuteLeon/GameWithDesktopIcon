Imports System.Runtime.InteropServices

Public Class GameForm

    Private Sub GameForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        OpenExplorer()
        Dim DesktopIconCount As Integer = GetDesktopIconCount()
        If (DesktopIconCount > 0) Then
            For Index As Integer = 0 To DesktopIconCount - 1
                With GetIconTextAndPosition(Index)
                    Debug.Print("图标：{0}{1}坐标：{2}，{3}", .Text, vbCrLf, .Position.X, .Position.Y)
                    SetIconPosition(Index, New Point(.Position.X + 100, .Position.Y + 100))
                End With
            Next
        End If
        CloseExplorer()
    End Sub

End Class