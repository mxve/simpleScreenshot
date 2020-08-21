Imports System.Drawing
Imports System.Windows.Forms

Public Class misc
    Public Enum screenshotFormat
        Bitmap
        Base64
    End Enum

    Public Structure Settings
        Public Shared Opacity As Single = 0.7
        Public Shared Pen As Pen = New Pen(Color.FromArgb(190, 100, 100, 100), 2)
        Public Shared Size As Point = New Point(450, 300)
        Public Shared BackColor As Color = Color.LightGreen
        Public Shared DashStyle As Drawing2D.DashStyle = Drawing2D.DashStyle.Dot
        Public Shared Hint As String = "The green area marks the screenshot." + vbNewLine + vbNewLine + "Press <ENTER> to capture"
        Public Shared ResizeWidth As Integer = 6
        Public Shared Format As screenshotFormat = screenshotFormat.Bitmap
    End Structure
End Class

Public Class selection
    Private Shared Function _capture(settings As misc.Settings)
        Dim screenshotDialog As screenshotForm
        screenshotDialog = New screenshotForm(settings)
        screenshotDialog.ShowDialog()
        If screenshotDialog.DialogResult = DialogResult.OK Then
            Select Case settings.Format
                Case misc.screenshotFormat.Bitmap
                    Return screenshotDialog.bmp
                Case misc.screenshotFormat.Base64
                    Return screenshotDialog.base64
            End Select
        End If
        Return Nothing
    End Function

    Public Shared Function capture()
        Dim settings As New misc.Settings
        Return _capture(settings)
    End Function

    Public Shared Function capture(settings As misc.Settings)
        Return _capture(settings)
    End Function
End Class