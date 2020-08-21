Imports System.Drawing
Imports System.Windows.Forms

Public Class selection
    Public Shared Function captureBmp(Optional hint As String = "default") As Bitmap
        Dim screenshotDialog As screenshotForm
        If Not hint = "default" Then
            screenshotDialog = New screenshotForm(False, hint)
        Else
            screenshotDialog = New screenshotForm(False)
        End If
        screenshotDialog.ShowDialog()
        If screenshotDialog.DialogResult = DialogResult.OK Then
            Return screenshotDialog.bmp
        End If
        Return Nothing
    End Function

    Public Shared Function captureBase64(Optional hint As String = "default") As String
        Dim screenshotDialog As screenshotForm
        If Not hint = "default" Then
            screenshotDialog = New screenshotForm(True, hint)
        Else
            screenshotDialog = New screenshotForm(True)
        End If
        screenshotDialog.ShowDialog()
        If screenshotDialog.DialogResult = DialogResult.OK Then
            Return screenshotDialog.base64
        End If
        Return Nothing
    End Function
End Class