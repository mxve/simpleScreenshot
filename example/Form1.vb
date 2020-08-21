Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PictureBox1.Image = simpleScreenshot.selection.captureBmp()
        TextBox1.Text = simpleScreenshot.selection.captureBase64("Now another screenshot to encode as base64")
    End Sub
End Class
