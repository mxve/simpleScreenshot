Imports simpleScreenshot

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Settings are optional
        PictureBox1.Image = selection.capture()

        Dim settings As misc.Settings
        settings.Opacity = 0.9
        settings.BackColor = Color.Pink
        settings.Size = New Point(600, 800)
        settings.Hint = "bitmap"
        settings.ResizeWidth = 32
        settings.Pen = New Pen(Color.Green, 32)

        PictureBox1.Image = selection.capture(settings)
    End Sub
End Class
