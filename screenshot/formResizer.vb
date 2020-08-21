' 
' I couldn't find any other source for this code than this forum post:
' https://www.vb-paradise.de/index.php/Thread/123876-Windows-Forms-FormBorderStyle-None-Resize/?postID=1075989#post1075989
'

Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Friend Class FormResizer
    Dim TargetForm As Form
    Private _BorderWidth As Integer
    Private _ResizeDirect As resizeDirectection = resizeDirectection.None
    ''' <summary>
    ''' Make borderless Form resizeable.
    ''' </summary>
    ''' <param name="value">As Form</param>
    ''' <param name="borderWidth">(Optional) BorderWidth as Integer (Default = 4)</param>
    Sub New(ByVal value As Form, Optional borderWidth As Integer = 4)
        TargetForm = value
        _BorderWidth = borderWidth
        AddHandler TargetForm.MouseDown, AddressOf Form_MouseDown
        AddHandler TargetForm.MouseMove, AddressOf Form_MouseMove
    End Sub
    Public Enum resizeDirectection
        None = 0
        Left = 1
        TopLeft = 2
        Top = 3
        TopRight = 4
        Right = 5
        BottomRight = 6
        Bottom = 7
        BottomLeft = 8
    End Enum
    Public Property resizeDirect() As resizeDirectection
        Get
            Return _ResizeDirect
        End Get
        Set(ByVal value As resizeDirectection)
            _ResizeDirect = value
            Select Case value
                Case resizeDirectection.Left
                    TargetForm.Cursor = Cursors.SizeWE
                Case resizeDirectection.Right
                    TargetForm.Cursor = Cursors.SizeWE
                Case resizeDirectection.Top
                    TargetForm.Cursor = Cursors.SizeNS
                Case resizeDirectection.Bottom
                    TargetForm.Cursor = Cursors.SizeNS
                Case resizeDirectection.BottomLeft
                    TargetForm.Cursor = Cursors.SizeNESW
                Case resizeDirectection.TopRight
                    TargetForm.Cursor = Cursors.SizeNESW
                Case resizeDirectection.BottomRight
                    TargetForm.Cursor = Cursors.SizeNWSE
                Case resizeDirectection.TopLeft
                    TargetForm.Cursor = Cursors.SizeNWSE
                Case Else
                    TargetForm.Cursor = Cursors.Default
            End Select
        End Set
    End Property
    <DllImport("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTBORDER As Integer = 18
    Private Const HTBOTTOM As Integer = 15
    Private Const HTBOTTOMLEFT As Integer = 16
    Private Const HTBOTTOMRIGHT As Integer = 17
    Private Const HTCAPTION As Integer = 2
    Private Const HTLEFT As Integer = 10
    Private Const HTRIGHT As Integer = 11
    Private Const HTTOP As Integer = 12
    Private Const HTTOPLEFT As Integer = 13
    Private Const HTTOPRIGHT As Integer = 14
    Private Sub ResizeForm(ByVal direction As resizeDirectection)
        Dim dir As Integer = -1
        Select Case direction
            Case resizeDirectection.Left
                dir = HTLEFT
            Case resizeDirectection.TopLeft
                dir = HTTOPLEFT
            Case resizeDirectection.Top
                dir = HTTOP
            Case resizeDirectection.TopRight
                dir = HTTOPRIGHT
            Case resizeDirectection.Right
                dir = HTRIGHT
            Case resizeDirectection.BottomRight
                dir = HTBOTTOMRIGHT
            Case resizeDirectection.Bottom
                dir = HTBOTTOM
            Case resizeDirectection.BottomLeft
                dir = HTBOTTOMLEFT
        End Select
        If dir <> -1 Then
            ReleaseCapture()
            SendMessage(TargetForm.Handle, WM_NCLBUTTONDOWN, dir, 0)
        End If
    End Sub
    Private Sub Form_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Button = Windows.Forms.MouseButtons.Left And TargetForm.WindowState <> FormWindowState.Maximized Then
            ResizeForm(resizeDirect)
        End If
    End Sub
    Private Sub Form_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Location.X < _BorderWidth And e.Location.Y < _BorderWidth Then
            resizeDirect = resizeDirectection.TopLeft
        ElseIf e.Location.X < _BorderWidth And e.Location.Y > TargetForm.Height - _BorderWidth Then
            resizeDirect = resizeDirectection.BottomLeft
        ElseIf e.Location.X > TargetForm.Width - _BorderWidth And e.Location.Y > TargetForm.Height - _BorderWidth Then
            resizeDirect = resizeDirectection.BottomRight
        ElseIf e.Location.X > TargetForm.Width - _BorderWidth And e.Location.Y < _BorderWidth Then
            resizeDirect = resizeDirectection.TopRight
        ElseIf e.Location.X < _BorderWidth Then
            resizeDirect = resizeDirectection.Left
        ElseIf e.Location.X > TargetForm.Width - _BorderWidth Then
            resizeDirect = resizeDirectection.Right
        ElseIf e.Location.Y < _BorderWidth Then
            resizeDirect = resizeDirectection.Top
        ElseIf e.Location.Y > TargetForm.Height - _BorderWidth Then
            resizeDirect = resizeDirectection.Bottom
        Else resizeDirect = resizeDirectection.None
        End If
    End Sub
End Class