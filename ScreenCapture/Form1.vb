Public Class Form1

#Region "宣告變數"
    Private s As New Screen
    Private HotKey As New CHotKey(Me)
    Private imageFileName As String
    Private imageFormat As Imaging.ImageFormat

    Private Enum modKey
        WM_HOTKEY = &H312
        NONE = &H0
        ALT = &H1
        CONTROL = &H2
        CTRL_ALT = &H3
        SHIFT = &H4
        ALT_SHIFT = &H5
        CTRL_SHIFT = &H6
        CTRL_SHIFT_ALT = &H7
        GWL_WNDPROC = (-4)
    End Enum
#End Region

    '覆載 WndProC
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = modKey.WM_HOTKEY Then
            ScreenCapture()
            Me.Show()
            Me.WindowState = FormWindowState.Normal
        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox1.Visible = False
        SplitContainer1.Panel2.BackgroundImage = My.Resources.Resource.backgroundImage
        CheckBox1.Checked = True
        PictureBox1.Enabled = False
        RadioButton1.Checked = True
        TextBox1.Text = Environ("SystemDrive") & Environ("HomePath") & "\Desktop"
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        s.Close()
        HotKey.Close()
        NotifyIcon1.Dispose()
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PictureBox1.Size = New Size(SplitContainer1.Panel2.Size.Width, SplitContainer1.Panel2.Size.Height)
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
        End If
    End Sub

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        s.X = e.X
        s.Y = e.Y
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If e.Button = MouseButtons.Left Then
            Using g As Graphics = PictureBox1.CreateGraphics
                Using p As New Pen(Color.Black, 1)
                    p.DashStyle = Drawing2D.DashStyle.Dash
                    g.DrawImage(s.screenBMP, 0, 0, s.screenBMP.Width, s.screenBMP.Height)
                    g.DrawRectangle(p, s.MakeRect(e.X, e.Y))
                End Using
            End Using
        End If
    End Sub

    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        Dim srcRect As Rectangle = s.MakeRect(e.X, e.Y)
        s.selectBMP = New Bitmap(srcRect.Width, srcRect.Height)

        Using g As Graphics = Graphics.FromImage(s.selectBMP)
            g.DrawImage(s.screenBMP, 0, 0, srcRect, GraphicsUnit.Pixel)
            PictureBox1.Image = s.selectBMP
        End Using

        SaveFile()
        PictureBox1.Enabled = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            imageFormat = Imaging.ImageFormat.Png
            imageFileName = ".png"
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            imageFormat = Imaging.ImageFormat.Jpeg
            imageFileName = ".jpg"
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            imageFormat = Imaging.ImageFormat.Bmp
            imageFileName = ".bmp"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub NotifyIcon1_Click(sender As Object, e As EventArgs) Handles NotifyIcon1.Click
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub

    ' 熱鍵
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            HotKey.RegisterKey(0, modKey.NONE, Keys.PrintScreen)
        End If

        If CheckBox1.Checked = False Then
            HotKey.UnRegisterKey(0)
        End If
    End Sub

    ''' <summary>
    ''' 截取螢幕畫面
    ''' </summary>
    ''' <returns>傳回 Bitmap 格式</returns>
    Private Function GetScreen() As Bitmap
        Dim screenWidth As Integer = My.Computer.Screen.Bounds.Width
        Dim ScreenHeight As Integer = My.Computer.Screen.Bounds.Height
        Dim screenBitMAP As New Bitmap(screenWidth, ScreenHeight)

        Using g As Graphics = Graphics.FromImage(screenBitMAP)
            g.CopyFromScreen(New Point(0, 0), New Point(0, 0), New Size(screenWidth, ScreenHeight))
        End Using

        Return screenBitMAP
    End Function

    ''' <summary>
    ''' 截取螢幕畫面後，設定為 PictureBox 的圖片
    ''' </summary>
    Private Sub ScreenCapture()
        s.screenBMP = New Bitmap(GetScreen)
        PictureBox1.Image = s.screenBMP
        PictureBox1.Visible = True
        PictureBox1.Enabled = True
    End Sub

    Private Function CheckChar(ByVal tmpString As String) As String
        If Strings.Right(tmpString, 1) <> "\" Then
            Return tmpString & "\"
        End If

        Return tmpString
    End Function

    Private Sub SaveFile()
        If My.Computer.FileSystem.DirectoryExists(TextBox1.Text) Then
            s.selectBMP.Save(CheckChar(TextBox1.Text) & Format(Now, "yyyyMMdd-HH.mm.ss") & imageFileName, imageFormat)
        End If
        My.Computer.Clipboard.SetImage(s.selectBMP)
    End Sub
End Class