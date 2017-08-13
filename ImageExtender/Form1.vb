Imports System.IO
Imports System.Drawing
Imports System.Threading

Public Class Form1

    Private Delegate Sub SetProgressBarMaxD(ByRef i As Integer)
    Private Delegate Sub NoParamD()

    Private ProcessThread As Thread

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OFDialg As New OpenFileDialog()
        OFDialg.Filter = "图像文件|*.bmp;*.jpg;*.gpeg;*.png;*.tif;*.tiff"
        OFDialg.Multiselect = False
        Dim OFRet = OFDialg.ShowDialog()
        If (OFRet = DialogResult.OK) Then
            TextBox1.Text = OFDialg.FileName
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Enabled = False
        ProcessThread = New Thread(AddressOf ImageExtender)
        ProcessThread.Start(Val(ComboBox1.SelectedItem))
    End Sub

    Private Sub SetProgressBarMax(ByRef i As Integer)
        ProgressBar1.Maximum = i
        ProgressBar1.Step = 1
        ProgressBar1.Value = 0
    End Sub
    Private Sub ProgressBarDoProgress()
        ProgressBar1.PerformStep()
    End Sub

    Private Sub ProcessFinish()
        Button2.Enabled = True
        MessageBox.Show("扩大完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' 用了简单低效的方法，反正放大两张图片而已。
    ''' </summary>
    ''' <param name="MagT"></param>
    Private Sub ImageExtender(MagT)
        Dim iFile As New FileInfo(TextBox1.Text)
        Dim MagTimes = CInt(MagT)
        Debug.WriteLine(iFile.FullName)
        Dim nFilename As String = iFile.FullName.Replace(iFile.Extension, "-extended" + iFile.Extension)
        If (File.Exists(nFilename)) Then
            nFilename = nFilename.Replace(iFile.Extension, "-" + MagTimes.ToString() + "x-" + Now.ToFileTimeUtc().ToString() + iFile.Extension)
        End If
        Dim OrigBitmap As Bitmap = Image.FromFile(iFile.FullName)
        BeginInvoke(New SetProgressBarMaxD(AddressOf SetProgressBarMax), OrigBitmap.Width)
        Dim NewBitmap = New Bitmap(OrigBitmap.Width * MagTimes, OrigBitmap.Height * MagTimes)
        For x = 0 To OrigBitmap.Width - 1
            For y = 0 To OrigBitmap.Height - 1
                For x1 = 0 To MagTimes - 1
                    For y1 = 0 To MagTimes - 1
                        NewBitmap.SetPixel(x * MagTimes + x1, y * MagTimes + y1, OrigBitmap.GetPixel(x, y))
                    Next
                Next
            Next
            BeginInvoke(New NoParamD(AddressOf ProgressBarDoProgress))
        Next
        NewBitmap.Save(nFilename, OrigBitmap.RawFormat)
        BeginInvoke(New NoParamD(AddressOf ProcessFinish))
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 1
    End Sub
End Class
