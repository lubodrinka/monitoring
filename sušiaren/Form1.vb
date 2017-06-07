Imports System.Text
Imports System.IO.Ports
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Module porting
    Public Declare Function Inp Lib "inpout32.dll" _
    Alias "Inp32" (ByVal PortAddress As Integer) As Integer
End Module
Public Class Form1
    Dim APP As New Excel.Application
    Dim worksheet As Excel.Worksheet
    Dim workbook As Excel.Workbook
    Dim com As String = "COM1"
    Dim com2 As String = "COM2"
    Dim line
    Dim buffer As New StringBuilder
    Dim recpath As String = Application.StartupPath & "\records.txt"
    Dim excelpath As String
    Dim intAddress As Integer
    Dim intReadVal As Integer
    Public Declare Function Inp Lib "inpoutx64.dll" Alias "Inp32" _
    (ByVal PortAddress As Integer) As Integer
    'Declare variables to store read information and Address values
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        'Get the address and assign
        intAddress = 889

        'Read the corresponding value and store
        intReadVal = porting.Inp(intAddress)

        Label1.Text = intReadVal.ToString()
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            SerialPort1.Close()
            com = ComboBox1.Text
            My.Settings.com = ComboBox1.SelectedIndex
            SerialPort1.PortName = com
            SerialPort1.Open()
            checkcomport()
        Catch ex As Exception
            Label1.Text = ex.ToString
        End Try
        Try
            If SerialPort1.CtsHolding = True Then RadioButton1.Checked = True
            If SerialPort1.DtrEnable = True Then CheckBox4.CheckState = CheckState.Checked
        Catch ex As Exception

        End Try
        If SerialPort1.IsOpen = True Then
            Label1.Text = SerialPort1.PortName & " open" & SerialPort1.DsrHolding
        End If

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            SerialPort2.Close()
            com2 = ComboBox2.Text
            My.Settings.com2 = ComboBox1.SelectedIndex
            SerialPort2.PortName = com2
            SerialPort2.Open()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.MaximumSize = New Point(((TableLayoutPanel1.ColumnStyles(1).Width)) / 150 * TableLayoutPanel1.Width, TableLayoutPanel1.RowStyles(0).Height / 2)
        Label1.Size = Label1.MaximumSize
        loadcolmn()
        loadcomtocomboxparameter()
        comboload()
        If My.Settings.delimter <> Nothing Then TextBox1.Text = My.Settings.delimter
        If My.Settings.sendtext <> Nothing Then TextBox2.Text = My.Settings.sendtext
        For Each comstr In My.Computer.Ports.SerialPortNames
            ComboBox1.Items.Add(comstr)
            ComboBox2.Items.Add(comstr)
        Next
        Try
            ComboBox1.SelectedIndex = My.Settings.com
        Catch ex As Exception

        End Try
        Try

            If My.Settings.com <> My.Settings.com2 And My.Settings.com <> ComboBox2.SelectedValue Then
                ComboBox2.SelectedIndex = My.Settings.com2
            End If
        Catch ex As Exception

        End Try
        Try


            If My.Computer.FileSystem.FileExists(recpath) = True Then
                Dim rec() As String = Split(My.Computer.FileSystem.ReadAllText(recpath), vbCrLf)
                For x = 0 To UBound(rec)
                    Dim line As String = rec(x)
                    If line.Contains(vbTab) Then
                        Dim column() As String = Split(line, vbTab)
                        ListView1.Items.Add(column(0))
                        For y = 1 To UBound(column)
                            ListView1.Items(x).SubItems.Add(column(y))

                        Next
                    Else
                        ListView1.Items.Add(rec(x))
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
        Try
            ListView1.Items(ListView1.Items.Count - 1).EnsureVisible()
        Catch ex As Exception

        End Try
        'readThread.Start()
    End Sub
    Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            Dim newstr As New StringBuilder
            For Each li As ListViewItem In ListView1.Items
                For x = 0 To li.SubItems.Count - 1
                    newstr.Append(li.SubItems(x).Text & vbTab)
                Next
                newstr.Append(vbCrLf)
            Next


            If newstr.ToString <> Nothing Then
                My.Computer.FileSystem.WriteAllText(Application.StartupPath & "\records.txt", newstr.ToString, False)
            End If
        Catch ex As Exception

        End Try
        savecolumnnameandwidth()
        combosave()
        Try

            APP.DisplayAlerts = False
            'APP.ActiveWorkbook.Save()
            APP.ActiveWorkbook.Close(SaveChanges:=True)
            APP.Quit()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub text1()
        Timer2.Enabled = True
        Timer2.Start()
        'MsgBox(buffer.ToString)
        Dim excelclupboard As New StringBuilder
        Dim LC As Integer = ListView1.Items.Count
        If buffer.ToString <> Nothing Then
            Dim datenow As String = Date.Now


            'Dim text2 As String = Replace(, vbCr, "")
            Dim text1 As String = buffer.ToString

            Me.Label1.Text = "Data prijaté:  " & datenow & "  " & text1
            ListView1.Items.Add(datenow)
            excelclupboard.Append(datenow & vbTab)
            If text1.Contains(TextBox1.Text) And TextBox1.Text <> String.Empty And TextBox1.Text <> Nothing Then
                Dim lisplited() As String = Split(text1, TextBox1.Text)
                For y = 0 To UBound(lisplited)
                    Dim subitext As String = lisplited(y)
                    If y >= ListView1.Columns.Count Then ListView1.Columns.Add(CStr(ListView1.Columns.Count), CInt(subitext.Length), HorizontalAlignment.Center).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                    ListView1.Items(LC).SubItems.Add(subitext)
                    excelclupboard.Append(subitext & vbTab)
                Next
                excelclupboard.AppendLine(vbCrLf)
            Else
                ListView1.Items(LC).SubItems.Add(text1)
                excelclupboard.Append(text1)
            End If


            Try : My.Computer.Clipboard.SetText(excelclupboard.ToString) : Catch ex As Exception : End Try


            If CheckBox2.CheckState = CheckState.Checked Then

                Me.ListView1.BeginInvoke(New InvokeDelegate3(AddressOf excelonlinewrite))



            End If
            'MsgBox(text1)

            Try
                ListView1.Items(LC).EnsureVisible()

            Catch ex As Exception

            End Try
            buffer.Clear()
        End If
    End Sub
    Delegate Sub InvokeDelegate()
    Dim DTIME As Integer
    
    Private Sub SerialPort1_DataReceived(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Try


            Dim line As String = SerialPort1.ReadLine




            If CheckBox3.CheckState = CheckState.Unchecked Then
                buffer.Append(line)
            ElseIf CheckBox3.CheckState = CheckState.Checked Then
                buffer.AppendLine(SerialPort1.ReadByte)

            End If

            'If Date.Now.Second > DTIME Then
            ListView1.BeginInvoke(New InvokeDelegate(AddressOf text1))
            'End If
            '    DTIME = Date.Now.Second
        Catch ex As Exception

        End Try
    End Sub
    

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        PrintDialog1.AllowSomePages = True
        PrintDialog1.ShowHelp = True
        PrintDialog1.AllowSelection = True
        PrintDialog1.AllowCurrentPage = True
        PrintDialog1.UseEXDialog = True
        PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings

        PrintDocument1.DefaultPageSettings.Margins.Left = 10

        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.DefaultPageSettings.Margins.Left = 10
            PrintDocument1.DefaultPageSettings.Margins.Top = 10
            PrintDocument1.DefaultPageSettings.Margins.Bottom = 15
            PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings

            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
        End If

    End Sub
    Private Sub pdoc_PrintPage3(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static CurrentYPosition As Integer = 0
        If ListView1.View = View.Details Or ListView2.View = View.Details Then
            PrintDetails1(e)
        End If
    End Sub
    Private Sub PrintDetails1(ByRef e As System.Drawing.Printing.PrintPageEventArgs)
        Dim lv As ListView
        If ListView2.Visible = True Then
            lv = ListView2
        Else
            lv = ListView1
        End If
        Static LastIndex As Integer = 0
        Static CurrentPage As Integer = 0
        Dim DpiGraphics As Graphics = Me.CreateGraphics
        Dim DpiX As Integer = DpiGraphics.DpiX
        Dim DpiY As Integer = DpiGraphics.DpiY
        DpiGraphics.Dispose()
        Dim X, Y As Integer
        Dim ImageWidth As Integer
        Dim TextRect As Rectangle = Rectangle.Empty
        Dim TextLeftPad As Single = CSng(4 * (DpiX / 96)) '4 pixel pad on the left.
        Dim ColumnHeaderHeight As Single = CSng(lv.Font.Height + (10 * (DpiX / 96))) '5 pixel pad on the top an bottom
        Dim StringFormat As New StringFormat
        Dim PageNumberWidth As Single = e.Graphics.MeasureString(CStr(CurrentPage), lv.Font).Width
        StringFormat.FormatFlags = StringFormatFlags.NoWrap
        StringFormat.Trimming = StringTrimming.EllipsisCharacter
        StringFormat.LineAlignment = StringAlignment.Center
        CurrentPage += 1
        X = CInt(e.MarginBounds.X)
        Y = CInt(e.MarginBounds.Y)
        For ColumnIndex As Integer = 0 To lv.Columns.Count - 1
            TextRect.X = X
            TextRect.Y = Y
            TextRect.Width = lv.Columns(ColumnIndex).Width
            TextRect.Height = ColumnHeaderHeight
            e.Graphics.FillRectangle(Brushes.LightGray, TextRect)
            e.Graphics.DrawRectangle(Pens.DarkGray, TextRect)
            TextRect.X += TextLeftPad
            TextRect.Width -= TextLeftPad
            e.Graphics.DrawString(lv.Columns(ColumnIndex).Text, lv.Font, Brushes.Black, TextRect, StringFormat)
            X += TextRect.Width + TextLeftPad
        Next
        Y += ColumnHeaderHeight
        For i = LastIndex To lv.Items.Count - 1
            With lv.Items(i)
                X = CInt(e.MarginBounds.X)
                If Y + .Bounds.Height > e.MarginBounds.Bottom Then
                    LastIndex = i - 1
                    e.HasMorePages = True
                    StringFormat.Dispose()
                    e.Graphics.DrawString(CStr(CurrentPage), lv.Font, Brushes.Black, (e.PageBounds.Width - PageNumberWidth) / 2, e.PageBounds.Bottom - lv.Font.Height * 2)
                    Exit Sub
                End If
                ImageWidth = 0
                If lv.SmallImageList IsNot Nothing Then
                    If Not String.IsNullOrEmpty(.ImageKey) Then
                        e.Graphics.DrawImage(lv.SmallImageList.Images(.ImageKey), X, Y)
                    ElseIf .ImageIndex >= 0 Then
                        e.Graphics.DrawImage(lv.SmallImageList.Images(.ImageIndex), X, Y)
                    End If
                    ImageWidth = lv.SmallImageList.ImageSize.Width
                End If
                For ColumnIndex As Integer = 0 To lv.Columns.Count - 1
                    TextRect.X = X
                    TextRect.Y = Y
                    TextRect.Width = lv.Columns(ColumnIndex).Width
                    TextRect.Height = .Bounds.Height
                    Dim bColor As Integer = lv.Items(i).BackColor.ToArgb
                    Dim myBColor As Color = ColorTranslator.FromHtml(bColor)
                    e.Graphics.FillRectangle(New SolidBrush(myBColor), TextRect)
                    If lv.GridLines Then
                        e.Graphics.DrawRectangle(Pens.DarkGray, TextRect)
                    End If
                    If ColumnIndex = 0 Then TextRect.X += ImageWidth
                    TextRect.X += TextLeftPad
                    TextRect.Width -= TextLeftPad
                    If ColumnIndex < .SubItems.Count Then
                        Dim TColor As Integer = lv.Items(i).SubItems(ColumnIndex).ForeColor.ToArgb
                        Dim myTColor As Color = ColorTranslator.FromHtml(TColor)
                        e.Graphics.DrawString(.SubItems(ColumnIndex).Text, lv.Font, New SolidBrush(myTColor), TextRect, StringFormat)
                    End If
                    X += TextRect.Width + TextLeftPad
                Next

                Y += .Bounds.Height
            End With
        Next
        e.Graphics.DrawString(CStr(CurrentPage), lv.Font, Brushes.Black, (e.PageBounds.Width - PageNumberWidth) / 2, e.PageBounds.Bottom - lv.Font.Height * 2)
        StringFormat.Dispose()
        LastIndex = 0
        CurrentPage = 0
    End Sub
    Private Sub SerialPort1_ErrorReceived(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialErrorReceivedEventArgs) Handles SerialPort1.ErrorReceived
        MsgBox("error" & e.ToString, MsgBoxStyle.Exclamation)
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SerialPort1.Close()
        If SerialPort1.IsOpen = False Then Label1.Text = com & "close"
        BackgroundWorker1.CancelAsync()
        Button3.Enabled = True
    End Sub
    Public Sub loadcolmn()
        Try
            If My.Settings.stlpcemeno <> String.Empty Then
                Dim csplit() As String = Split(My.Settings.stlpcemeno, vbCrLf)
                For x = 0 To UBound(csplit)
                    If csplit(x) <> String.Empty Then
                        If csplit(x).Contains(vbTab) Then
                            Dim line() As String = Split(csplit(x), vbTab)
                            Dim cname As String = line(0)
                            Dim cwidth As Integer = line(1)
                            If x < 2 Then
                                ListView1.Columns(x).Text = cname
                                ListView2.Columns(x).Text = cname
                            Else
                                ListView1.Columns.Add(cname).Width = cwidth
                                ListView2.Columns.Add(cname).Width = cwidth
                            End If
                        Else
                            If x < 2 Then ListView1.Columns.Add(csplit(x)) : ListView2.Columns.Add(csplit(x))
                        End If
                    End If
                Next
                If ListView1.Columns.Count > 2 Then
                    ListView1.Columns(1).Width = ListView1.Columns(1).Width - ListView1.Columns(ListView1.Columns.Count - 1).Width
                    ListView2.Columns(1).Width = ListView1.Columns(1).Width
                End If


            End If
        Catch ex As Exception

        End Try
        zmenaširkystlpcov()
    End Sub
    Public Sub zmenaširkystlpcov()
        ListView1.Width = Me.Width - 40
        Dim width As Integer = 0
        For Each licol As ColumnHeader In ListView1.Columns
            width += licol.Width
        Next
        Dim sumwidth As Decimal = (ListView1.Width - 25) / width
        For Each licol As ColumnHeader In ListView1.Columns
            licol.Width = licol.Width * sumwidth
        Next

        Dim sumwidth2 As Decimal = (ListView2.Width - 25) / width
        For Each licol As ColumnHeader In ListView2.Columns
            licol.Width = licol.Width * sumwidth
        Next
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try : SerialPort1.Open() : Catch ex As Exception : End Try
        Try : SerialPort2.Open() : Catch ex As Exception : End Try
        Button3.Enabled = False
    End Sub
    Private Sub Form1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        zmenaširkystlpcov()
        Me.ListView1.Height = Me.Height - TableLayoutPanel1.Height - 100

    End Sub
    Public Sub savecolumnnameandwidth()
        Try
            Dim newstr As New StringBuilder
            For Each licol As ColumnHeader In ListView1.Columns
                newstr.Append(licol.Text & vbTab & licol.Width & vbCrLf)
            Next
            My.Settings.stlpcemeno = newstr.ToString
        Catch ex As Exception
        End Try
    End Sub
    Private Sub ListView1_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        Dim ci As Integer = e.Column.ToString
        ListView1.Columns(ci).Text = InputBox("Newname " & ListView1.Columns(ci).Text)
        ListView2.Columns(ci).Text = ListView1.Columns(ci).Text
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim menostlpca As String = InputBox("Name?")
        ListView1.Columns.Add(menostlpca)
        ListView2.Columns.Add(menostlpca)
        My.Settings.stlpcemeno = My.Settings.stlpcemeno & menostlpca & vbTab & ListView1.Columns(ListView1.Columns.Count - 1).Width & vbCrLf
    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        My.Settings.stlpcemeno = String.Empty
        For Each licol As ColumnHeader In ListView1.Columns
            If licol.Index > 1 Then ListView1.Columns.Remove(licol)
        Next
        For Each licol As ColumnHeader In ListView2.Columns
            If licol.Index > 1 Then ListView1.Columns.Remove(licol)
        Next
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> String.Empty Then
            My.Settings.delimter = TextBox1.Text
        Else
            My.Settings.delimter = Nothing
        End If
    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        My.Settings.sendtext = TextBox2.Text
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            SerialPort2.WriteLine(TextBox2.Text)
            If CheckBox1.Checked Then Timer1.Start() : Timer1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            SerialPort2.WriteLine(TextBox2.Text)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Timer1.Stop() : Timer1.Enabled = False
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        BackgroundWorker2.RunWorkerAsync()
    End Sub
    Public Sub savetoexcelfile()
        Dim lv As ListView = ListView1
        If ListView2.Visible = True Then lv = ListView2
        Try
            ProgressBar1.Maximum = (lv.Columns.Count * lv.Items.Count) : ProgressBar1.Visible = True : ProgressBar1.Value = 0
            Dim x As Integer = 0 : Dim r As Integer = 1 : lv.UseWaitCursor = True
            SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
            SaveFileDialog1.DefaultExt = ".xls"
            SaveFileDialog1.FileName = Format(Date.Now.Date, "dd/MM/yyyy ")
            SaveFileDialog1.Filter = "excel files (*.xls)|.xls|*.xlsx)|.xlsx|All files (*.*)|*.*"
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim excelfilepat As String = SaveFileDialog1.FileName
                Dim fileExists1 As Boolean = My.Computer.FileSystem.FileExists(excelfilepat)
                If fileExists1 = False Then
                    workbook = APP.Workbooks.Add(System.Reflection.Missing.Value) : workbook.UnprotectSharing()
                    APP.ActiveWorkbook.SaveAs(Filename:=excelfilepat) : APP.ActiveWorkbook.Close()
                    APP.Quit()
                End If
                workbook = APP.Workbooks.Open(excelfilepat)
                worksheet = workbook.Sheets.Item(1)
                Dim bColor As Integer = 2
                If lv.Visible = True Then
                    For Each col As ColumnHeader In lv.Columns
                        x += 1
                        worksheet.Cells(1, x).value = col.Text
                        worksheet.Cells(1, x).Interior.Color = Color.LightGray
                        worksheet.Cells(r, x).BorderAround(1, 2, 4105, Color.Black)
                    Next
                    For Each item As ListViewItem In lv.Items
                        r += 1
                        Try
                            ProgressBar1.Value = (r * x)
                        Catch ex As Exception
                        End Try
                        Try
                            For count = 0 To item.SubItems.Count - 1
                                worksheet.Cells(r, count + 1).Value = item.SubItems(count).Text
                                For xx = 0 To lv.Columns.Count - 1
                                    bColor = item.BackColor.ToArgb
                                    Dim myBColor As Color = ColorTranslator.FromHtml(bColor)
                                    worksheet.Cells(r, xx + 1).Interior.Color = myBColor
                                    worksheet.Cells(r, xx + 1).BorderAround(1, 2, 4105, Color.Black)
                                Next
                            Next
                        Catch ex As Exception

                        End Try
                    Next

                End If
                Try
                    worksheet.Columns("A:Q").EntireColumn.Autofit()
                    Dim name As String = IO.Path.GetFullPath(SaveFileDialog1.FileName)
                    APP.DisplayAlerts = False : APP.ActiveWorkbook.SaveAs(Filename:=name) : workbook.Close()
                    APP.Quit()
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception
            APP.DisplayAlerts = False : APP.Quit() : MsgBox(ex.ToString)
        End Try
        lv.UseWaitCursor = False : ProgressBar1.Visible = False : BackgroundWorker2.CancelAsync()
    End Sub
    Delegate Sub InvokeDelegate2()
    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Invoke(New InvokeDelegate2(AddressOf savetoexcelfile))
    End Sub
    Public Sub excelonlinewrite()
        Try
            If CheckBox2.CheckState = CheckState.Checked Then
                Dim LastRow As Long
                With worksheet
                    LastRow = .Cells(.Rows.Count, 2).End(-4162).Row
                End With
                Dim licol As ListViewItem = ListView1.Items(ListView1.Items.Count - 1)

                For x = 0 To licol.SubItems.Count - 1
                    worksheet.Cells(LastRow + 1, x + 1).Value = licol.SubItems(x).Text
                Next

                worksheet.Columns("A:Q").EntireColumn.Autofit()
            End If
        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try
        MsgBox(excelpath)
    End Sub
    Delegate Sub InvokeDelegate3()




    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
        SaveFileDialog1.DefaultExt = ".xls"

        SaveFileDialog1.FileName = Format(Date.Now.Date, "dd/MM/yyyy ")
        SaveFileDialog1.Filter = "excel files (*.xls)|*.xlsx|All files (*.*)|*.*"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            My.Settings.excelfilepath = SaveFileDialog1.FileName
        End If
        excelpath = My.Settings.excelfilepath
        excelopenfile()
    End Sub
    Public Sub excelopenfile()
        If My.Computer.FileSystem.FileExists(My.Settings.excelfilepath) = False Then
            excelpath = Application.StartupPath & "\book1.xls"

            Dim fileExists1 As Boolean = My.Computer.FileSystem.FileExists(excelpath)
            If fileExists1 = False Then

                workbook = APP.Workbooks.Add(System.Reflection.Missing.Value)
                workbook.UnprotectSharing()
                APP.ActiveWorkbook.SaveAs(Filename:=excelpath)
                APP.ActiveWorkbook.Close()
                APP.Quit()
            End If
        ElseIf My.Computer.FileSystem.FileExists(My.Settings.excelfilepath) = True Then
            excelpath = My.Settings.excelfilepath
        End If

        workbook = APP.Workbooks.Open(excelpath)

        worksheet = workbook.Sheets.Item(1)
    End Sub
    Dim datum, datum1, datum2, timestr As Date
    Public Sub filterdatum()
        ListView2.Items.Clear()
        ListView2.Visible = True : ListView2.BringToFront()
        Dim dat1, dat2 As Integer
        For Each li As ListViewItem In ListView1.Items
            If IsDate(li.Text) Then
                datum = FormatDateTime(li.Text, DateFormat.ShortDate)
                datum1 = FormatDateTime(DateTimePicker1.Value, DateFormat.ShortDate)
                datum2 = FormatDateTime(DateTimePicker2.Value, DateFormat.ShortDate)
                timestr = FormatDateTime(li.Text, DateFormat.ShortTime)
                dat1 = DateDiff(DateInterval.Day, datum1, datum)
                dat2 = DateDiff(DateInterval.Day, datum, datum2)
                'MsgBox(datum & vbCrLf & datum1 & vbTab & dat1 & vbCrLf & datum2 & vbTab & dat2)
                If dat1 >= 0 AndAlso dat2 >= 0 Then
                    ListView2.Items.Add(li.Text)
                    For x = 1 To li.SubItems.Count - 1
                        ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(li.SubItems(x).Text)
                    Next
                End If
            End If
        Next
    End Sub
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        filterdatum()
    End Sub
    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        filterdatum()
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        ListView2.Visible = False : ListView2.SendToBack()
    End Sub
    Private Sub ListView1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.Resize
        ListView2.Size = New Size(ListView1.Size)
        ListView2.Location = New Point(ListView1.Location)
    End Sub
    Private Sub CheckBox2_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.MouseHover
        ToolTip1.Active = True
        ToolTip1.SetToolTip(CheckBox2, excelpath)
    End Sub
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.CheckState = CheckState.Checked Then
            excelopenfile()
        ElseIf CheckBox2.CheckState = CheckState.Unchecked Then
            Try

                APP.DisplayAlerts = False
                'APP.ActiveWorkbook.Save()
                APP.ActiveWorkbook.Close(SaveChanges:=True)
                APP.Quit()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim excelclupboard As New StringBuilder
        For Each li As ListViewItem In ListView1.SelectedItems
            For x = 0 To li.SubItems.Count - 1
                excelclupboard.Append(li.SubItems(x).Text & vbTab)
            Next
            excelclupboard.AppendLine(vbCrLf)
        Next

        If excelclupboard.ToString = Nothing Then
            For Each li As ListViewItem In ListView1.Items
                For x = 0 To li.SubItems.Count - 1
                    excelclupboard.Append(li.SubItems(x).Text & vbTab)
                Next
                excelclupboard.AppendLine(vbCrLf)
            Next
        End If

        My.Computer.Clipboard.SetText(excelclupboard.ToString)
    End Sub
    Public Sub loadcomtocomboxparameter()
        Dim br As String = "75,110,134,150,300,600,1200,1800,2400,4800,7200,9600,14400,19200,38400,57600,115200,128000"
        For Each baudrate As String In br.Split(",")
            ComboBox3.Items.Add(baudrate)
            If baudrate = SerialPort1.BaudRate.ToString Then ComboBox3.SelectedIndex = ComboBox3.Items.Count - 1
        Next
        Dim DBits As String = "5,6,7,8"
        For Each DataBits As String In DBits.Split(",")
            ComboBox4.Items.Add(DataBits)
            If DataBits = SerialPort1.DataBits.ToString Then ComboBox4.SelectedIndex = ComboBox4.Items.Count - 1
        Next

        Dim stopBitsstr As String = StopBits.One.ToString & "," & StopBits.OnePointFive.ToString & "," & StopBits.Two.ToString
        For Each stopBits As String In stopBitsstr.Split(",")
            ComboBox5.Items.Add(stopBits)
            If stopBits = SerialPort1.StopBits.ToString Then ComboBox5.SelectedIndex = ComboBox5.Items.Count - 1
        Next

        Dim Paritystr As String = Parity.None.ToString & "," & Parity.Odd.ToString & "," & Parity.Even.ToString & "," & Parity.Mark.ToString & "," & Parity.Space.ToString
        For Each Parity As String In Paritystr.Split(",")
            ComboBox6.Items.Add(Parity)
            If Parity = SerialPort1.Parity.ToString Then ComboBox6.SelectedIndex = ComboBox6.Items.Count - 1
        Next

        Dim handshakestr As String = Handshake.None.ToString & "," & Handshake.XOnXOff.ToString & "," & Handshake.RequestToSend.ToString & "," & Handshake.RequestToSendXOnXOff.ToString
        For Each handshake As String In handshakestr.Split(",")
            ComboBox7.Items.Add(handshake)
            If handshake = SerialPort1.Handshake.ToString Then ComboBox7.SelectedIndex = ComboBox7.Items.Count - 1
        Next
        Dim encodingstr As String = "ASCII,Unicode,UTF32,UTF7,UTF8,BigEndianUnicode,Default"
        For Each endcoding As String In encodingstr.Split(",")
            ComboBox8.Items.Add(endcoding)
            If SerialPort1.Encoding.ToString.Contains(endcoding) Then ComboBox8.SelectedIndex = ComboBox8.Items.Count - 1
        Next
    End Sub


    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.CheckState = CheckState.Checked Then SerialPort1.DtrEnable = True
    End Sub
    Dim baudrate_ob As Object
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

        If IsNumeric(ComboBox3.SelectedItem) Then

            If SerialPort1.BaudRate <> ComboBox3.SelectedItem Then SerialPort1.BaudRate = ComboBox3.SelectedItem
        End If

        baudrate_ob = SerialPort1.BaudRate
        Label30.Text = baudrate_ob.ToString

    End Sub

    Dim databit_ob As Object
    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        'datbits
        If IsNumeric(ComboBox4.SelectedItem) Then

            If SerialPort1.DataBits <> ComboBox4.SelectedItem Then SerialPort1.DataBits = ComboBox4.SelectedItem
        End If
        databit_ob = SerialPort1.DataBits
        Label40.Text = databit_ob.ToString
    End Sub
    Dim stopbit_ob As Object
    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged

        If ComboBox5.SelectedIndex = 0 Then
            SerialPort1.StopBits = StopBits.One
        ElseIf ComboBox5.SelectedIndex = 1 Then
            SerialPort1.StopBits = StopBits.OnePointFive
        ElseIf ComboBox5.SelectedIndex = 2 Then
            SerialPort1.StopBits = StopBits.Two
        End If
        stopbit_ob = SerialPort1.StopBits
        Label50.Text = stopbit_ob.ToString
    End Sub
    Dim parity_ob As Object
    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        'parity.None &  parity.Odd. & parity.Even & parity.Mark & parity.Space
        If ComboBox6.SelectedIndex = 0 Then
            SerialPort1.Parity = Parity.None
        ElseIf ComboBox6.SelectedIndex = 1 Then
            SerialPort1.Parity = Parity.Odd
        ElseIf ComboBox6.SelectedIndex = 2 Then
            SerialPort1.Parity = Parity.Even
        ElseIf ComboBox6.SelectedIndex = 3 Then
            SerialPort1.Parity = Parity.Mark
        ElseIf ComboBox6.SelectedIndex = 4 Then
            SerialPort1.Parity = Parity.Space
        End If
        parity_ob = SerialPort1.Parity
        Label60.Text = parity_ob.ToString
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        If ComboBox7.SelectedIndex = 0 Then
            SerialPort1.Handshake = Handshake.None
        ElseIf ComboBox7.SelectedIndex = 1 Then
            SerialPort1.Handshake = Handshake.XOnXOff
        ElseIf ComboBox7.SelectedIndex = 2 Then
            SerialPort1.Handshake = Handshake.RequestToSend
        ElseIf ComboBox7.SelectedIndex = 3 Then
            SerialPort1.Handshake = Handshake.RequestToSendXOnXOff

        End If
        Label70.Text = SerialPort1.Handshake.ToString
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        Try


            'Dim encodingstr As String = "ASCII,Unicode,UTF32,UTF7,UTF8,BigEndianUnicode,Default"
            If ComboBox8.SelectedItem <> Nothing Then
                If ComboBox8.SelectedIndex = 0 Then
                    SerialPort1.Encoding = ASCIIEncoding.ASCII
                ElseIf ComboBox8.SelectedIndex = 1 Then
                    SerialPort1.Encoding = ASCIIEncoding.Unicode
                ElseIf ComboBox8.SelectedIndex = 2 Then
                    SerialPort1.Encoding = ASCIIEncoding.UTF32
                ElseIf ComboBox8.SelectedIndex = 3 Then
                    SerialPort1.Encoding = ASCIIEncoding.UTF7
                ElseIf ComboBox8.SelectedIndex = 4 Then
                    SerialPort1.Encoding = ASCIIEncoding.UTF8
                ElseIf ComboBox8.SelectedIndex = 5 Then
                    SerialPort1.Encoding = ASCIIEncoding.BigEndianUnicode
                ElseIf ComboBox8.SelectedIndex = 6 Then
                    SerialPort1.Encoding = ASCIIEncoding.Default
                End If

            End If
            Label80.Text = Replace(Replace(SerialPort1.Encoding.ToString, "System.Text.", ""), "Encoding", "")
        Catch ex As Exception
            Label80.Text = ex.ToString
        End Try

    End Sub
    Public Sub checkcomport()
        Try

            Label30.Text = My.Computer.Ports.OpenSerialPort(com).BaudRate.ToString
            Label40.Text = My.Computer.Ports.OpenSerialPort(com).DataBits.ToString
            Label50.Text = My.Computer.Ports.OpenSerialPort(com).StopBits.ToString
            Label60.Text = My.Computer.Ports.OpenSerialPort(com).Parity.ToString
            Label70.Text = My.Computer.Ports.OpenSerialPort(com).Handshake.ToString
            Label80.Text = Replace(Replace(My.Computer.Ports.OpenSerialPort(com).Encoding.ToString, "System.Text", ""), "Encoding", "")


        Catch ex As Exception
            Label1.Text = ex.ToString
        End Try
    End Sub

    Private Sub ComboBox1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox1.MouseClick
        For Each comstr In My.Computer.Ports.SerialPortNames
            If ComboBox1.Items.Contains(comstr) = False Then
                ComboBox1.Items.Add(comstr)
            End If

        Next

    End Sub

    Private Sub ComboBox2_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComboBox2.MouseClick
        For Each comstr In My.Computer.Ports.SerialPortNames
            If ComboBox2.Items.Contains(comstr) = False Then
                ComboBox2.Items.Add(comstr)
            End If

        Next
    End Sub

    Dim pin As String
    Private Sub SerialPort1_PinChanged(ByVal sender As System.Object, ByVal e As System.IO.Ports.SerialPinChangedEventArgs) Handles SerialPort1.PinChanged
        pin = e.EventType.ToString
        If BackgroundWorker4.IsBusy = False Then
            BackgroundWorker4.RunWorkerAsync()
        End If

    End Sub
    Public Sub pinchanged()
        Label9.Text = pin
    End Sub

    Delegate Sub InvokeDelegate4()

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Label9.BeginInvoke(New InvokeDelegate4(AddressOf pinchanged))

    End Sub
    Public Sub combosave()
        'baudrate
        My.Settings.cb3 = ComboBox3.SelectedIndex
        'databits()
        My.Settings.cb4 = ComboBox4.SelectedIndex
        'StopBits()
        My.Settings.cb5 = ComboBox5.SelectedIndex
        'Parity
        My.Settings.cb6 = ComboBox6.SelectedIndex
        'handshake
        My.Settings.cb7 = ComboBox7.SelectedIndex
        'ASCIIEncoding
        My.Settings.cb8 = ComboBox8.SelectedIndex

    End Sub
    Public Sub comboload()
        If My.Settings.cb3 <> Nothing Then ComboBox3.SelectedIndex = CInt(My.Settings.cb3)
        If My.Settings.cb4 <> Nothing Then ComboBox4.SelectedIndex = CInt(My.Settings.cb4)
        If My.Settings.cb5 <> Nothing Then ComboBox5.SelectedIndex = CInt(My.Settings.cb5)
        If My.Settings.cb6 <> Nothing Then ComboBox6.SelectedIndex = CInt(My.Settings.cb6)
        If My.Settings.cb7 <> Nothing Then ComboBox7.SelectedIndex = CInt(My.Settings.cb7)
        If My.Settings.cb8 <> Nothing Then ComboBox8.SelectedIndex = CInt(My.Settings.cb8)
        CheckBox3.Checked = My.Settings.bytesenabled






    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        My.Settings.bytesenabled = CheckBox3.Checked
    End Sub



    
End Class
