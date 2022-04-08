Imports System.IO
Public Class Form1
    'File Operations

    ' Open File 
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog1.Filter = "Rich Text File |*.rtf|Word Doc (*.doc)|*.doc"
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        RichTextBox1.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.RichText)
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SaveFileDialog1.Filter = "Text File|*.rtf"
        SaveFileDialog1.Title = "Save an Text File"
        If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
        End If
    End Sub

    Public Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveFileDialog1.Filter = "RTF File|*.rtf"
        SaveFileDialog1.Title = "Save an Text File"

        If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        SaveAsToolStripMenuItem_Click(sender, e)
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        SaveFileDialog1.Filter = "Text File|*.rtf"
        SaveFileDialog1.Title = "Save an Text File"
        If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then
            My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
        End If
    End Sub

    Private Sub NewWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewWindowToolStripMenuItem.Click
        Dim myForm As New Form1
        myForm.Show()
    End Sub

    'End File Operations

    'Edit Operations
    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click

        If RichTextBox1.Text <> String.Empty Then
            If RichTextBox1.SelectedText = String.Empty Then
                Clipboard.SetText(RichTextBox1.Text)
                RichTextBox1.Text = ""
            Else
                Clipboard.SetText(RichTextBox1.SelectedText)
                RichTextBox1.SelectedText = ""
            End If
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        If RichTextBox1.Text <> String.Empty Then
            If RichTextBox1.SelectedText = String.Empty Then
                Clipboard.SetText(RichTextBox1.Text)
            Else
                Clipboard.SetText(RichTextBox1.SelectedText)
            End If
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
        Dim search_text, temp As String
        Dim index As Integer = 0
        Dim foundornot As Boolean
        temp = RichTextBox1.Text
        RichTextBox1.Text = ""
        RichTextBox1.Text = temp
        search_text = InputBox(
            "Enter Text To Search: ",
            "Search Text",
            "",
            100,
            100
            )
        While index < RichTextBox1.Text.LastIndexOf(search_text)
            RichTextBox1.Find(search_text, index, RichTextBox1.TextLength, RichTextBoxFinds.None)
            RichTextBox1.SelectionBackColor = Color.Red
            index = RichTextBox1.Text.IndexOf(search_text, index) + 1
            If RichTextBox1.Find(search_text, index, RichTextBox1.TextLength, RichTextBoxFinds.None) Then
                foundornot = True
            End If
        End While

        If foundornot = False Then
            MessageBox.Show("Word Not Found")
        End If
    End Sub

    Private Sub ReplaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReplaceToolStripMenuItem.Click
        Dim replace_text, search_text As String
        Dim index As Integer = 0
        search_text = InputBox(
            "Enter Text To Search: ",
            "Search Text",
            "",
            100,
            100
            )
        replace_text = InputBox(
                "Enter Text To Replace With: ",
                "Replace Text",
                "",
                100,
                100
                )
        While index < RichTextBox1.Text.LastIndexOf(search_text)
            RichTextBox1.Find(search_text, 0, RichTextBoxFinds.None)
            RichTextBox1.Focus()
            RichTextBox1.SelectedText = replace_text
        End While
    End Sub

    Private Sub ZoomOutToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ZoomOutToolStripMenuItem.Click
        RichTextBox1.ZoomFactor = RichTextBox1.ZoomFactor + 1
    End Sub

    Private Sub ZoomInToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ZoomInToolStripMenuItem.Click
        If RichTextBox1.ZoomFactor <= 1 Then
            MessageBox.Show("Cannot Zoom Less Than 0", "Cannot Zoom", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            RichTextBox1.ZoomFactor = RichTextBox1.ZoomFactor - 1
        End If
    End Sub


    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub
    'Edit Operations End


    'Font Operations
    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        FontDialog1.ShowDialog()
        If RichTextBox1.Text <> String.Empty Then
            If RichTextBox1.SelectedText = String.Empty Then
                RichTextBox1.Font = FontDialog1.Font
            Else
                RichTextBox1.SelectionFont = FontDialog1.Font
            End If
        End If
    End Sub
    'Font Operation End

    'Context Menu Operation
    Private Sub RedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedToolStripMenuItem.Click
        If RichTextBox1.Text <> String.Empty Then
            If RichTextBox1.SelectedText = String.Empty Then
                RichTextBox1.ForeColor = Color.Black
            Else
                RichTextBox1.SelectionColor = Color.Red
            End If
        End If
    End Sub

    Private Sub GreenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GreenToolStripMenuItem.Click
        If RichTextBox1.Text <> String.Empty Then
            If RichTextBox1.SelectedText = String.Empty Then
                RichTextBox1.ForeColor = Color.Black
            Else
                RichTextBox1.SelectionColor = Color.Green
            End If
        End If
    End Sub

    Private Sub BlueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlueToolStripMenuItem.Click
        If RichTextBox1.Text <> String.Empty Then
            If RichTextBox1.SelectedText = String.Empty Then
                RichTextBox1.ForeColor = Color.Black
            Else
                RichTextBox1.SelectionColor = Color.Blue
            End If
        End If
    End Sub
    'Context Menu Operation End

    'Form Related Functions
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.KeyPreview = True
    End Sub

    'For Save As Ctrl + S
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If (e.Control AndAlso e.KeyCode = Keys.S) Then
            SaveToolStripMenuItem_Click_1(sender, e)
        End If
    End Sub

    'For Closing Dialog
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.FormClosing
        Dim result = MessageBox.Show("Close Window?", "Form Closing", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result = DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub
    'Form Related Functions End
End Class
