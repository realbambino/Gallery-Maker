VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gallery Maker"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdTable 
      Caption         =   "&Tabled"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame frmBrowse 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5055
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4815
      End
      Begin VB.FileListBox flFile 
         Height          =   2235
         Left            =   120
         Pattern         =   "*.jpg;*.jpeg;*.bmp;*.png;*.jpe"
         TabIndex        =   5
         Top             =   2160
         Width           =   4815
      End
      Begin VB.DirListBox drvDrive 
         Height          =   1440
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label lblFileFound 
         AutoSize        =   -1  'True
         Caption         =   "File(s) found:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   4560
         Width           =   900
      End
   End
   Begin VB.TextBox txtDocName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "Untitled"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblASCII 
      AutoSize        =   -1  'True
      Caption         =   "ASCII codes"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   885
   End
   Begin VB.Label lblNote 
      Caption         =   "NOTE: Please make sure that the ""THUMB"" folder with the thumbnails already exist. Double-click on the filename to open the file."
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   5055
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Document &Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API declare for enabling manifest bundled in resource
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub cmdCreate_Click()
On Error Resume Next
Dim OutputText As String

' Check for a valid document name
If txtDocName = "" Then
    MsgBox "Please input a document name.", vbInformation, App.ProductName
    txtDocName.SetFocus
    Exit Sub
End If

' Check for total files found
If flFile.ListCount = 0 Then
    MsgBox "Unable to perform operation. No file(s) found.", vbInformation, App.ProductName
    Exit Sub
End If

    ' Remove all text found on the Output textfield
    frmOutput.txtOutput.Text = ""
    
    OutputText = "<html>" & vbCrLf & vbCrLf
    OutputText = OutputText & "<!--" & vbCrLf
    OutputText = OutputText & " Created using Gallery Maker." & vbCrLf
    OutputText = OutputText & " Date of creation: " & Format$(Now(), "Long Date") & vbCrLf
    OutputText = OutputText & " Time of creation: " & Time$ & vbCrLf
    OutputText = OutputText & "-->" & vbCrLf & vbCrLf
    
    OutputText = OutputText & "   <body bgcolor=""#6B6B6B""" & "></body>" & vbCrLf
    
    OutputText = OutputText & "   <title>" & txtDocName.Text & "</title>" & vbCrLf
    OutputText = OutputText & "   <center><h3>Created using Gallery Maker</h3>" & vbCrLf
    OutputText = OutputText & "   Total picture(s) on current gallery: " & flFile.ListCount & ".</center>" & vbCrLf
    OutputText = OutputText & "   <hr><br />" & vbCrLf
    
    flFile.SetFocus
    flFile.ListIndex = 0
    flFile.Enabled = False
    
    For X = 1 To (flFile.ListCount)
        OutputText = OutputText & "   <a href=""" & flFile.Filename & """><img src=""thumb/" & flFile.Filename & """></a> " & vbCrLf
        flFile.ListIndex = X
        DoEvents
    Next X
    
    ' Adding some information
    OutputText = OutputText & "   <hr><br />" & vbCrLf
    OutputText = OutputText & "   <center>Copyright &copy; 2001-" & Format$(Date, "yyyy") & " Ino Bambino. All rights reserved.</center>" & vbCrLf
    
    ' Finalizing the string
    OutputText = OutputText & "</html>" & vbCrLf
    
    flFile.Enabled = True
    
    ' Placing the string on to the textfield
    frmOutput.txtOutput.Text = OutputText
    frmOutput.Show
    
End Sub

Private Sub cmdExit_Click()
    Form_Unload 0
End Sub

Private Sub cmdTable_Click()
On Error Resume Next
Dim OutputText As String
Dim TotalGallerySize As Double
Dim CurrentFileSize As Double

' Check for a valid document name
If txtDocName = "" Then
    MsgBox "Please input a document name.", vbInformation, App.ProductName
    txtDocName.SetFocus
    Exit Sub
End If

' Check for total files found
If flFile.ListCount = 0 Then
    MsgBox "Unable to perform operation. No file(s) found.", vbInformation, App.ProductName
    Exit Sub
End If

    ' Remove all text found on the Output textfield
    OutputText = ""
    
    ' <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    
    OutputText = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbCrLf & vbCrLf
    OutputText = OutputText & "<html>" & vbCrLf & vbCrLf
    OutputText = OutputText & "<!--" & vbCrLf
    OutputText = OutputText & " Created using Gallery Maker." & vbCrLf
    OutputText = OutputText & " Version number: " & App.Major & "." & App.Minor & " (BUILD " & App.Revision & ")" & vbCrLf
    OutputText = OutputText & " Date of creation: " & Format$(Now(), "Long Date") & vbCrLf
    OutputText = OutputText & " Time of creation: " & Time$ & vbCrLf

    OutputText = OutputText & "-->" & vbCrLf & vbCrLf
    
    OutputText = OutputText & "<head>" & vbCrLf
    OutputText = OutputText & "   <style type=""text/css"">" & vbCrLf & _
                              "      a {text-decoration: none; font-family: Tahoma, Verdana, sans-serif;}" & vbCrLf & _
                              "      a:hover {text-decoration: underline; color: #C0C0C0;}" & vbCrLf & _
                              "      body {font-family: Tahoma, Verdana, sans-serif;}" & vbCrLf & _
                              "   </style>" & vbCrLf & _
                              vbCrLf
    OutputText = OutputText & "   <title>" & txtDocName.Text & "</title>" & vbCrLf
    OutputText = OutputText & "<head>" & vbCrLf & vbCrLf
    
    OutputText = OutputText & "   <body bgcolor=""#6B6B6B""" & "></body>" & vbCrLf
    
    OutputText = OutputText & "   <center><b>Created using " & App.ProductName & "</b><br /><br />" & vbCrLf
    OutputText = OutputText & "   Total picture(s) on current gallery: " & flFile.ListCount & ".</center>" & vbCrLf
    OutputText = OutputText & "   <hr><br />" & vbCrLf
    
    flFile.SetFocus
    flFile.ListIndex = 0
    flFile.Enabled = False
    
    ' Create table
    OutputText = OutputText & "   <table border=""1"">" & vbCrLf
    
    
    For X = 1 To (flFile.ListCount)
        OutputText = OutputText & "   <tr>" & vbCrLf
        OutputText = OutputText & "      <td><a href=""" & flFile.Filename & """><img src=""thumb/" & flFile.Filename & """></a></td>" & vbCrLf
        OutputText = OutputText & "      <td>&nbsp;&nbsp;&nbsp;<b>File Index:</b> <i>" & (flFile.ListIndex + 1) & "</i> out of <i>" & flFile.ListCount & "</i><br />" & vbCrLf
        OutputText = OutputText & "          &nbsp;&nbsp;&nbsp;<b>Filename:</b> " & flFile.Filename & "<br />" & vbCrLf
        OutputText = OutputText & "          &nbsp;&nbsp;&nbsp;<b>Original file location:</b> " & (drvDrive.Path & "\" & flFile.Filename) & " &nbsp;&nbsp;&nbsp;<br />" & vbCrLf
        OutputText = OutputText & "          &nbsp;&nbsp;&nbsp;<b>Original file size:</b> " & Format(FileLen(drvDrive.Path & "\" & flFile.Filename), "#,###") & " bytes (" & Int(FileLen(drvDrive.Path & "\" & flFile.Filename) / 1024) & " KB) &nbsp;&nbsp;&nbsp;<br />" & vbCrLf
        
        ' Get the current size for selected file & add it
        CurrentFileSize = FileLen(drvDrive.Path & "\" & flFile.Filename)
        TotalGallerySize = TotalGallerySize + CurrentFileSize
        
        OutputText = OutputText & "          &nbsp;&nbsp;&nbsp;<b>Thumbnail file size:</b> " & Format(FileLen(drvDrive.Path & "\thumb\" & flFile.Filename), "#,###") & " bytes (" & Int(FileLen(drvDrive.Path & "\thumb\" & flFile.Filename) / 1024) & " KB)<br /><br />" & vbCrLf
        OutputText = OutputText & "      &nbsp;&nbsp;&nbsp;<a href=""" & flFile.Filename & """>VIEW ORIGINAL PICTURE</a>" & vbCrLf
        OutputText = OutputText & "      </td>" & vbCrLf
        OutputText = OutputText & "   <tr>" & vbCrLf
        flFile.ListIndex = X
        DoEvents
    Next X
    
    ' Closing table
    OutputText = OutputText & "   </table><br />" & vbCrLf
    
    ' Add total gallery size
    'bla = ((TotalGallerySize / 1024) / 1024)
    OutputText = OutputText & "   Estimated current gallery total size: <b>" & Format(TotalGallerySize, "#,###") & "</b> bytes (" & Format(Int(TotalGallerySize / 1024), "#,###") & " KB / " & Format(((TotalGallerySize / 1024) / 1024), "#,###.#") & " MB).<br />" & vbCrLf
    
    ' Adding some information
    OutputText = OutputText & "   <hr>" & vbCrLf
    OutputText = OutputText & "   <p><center>@ Copyleft 2001-" & Format$(Date, "yyyy") & " Ino Bambino. All wrongs reserved.</center></p>" & vbCrLf
    OutputText = OutputText & "   <center><p style=""font-size:x-small;""><b>NOTICE:</b> THIS SOFTWARE IS PROVIDED ""AS IS"" AND ANY EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE REGENTS OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION). HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.</p></center>" & vbCrLf
    
    ' Finalizing the string
    OutputText = OutputText & "</html>" & vbCrLf
    
    flFile.Enabled = True
    
    ' Placing the string on to the textfield
    frmOutput.txtOutput.Text = OutputText
    frmOutput.Show
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    drvDrive.Path = Drive1.Drive
End Sub

Private Sub drvDrive_Change()
    flFile.Path = drvDrive.Path
    lblFileFound.Caption = "File(s) found: " & (flFile.ListCount)
End Sub

Private Sub flFile_DblClick()
Dim SelFileName As String

    If Right$(flFile.Path, 1) = "\" Then
        SelFileName = flFile.Path & flFile.Filename
    Else
        SelFileName = flFile.Path & "\" & flFile.Filename
    End If
    
    ShellOpen SelFileName, Me
    
End Sub

Private Sub Form_Initialize()
' Call the InitCommonControls Function for XP Styles
    InitCommonControls
End Sub

Private Sub Form_Load()
    Me.Caption = App.ProductName & " - v" & App.Major & "." & App.Minor & " (BUILD " & App.Revision & ")"
    lblFileFound.Caption = "File(s) found: " & (flFile.ListCount)
    lblNote.Caption = "NOTE: Please make sure that the ""THUMB"" folder with the thumbnails already exist. Double-click on the filename to open the file using your default picture viewer." & _
                      vbCrLf & vbCrLf & "@ Copyleft 2001-" & Format$(Date, "yyyy") & " Ino Bambino. All wrongs reserved."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblASCII.ForeColor = &HFF0000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lblASCII_Click()
    ShellOpen "http://www.ascii.cl/htmlcodes.htm", Me
End Sub

Private Sub lblASCII_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblASCII.ForeColor = vbRed
End Sub

