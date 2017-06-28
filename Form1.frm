VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   6105
   ClientTop       =   2325
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   Begin MSComDlg.CommonDialog cd 
      Left            =   -7500
      Top             =   -7500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnCopy 
      Caption         =   "Copy"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton btnClearCode 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtCode 
      Height          =   3255
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton btnToPtOn 
      Caption         =   "To Pt-On()"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox imgSmall 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   1320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   1425
      TabIndex        =   0
      Top             =   120
      Width           =   1425
   End
   Begin VB.PictureBox imgSmallHid 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   1455
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgLarge 
      Height          =   1800
      Left            =   120
      Picture         =   "Form1.frx":22F0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Close Up View of Image:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Image To Scan:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const IMG_HEIGHT As Integer = 95
Private Const IMG_WIDTH As Integer = 63

Private Sub Command1_Click()
    Dim clr As Long
    Dim s As String
    
    For y = 0 To IMG_WIDTH
        For x = 0 To IMG_HEIGHT
            clr = pic.Point(x, y)
            If clr = 0 Then
                s = s & "Pt-On(" & x & "," & (-1 * y) & ")" & vbCrLf
            End If
            Debug.Print "(" & x & "," & y & ") " & clr
        Next
    Next
    
    txtCode.Text = s
End Sub

Private Sub btnClearCode_Click()
    txtCode.Text = ""
End Sub

Private Sub btnCopy_Click()
    Clipboard.Clear
    Clipboard.SetText s
End Sub

Private Sub btnLoad_Click()
    Dim bimgClear As Boolean: bimgClear = True
        
    cd.DialogTitle = "Pick Image..."
    cd.Filter = "Images|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur"
    cd.ShowOpen
    If cd.CancelError = False Then
        imgSmall.Picture = LoadPicture(cd.FileName)
        imgSmallHid.Picture = imgSmall.Picture
        imgLarge.Picture = imgSmall.Picture
        
        bimgClear = False
        
        If imgSmallHid.Width > IMG_WIDTH Or imgSmallHid.Height > IMG_HEIGHT Then
            MsgBox "Error! Invalid image size. Please make sure it is no larger than " & IMG_WIDTH & "x" & IMG_HEIGHT & " pixels"
            bimgClear = True
        End If
    End If
    
    If bimgClear = True Then
        imgSmall.Picture = Nothing
        imgSmallHid.Picture = Nothing
        imgLarge.Picture = Nothing
    End If
End Sub

Private Sub cmdSave_Click()
    
End Sub
