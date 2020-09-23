VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImage 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Image Informer"
   ClientHeight    =   3720
   ClientLeft      =   2370
   ClientTop       =   1200
   ClientWidth     =   7320
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   1080
      Width           =   3375
      Begin VB.VScrollBar VScroll 
         Height          =   2190
         LargeChange     =   15
         Left            =   3120
         SmallChange     =   5
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar HScroll 
         Height          =   255
         LargeChange     =   15
         Left            =   0
         SmallChange     =   5
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.PictureBox PicImage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdAuthor 
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtSize 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Read Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "( Supoorts only Jpg, Gif, Bmp, Png )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "File Size (Bytes)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Image Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cdlgImage 
         Left            =   3120
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Browse To Open Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        '########################################'
        '   Programmed By Inderpal Singh         '
        '   Email: inderpal0@hotmail.com         '
        '   Date: April 9, 2002                  '
        '   Homepage: http://connect.to/lanserver'
        '########################################'

Option Explicit
Private ImageTypes(4) As String

Private Sub cmdAuthor_Click()
    frmAbout.Show
End Sub

Private Sub cmdBrowse_Click()
    Dim Info As String
On Error GoTo Err
    txtType = ""
    txtWidth = ""
    txtHeight = ""
    txtSize = ""
    cmdInfo.Enabled = True
    cdlgImage.ShowOpen
    txtPath.Text = cdlgImage.FileName
    Dim Lower As String
    Lower = LCase(txtPath)
    PicImage.Cls
    Info = Right$(Lower, Len(Lower) - InStr(Lower, "."))
    If Info = "jpg" Or Info = "gif" Or Info = "bmp" Or Info = "png" Then
        PicImage.Picture = LoadPicture(cdlgImage.FileName)
        If PicImage.Width > picScroll.Width Then
            HScroll.Max = PicImage.Width
            HScroll.Visible = True
        End If
        If PicImage.Height > picScroll.Height Then
            VScroll.Max = PicImage.Height
            VScroll.Visible = True
        End If
        If txtPath = "" Then
            PicImage.Visible = False
            Exit Sub
        End If
        PicImage.Visible = True
    Else
        MsgBox "Pls Enter A Valid File", vbInformation, "File Error"
        txtPath = ""
        cmdBrowse.SetFocus
        Exit Sub
    End If
Err:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdInfo_Click()
    If txtPath = "" Then
        MsgBox "Pls Enter The Image Path", vbInformation, "Path Error"
        cmdBrowse.SetFocus
        Exit Sub
    End If
    ImageTypes(0) = "Unknown"
    ImageTypes(1) = "GIF"
    ImageTypes(2) = "JPEG"
    ImageTypes(3) = "PNG"
    ImageTypes(4) = "BMP"
    ReadImageInfo (txtPath.Text)
    txtHeight.Text = ImageHeight
    txtWidth.Text = ImageWidth
    txtType.Text = ImageTypes(ImageType)
    txtSize.Text = FileSize
End Sub

Private Sub Form_Load()
    cmdInfo.Enabled = False
End Sub

Private Sub HScroll_Change()
    HScroll_Scroll
End Sub
Private Sub HScroll_Scroll()
   PicImage.Left = -HScroll.Value
End Sub
Private Sub VScroll_Change()
    VScroll_Scroll
End Sub

Private Sub VScroll_Scroll()
    PicImage.Top = -VScroll.Value
End Sub

