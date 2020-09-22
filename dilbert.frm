VERSION 5.00
Begin VB.Form frmDilbert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Dilbert"
   ClientHeight    =   1260
   ClientLeft      =   3672
   ClientTop       =   2556
   ClientWidth     =   4080
   Icon            =   "dilbert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4080
   Begin VB.Label lblDilbert 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading today's Dilbert..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3732
   End
End
Attribute VB_Name = "frmDilbert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function URLDownloadToCacheFile Lib "urlmon" Alias _
    "URLDownloadToCacheFileA" (ByVal lpUnkcaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwBufLength As Long, _
    ByVal dwReserved As Long, _
    ByVal IBindStatusCallback As Long) As Long

Public Sub Create()
    Me.Show
End Sub

Private Sub Form_Load()
    Dim lStartLocation As Long
    Dim lEndLocation As Long
    Dim lLen As Long
    Dim sPath As String
    Dim sHTM As String
    Dim sLocalFilename As String
    Dim fso As FileSystemObject
    Dim ts As TextStream
        
    On Error GoTo ErrorHandler
        
    Set Me.Font = lblDilbert.Font
    lblDilbert.Width = Me.TextWidth(lblDilbert.Caption) + 500
    lblDilbert.Height = Me.TextHeight(lblDilbert.Caption) + 500
    Me.Height = lblDilbert.Height
    Me.Width = lblDilbert.Width
    Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Me.Show
    DoEvents
    
    sLocalFilename = DownloadFile("http://www.unitedmedia.com/comics/dilbert/")

    Set fso = New FileSystemObject
    Set ts = fso.OpenTextFile(sLocalFilename, ForReading, False)
    sHTM = ts.ReadAll
    
    lStartLocation = InStr(sHTM, "/comics/dilbert/archive/images/dilbert")
    lEndLocation = InStr(lStartLocation, sHTM, ".gif")
    lLen = lEndLocation - lStartLocation + 4
    sPath = Mid(sHTM, lStartLocation, lLen)

    sLocalFilename = DownloadFile("http://www.unitedmedia.com" & sPath)

    Set Me.Picture = LoadPicture(sLocalFilename)
    
    lblDilbert.Visible = False
    
    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
    Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
    
    Me.Enabled = True
    Screen.MousePointer = vbNormal
    
    Exit Sub
ErrorHandler:
    MsgBox "Unable to download image from Internet.", vbOKOnly, "Error..."
    Me.Enabled = True
    Screen.MousePointer = vbNormal
    Unload Me
End Sub

Private Function DownloadFile(URL As String) As String
    Dim lngRetVal As Long
    Dim sLocalFilename As String
    
    sLocalFilename = Space(300)
    lngRetVal = URLDownloadToCacheFile(0, URL, sLocalFilename, Len(sLocalFilename), 0, 0)
    
    If lngRetVal = 0 Then
        DownloadFile = Trim(sLocalFilename)
    End If
End Function
