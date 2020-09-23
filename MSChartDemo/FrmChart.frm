VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmChart 
   BackColor       =   &H00C7BEAD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart Demo"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   9240
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7BEAD&
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   8280
      TabIndex        =   3
      Top             =   360
      Width           =   3135
      Begin VB.CommandButton CmdSaveas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save As..."
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print Chart"
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3960
         Width           =   2415
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6735
      Left            =   0
      OleObjectBlob   =   "FrmChart.frx":0000
      TabIndex        =   4
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "FrmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Con As New ADODB.Connection ' Connection
Dim Rss As New ADODB.Recordset ' Recordset

Dim ArrayChart() 'Array
Dim X As Integer
Private Sub CmdPrint_Click()
    MSChart1.EditCopy 'This Makes MSChart Control to be Copied
    DoEvents   ' may be needed for large datasets
    Printer.Print " "
    Picture1.PaintPicture Clipboard.GetData(), 0, 0
    Printer.EndDoc
End Sub
Private Sub CmdSaveas_Click()
    On Error GoTo Hell
    Dim strsavefile As String
    With CommonDialog1
        .Filter = "Pictures (*.bmp)|*.bmp" ' You can Also Save the Pic in JPG/GIF/TIFF
        .DefaultExt = "bmp"
        .CancelError = False
        .ShowSave
        strsavefile = .FileName
        If strsavefile = "" Then Exit Sub
    End With
    MSChart1.EditCopy
    SavePicture Clipboard.GetData, strsavefile ' File Saved
    Exit Sub
Hell:
    MsgBox Err.Description
End Sub
Private Sub Command1_Click()
    End ' Quitting
End Sub
Private Sub Form_Load()
Set Con = Nothing
Con.Open "Provider = Microsoft.jet.OLEDB.4.0; Data Source =" & App.Path & "\DBMain.Mdb" ' Making Connection to with the Database
Set Rss = Nothing
    Rss.Open "Select * from Main", Con, adOpenKeyset ' Making Recordset
    If Rss.RecordCount = 0 Then ' If no Record in Database, then Show an Error Msg and Exit the Sub
        MsgBox "No Data to Show on Chart!!!", vbCritical, "MSChart Demo": Exit Sub
    Else
        ReDim ArrayChart(1 To Rss.RecordCount, 1 To 2) ' Array
        'Puuting Records from Database to Array
        For X = 1 To Rss.RecordCount
        ArrayChart(X, 1) = Rss!Name
        ArrayChart(X, 2) = CInt(Rss!Total)
        Rss.MoveNext
        Next X
'# Assigns our array to the MSChart control #
    MSChart1.ChartData = ArrayChart
    MSChart1.Refresh
    End If
End Sub
