VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00C7BEAD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Records"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
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
   ScaleHeight     =   3675
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdShowChart 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Chart"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox TxtBasicPay 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   3
      Top             =   765
      Width           =   2415
   End
   Begin VB.TextBox TxtHRA 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1275
      Width           =   2415
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      Height          =   495
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
      Height          =   495
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   1800
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   450
      TabIndex        =   8
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Pay :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   765
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HRA :"
      Height          =   195
      Left            =   585
      TabIndex        =   6
      Top             =   1275
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      Height          =   195
      Left            =   525
      TabIndex        =   5
      Top             =   1800
      Width           =   555
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

