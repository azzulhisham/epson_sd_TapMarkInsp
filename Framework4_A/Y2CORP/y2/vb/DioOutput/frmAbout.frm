VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "バージョン情報"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   3240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   3240
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "DIO 出力確認サンプル"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1740
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 0.00"
      Height          =   180
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   945
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "(C) 2008 株式会社ワイツー"
      Height          =   180
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   2085
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
End Sub
