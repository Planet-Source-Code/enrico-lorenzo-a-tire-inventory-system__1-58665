VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00808080&
   Caption         =   "Administrator Password"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
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
   ScaleHeight     =   3405
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   6255
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "Ï"
         TabIndex        =   0
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   240
         Picture         =   "frmAdmin.frx":0000
         Top             =   360
         Width           =   3120
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mb

Private Sub cmdEnter_Click()
If txtPassword = "DBMS" Then
    frmUser.Show
    Unload Me
Else
    mb = MsgBox("Invalid password", vbCritical, "Yokohama")
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
End If

End Sub

