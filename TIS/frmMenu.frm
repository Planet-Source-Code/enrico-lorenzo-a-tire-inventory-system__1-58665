VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00808080&
   Caption         =   "Tire Inventory System"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
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
   ScaleHeight     =   6435
   ScaleWidth      =   8505
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
      Height          =   5295
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   7455
      Begin VB.CommandButton cmdMonthly 
         Caption         =   "Monthly Earnings"
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
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "Order Slip"
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
         Left            =   240
         TabIndex        =   4
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton cmdInvoice 
         Caption         =   "Sales Invoice"
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
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton cmdCustomer 
         Caption         =   "Customer Maintenance"
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
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CommandButton cmdTire 
         Caption         =   "Tire Maintenance"
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
         Left            =   240
         TabIndex        =   0
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   705
         Left            =   240
         Picture         =   "frmMenu.frx":0000
         Top             =   360
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()

End Sub

Private Sub cmdCustomer_Click()
frmCustomer.Show
Unload Me
End Sub

Private Sub cmdInvoice_Click()
frmInvoice.Show
Unload Me
End Sub

Private Sub cmdMonthly_Click()
frmMonthly.Show
Unload Me
End Sub

Private Sub cmdOrder_Click()
frmPurchase.Show
Unload Me
End Sub

Private Sub cmdTire_Click()
frmTire.Show
Unload Me
End Sub

