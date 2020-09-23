VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInvoice 
   BackColor       =   &H00808080&
   Caption         =   "Sales Invoice"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14640
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
   ScaleHeight     =   10080
   ScaleWidth      =   14640
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
      Height          =   9855
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   13935
      Begin VB.CommandButton cmdReceipt 
         Caption         =   "Receipt"
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
         Left            =   6240
         TabIndex        =   64
         Top             =   8400
         Width           =   1215
      End
      Begin VB.TextBox txtMonth 
         DataField       =   "Month"
         DataSource      =   "Adodc5"
         Height          =   495
         Left            =   9840
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtYear 
         DataField       =   "Year"
         DataSource      =   "Adodc5"
         Height          =   495
         Left            =   9840
         TabIndex        =   60
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAmount2 
         DataField       =   "Amount"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   59
         Top             =   5880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDiscount2 
         DataField       =   "Discount"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   58
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPaid2 
         DataField       =   "Paid"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   11280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   57
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTotal2 
         DataField       =   "Total"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   56
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtChange2 
         DataField       =   "Change"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   55
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtQuantity2 
         DataField       =   "Quantity"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   54
         Top             =   9000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDescription3 
         DataField       =   "Tire Description"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   53
         Top             =   8520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPrice3 
         DataField       =   "Price"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   10680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         Top             =   9000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTireCode3 
         DataField       =   "Tire Code"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   10680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         Top             =   8520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAddress3 
         DataField       =   "Customer Address"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Top             =   8040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtName3 
         DataField       =   "Customer Name"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   12000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Top             =   7560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtDate2 
         DataField       =   "Date"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   10680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Top             =   8040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCode2 
         DataField       =   "Sales Invoice Code"
         DataSource      =   "Adodc4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   10680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Top             =   7560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "Compute"
         Enabled         =   0   'False
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
         Left            =   9960
         TabIndex        =   45
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Tire Description"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Customer Address"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtDescription2 
         DataField       =   "Size"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   6120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Date"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Sales Invoice Code"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<Menu>"
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
         Left            =   11400
         TabIndex        =   34
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
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
         Left            =   11400
         TabIndex        =   33
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   11400
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add New"
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
         Left            =   11400
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtMax 
         DataField       =   "Max"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   6120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtMin 
         DataField       =   "Min"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   6120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Amount"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtPrice2 
         DataField       =   "Price"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   6120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
         Enabled         =   0   'False
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
         Left            =   9960
         TabIndex        =   25
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtAddress2 
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtName2 
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtChange 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Change"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox txtPaid 
         DataField       =   "Paid"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   5520
         Width           =   3495
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Total"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Top             =   4080
         Width           =   3495
      End
      Begin VB.TextBox txtPrice 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Price"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtTireCode2 
         DataField       =   "Tire Code"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   6120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTireCode 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Tire Code"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Customer Name"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   2280
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   2280
         Top             =   3720
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Customers"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   3720
         Top             =   5520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Tires"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComCtl2.UpDown updQuantity 
         Height          =   330
         Left            =   10920
         TabIndex        =   12
         Top             =   2280
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Max             =   9999
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown updDiscount 
         Height          =   330
         Left            =   10920
         TabIndex        =   17
         Top             =   3480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Max             =   100
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtDiscount 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Discount"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3480
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   495
         Left            =   4320
         Top             =   6960
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Invoice"
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmInvoice.frx":0000
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   2760
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmInvoice.frx":0015
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   5160
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmInvoice.frx":002A
         Height          =   1455
         Left            =   240
         TabIndex        =   23
         Top             =   6840
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtAvailable 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Quantity"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox txtQuantity 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Quantity"
         DataSource      =   "Adodc3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2280
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   495
         Left            =   8040
         Top             =   7680
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "IDummy"
         Caption         =   "Adodc4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmInvoice.frx":003F
         Height          =   1455
         Left            =   7800
         TabIndex        =   46
         Top             =   7200
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmInvoice.frx":0054
         Height          =   1455
         Left            =   7200
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   6000
         Top             =   720
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Monthly"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtEarnings 
         DataField       =   "Earnings"
         DataSource      =   "Adodc5"
         Height          =   495
         Left            =   5880
         TabIndex        =   63
         Top             =   7440
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tire Description"
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
         Left            =   3960
         TabIndex        =   44
         Top             =   4440
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Address"
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
         Left            =   3960
         TabIndex        =   42
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   3960
         TabIndex        =   39
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         Left            =   7680
         TabIndex        =   37
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         TabIndex        =   36
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tire Code"
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
         TabIndex        =   26
         Top             =   4440
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Change"
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
         Left            =   7680
         TabIndex        =   19
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Amount Paid"
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
         Left            =   7680
         TabIndex        =   16
         Top             =   5280
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
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
         Left            =   7680
         TabIndex        =   14
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   7680
         TabIndex        =   11
         Top             =   3840
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Price"
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
         Left            =   7680
         TabIndex        =   9
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         Left            =   7680
         TabIndex        =   7
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
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
         TabIndex        =   3
         Top             =   2040
         Width           =   1545
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   240
         Picture         =   "frmInvoice.frx":0069
         Top             =   360
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mb, monthx

Private Sub cmdAdd_Click()
If Adodc4.Recordset.RecordCount > 0 Then
    Adodc4.Recordset.delete
End If

Adodc3.Recordset.AddNew
addupdate

Randomize
txtCode = "INV" & Round(Rnd() * 999999) & txtCode + Chr(Round(Rnd() * 25) + 65)
txtDate = Now
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmInvoice.Show
End Sub

Private Sub cmdCompute_Click()
If Val(txtPaid) >= Val(txtTotal) And IsNumeric(txtPaid) = True Then
    txtChange = Val(txtPaid) - Val(txtTotal)
    cmdCompute.Enabled = False
    cmdSave.Enabled = True
Else
    mb = MsgBox("Invalid data.", vbCritical, "Yokohama")
End If

End Sub

Private Sub cmdDelete_Click()
Adodc3.Recordset.delete

delete
End Sub

Private Sub cmdReceipt_Click()
If Adodc4.Recordset.RecordCount > 0 Then
    Adodc4.Recordset.delete
End If

Adodc4.Recordset.AddNew
txtCode2 = txtCode
txtDate2 = txtDate
txtName3 = txtName
txtAddress3 = txtAddress
txtTireCode3 = txtTireCode
txtDescription3 = txtDescription
txtPrice3 = txtPrice
txtQuantity2 = txtQuantity
txtAmount2 = txtAmount
txtDiscount2 = txtDiscount
txtTotal2 = txtTotal
txtPaid2 = txtPaid
txtChange2 = txtChange
Adodc4.Recordset.UpdateBatch

Set DataReport2.DataSource = Adodc4
DataReport2.Show
frmInvoice.Enabled = False
End Sub

Private Sub cmdSave_Click()
If Len(txtName) > 0 Then
    Adodc2.Recordset.UpdateBatch
    Adodc3.Recordset.UpdateBatch
    savecancel
    
    If Month(Now) = 1 Then
        monthx = "January"
    ElseIf Month(Now) = 2 Then
        monthx = "February"
    ElseIf Month(Now) = 3 Then
        monthx = "March"
    ElseIf Month(Now) = 4 Then
        monthx = "April"
    ElseIf Month(Now) = 5 Then
        monthx = "May"
    ElseIf Month(Now) = 6 Then
        monthx = "June"
    ElseIf Month(Now) = 7 Then
        monthx = "July"
    ElseIf Month(Now) = 8 Then
        monthx = "August"
    ElseIf Month(Now) = 9 Then
        monthx = "September"
    ElseIf Month(Now) = 10 Then
        monthx = "October"
    ElseIf Month(Now) = 11 Then
        monthx = "November"
    ElseIf Month(Now) = 12 Then
        monthx = "December"
    End If

    Adodc5.Recordset.Find "Month=" & "'" & monthx & "'"

    If Len(txtMonth) > 0 And Year(Now) = txtYear Then
        txtEarnings = Val(txtEarnings) + Val(txtTotal)
        txtEarnings.SetFocus
        Adodc5.Recordset.UpdateBatch
    Else
        Adodc5.Recordset.AddNew
        txtMonth = monthx
        txtYear = Year(Now)
        txtEarnings = txtTotal
        Adodc5.Recordset.UpdateBatch
    End If
    
    cmdReceipt.Enabled = True
Else
    mb = MsgBox("Invalid data.", vbCritical, "Yokohama")
End If

End Sub

Private Sub cmdTotal_Click()
If Len(txtPrice) > 0 And Len(txtQuantity) > 0 And Len(txtDiscount) > 0 And Val(txtAvailable) >= Val(txtQuantity) Then
    txtAmount = Val(txtPrice) * Val(txtQuantity)
    txtTotal = Val(txtAmount) - Val(txtAmount) * Val(txtDiscount) / 100
    txtAvailable = Val(txtAvailable) - Val(txtQuantity)
    txtAvailable.SetFocus
    txtPaid.Locked = False
    
    cmdCompute.Enabled = True
    cmdTotal.Enabled = False
    
    updQuantity.Enabled = False
    updDiscount.Enabled = False
    
    DataGrid2.Enabled = False
    
    If Val(txtAvailable) > Val(txtMax) Then
        mb = MsgBox("Reminder: This tire is above its maximum quantity.", vbExclamation, "Yokohama")
    ElseIf Val(txtAvailable) < Val(txtMin) Then
        mb = MsgBox("Reminder: This tire is below its minimum quantity.", vbExclamation, "Yokohama")
    End If
Else
    mb = MsgBox("Invalid data.", vbCritical, "Yokohama")
End If


End Sub

Private Sub DataGrid1_Click()
txtName = txtName2
txtAddress = txtAddress2
End Sub

Private Sub DataGrid2_Click()
txtTireCode = txtTireCode2
txtDescription = txtDescription2
txtPrice = txtPrice2
End Sub

Private Function addupdate()
cmdAdd.Enabled = False
cmdCancel.Enabled = True

cmdTotal.Enabled = True
cmdReceipt.Enabled = False

DataGrid1.Enabled = True
DataGrid2.Enabled = True
DataGrid3.Enabled = False

updQuantity.Enabled = True
updDiscount.Enabled = True

txtPaid.Locked = False
End Function

Private Function savecancel()
DataGrid3.Refresh

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False

cmdReceipt.Enabled = True

DataGrid1.Enabled = False
DataGrid3.Enabled = True

updQuantity.Enabled = False
End Function

Private Sub cmdBack_Click()
frmMenu.Show
Unload Me
End Sub

Private Sub Form_Activate()
If Adodc4.Recordset.RecordCount = 0 Then
    cmdReceipt.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc2.Recordset.CancelBatch
Adodc3.Recordset.CancelBatch
End Sub

Private Sub txtAddress2_Change()
txtAddress = txtAddress2
End Sub

Private Sub txtDescription2_Change()
txtDescription = txtDescription2
End Sub

Private Sub txtName2_Change()
txtName = txtName2
End Sub

Private Sub txtPrice2_Change()
txtPrice = txtPrice2
End Sub

Private Sub txtTireCode2_Change()
txtTireCode = txtTireCode2
End Sub

Private Sub updDiscount_Change()
txtDiscount = updDiscount
End Sub

Private Sub updQuantity_DownClick()
txtQuantity = Val(txtQuantity) - 1
txtQuantity.SetFocus

If txtQuantity < 0 Then
    txtQuantity = 0
End If

End Sub

Private Sub updQuantity_UpClick()
txtQuantity = Val(txtQuantity) + 1
txtQuantity.SetFocus
End Sub

