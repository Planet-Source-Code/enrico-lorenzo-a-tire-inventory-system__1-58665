VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTire 
   BackColor       =   &H00808080&
   Caption         =   "Tire Maintenance"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
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
   ScaleHeight     =   9990
   ScaleWidth      =   14685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
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
      TabIndex        =   35
      Top             =   6480
      Width           =   1215
   End
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
      Height          =   8655
      Left            =   0
      TabIndex        =   21
      Top             =   -120
      Width           =   13095
      Begin VB.TextBox txtCode 
         BackColor       =   &H00E0E0E0&
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtPattern2 
         DataField       =   "Pattern"
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSize2 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCategory2 
         DataField       =   "Category"
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   19
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSize 
         Caption         =   "Add Data"
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
         TabIndex        =   4
         Top             =   4320
         Width           =   1215
      End
      Begin MSComCtl2.UpDown updQuantity 
         Height          =   330
         Left            =   3480
         TabIndex        =   9
         Top             =   5880
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdPattern 
         Caption         =   "Add Data"
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
         TabIndex        =   6
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtPattern 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Pattern"
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtSize 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Size"
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   3495
      End
      Begin VB.CommandButton cmdCategory 
         Caption         =   "Add Data"
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
         TabIndex        =   2
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         TabIndex        =   18
         Top             =   2760
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
         TabIndex        =   17
         Top             =   2160
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
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.UpDown updMin 
         Height          =   330
         Left            =   3480
         TabIndex        =   11
         Top             =   6480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown updMax 
         Height          =   330
         Left            =   3480
         TabIndex        =   13
         Top             =   7080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtMax 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Max"
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   7080
         Width           =   3495
      End
      Begin VB.TextBox txtCategory 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Category"
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2280
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1320
         Top             =   3600
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
         RecordSource    =   "Categories"
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
         Height          =   330
         Left            =   4800
         Top             =   3480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "Sizes"
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   8640
         Top             =   3360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "Patterns"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   495
         Left            =   5040
         Top             =   6240
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmTire.frx":0000
         Height          =   1455
         Left            =   3960
         TabIndex        =   32
         Top             =   2760
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmTire.frx":0015
         Height          =   1455
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   3495
         _ExtentX        =   6165
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
         Bindings        =   "frmTire.frx":002A
         Height          =   1455
         Left            =   7680
         TabIndex        =   33
         Top             =   2760
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.TextBox txtQuantity 
         BackColor       =   &H00E0E0E0&
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5880
         Width           =   3495
      End
      Begin VB.TextBox txtMin 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Min"
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   6480
         Width           =   3495
      End
      Begin VB.TextBox txtPrice 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5280
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmTire.frx":003F
         Height          =   1455
         Left            =   3960
         TabIndex        =   34
         Top             =   5040
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
      Begin VB.Label Label1 
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
         TabIndex        =   36
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   26
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Selling Price"
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
         TabIndex        =   25
         Top             =   5040
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tread Pattern"
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
         TabIndex        =   24
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tire Size"
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
         TabIndex        =   23
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Category"
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
         TabIndex        =   22
         Top             =   2040
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   240
         Picture         =   "frmTire.frx":0054
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Minimum Tire Quantity"
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
         TabIndex        =   27
         Top             =   6240
         Width           =   2205
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Maximum Tire Quantity"
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
         TabIndex        =   28
         Top             =   6840
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmTire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mb

Private Sub cmdAdd_Click()
Adodc4.Recordset.AddNew

addupdate

Randomize
txtCode = "TIR" & Round(Rnd() * 999999) & txtCode + Chr(Round(Rnd() * 25) + 65)
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmTire.Show
End Sub

Private Sub cmdDelete_Click()
Adodc4.Recordset.delete

delete
End Sub

Private Sub cmdReport_Click()
Set DataReport1.DataSource = Adodc4
DataReport1.Show
frmTire.Enabled = False
End Sub

Private Sub cmdSave_Click()

If Len(txtCategory) > 0 And Len(txtSize) > 0 And Len(txtPattern) > 0 And IsNumeric(txtPrice) = True And Len(txtQuantity) > 0 And Len(txtMin) > 0 And Len(txtMax) > 0 And Val(txtMin) <= Val(txtMax) Then
    Adodc4.Recordset.UpdateBatch
    savecancel
    delete
    
    If Val(txtQuantity) > Val(txtMax) Then
        mb = MsgBox("Reminder: This tire is above its maximum quantity.", vbExclamation, "Yokohama")
    ElseIf Val(txtQuantity) < Val(txtMin) Then
        mb = MsgBox("Reminder: This tire is below its minimum quantity.", vbExclamation, "Yokohama")
    End If
    
Else
    mb = MsgBox("Invalid data.", vbCritical, "Yokohama")
End If



End Sub

Private Sub cmdUpdate_Click()
addupdate

End Sub

Private Sub DataGrid1_Click()
txtCategory = txtCategory2
txtCategory.SetFocus
End Sub

Private Sub DataGrid2_Click()
txtSize = txtSize2
txtSize.SetFocus
End Sub

Private Sub DataGrid3_Click()
txtPattern = txtPattern2
txtPattern.SetFocus
End Sub

Private Sub Form_Activate()

delete

End Sub

Private Function addupdate()
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
cmdDelete.Enabled = False
cmdReport.Enabled = False

cmdCategory.Enabled = False
cmdSize.Enabled = False
cmdPattern.Enabled = False

DataGrid1.Enabled = True
DataGrid2.Enabled = True
DataGrid3.Enabled = True
DataGrid4.Enabled = False

updQuantity.Enabled = True
updMin.Enabled = True
updMax.Enabled = True

txtPrice.Locked = False

End Function

Private Function savecancel()
DataGrid4.Refresh

cmdAdd.Enabled = True
cmdUpdate.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmdReport.Enabled = True

cmdCategory.Enabled = True
cmdSize.Enabled = True
cmdPattern.Enabled = True

updQuantity.Enabled = False
updMin.Enabled = False
updMax.Enabled = False

DataGrid1.Enabled = False
DataGrid2.Enabled = False
DataGrid3.Enabled = False
DataGrid4.Enabled = True

txtPrice.Locked = True

End Function

Private Function delete()
DataGrid4.Refresh

If Adodc4.Recordset.RecordCount = 0 Then
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    cmdReport.Enabled = False
Else

    cmdDelete.Enabled = True
    cmdUpdate.Enabled = True
    cmdReport.Enabled = True
End If

End Function

Private Sub cmdBack_Click()
frmMenu.Show
Unload Me
End Sub

Private Sub cmdCategory_Click()
frmCategory.Show
Unload Me
End Sub

Private Sub cmdPattern_Click()
frmPattern.Show
Unload Me
End Sub

Private Sub cmdSize_Click()
frmSize.Show
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc4.Recordset.CancelBatch
End Sub

Private Sub txtCategory2_Change()
txtCategory = txtCategory2
End Sub

Private Sub txtPattern2_Change()
txtPattern = txtPattern2
End Sub

Private Sub txtSize2_Change()
txtSize = txtSize2
End Sub

Private Sub updMax_DownClick()
txtMax = Val(txtMax) - 1

If txtMax < 0 Then
    txtMax = 0
End If

txtMax.SetFocus
End Sub

Private Sub updMax_UpClick()
txtMax = Val(txtMax) + 1
txtMax.SetFocus
End Sub

Private Sub updMin_DownClick()
txtMin = Val(txtMin) - 1

If txtMin < 0 Then
    txtMin = 0
End If

txtMin.SetFocus
End Sub

Private Sub updMin_UpClick()
txtMin = Val(txtMin) + 1
txtMin.SetFocus
End Sub

Private Sub updQuantity_DownClick()
txtQuantity = Val(txtQuantity) - 1

If txtQuantity < 0 Then
    txtQuantity = 0
End If

txtQuantity.SetFocus
End Sub

Private Sub updQuantity_UpClick()
txtQuantity = Val(txtQuantity) + 1
txtQuantity.SetFocus
End Sub
