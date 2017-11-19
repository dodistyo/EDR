VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees Database"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Search"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Ado1 
      Height          =   375
      Left            =   4200
      Top             =   5280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   8454143
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "EmployeeID"
         Caption         =   "EmployeeID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """Rp""#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "          Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Gender"
         Caption         =   "  Gender"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Contact"
         Caption         =   "    Contact"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Position"
         Caption         =   "        Position"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Salary"
         Caption         =   "       Salary"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form3.Text1.Text = Ado1.Recordset("EmployeeID")
Form3.Text2.Text = Ado1.Recordset("Name")
Form3.Combo1.Text = Ado1.Recordset("Gender")
Form3.Text3.Text = Ado1.Recordset("Height/Weight")
Form3.Text4.Text = Ado1.Recordset("Contact")
Form3.Text5.Text = Ado1.Recordset("Position")
Form3.Text6.Text = Ado1.Recordset("Salary")
Form3.Text7.Text = Ado1.Recordset("Address")
Form3.Text8.Text = Ado1.Recordset("Religion")
Form3.Combo2.Text = Ado1.Recordset("Status")
Form3.Text9.Text = Ado1.Recordset("Birthday")
Form3.Text10.Text = Ado1.Recordset("Education")
Form3.Command2.Enabled = True
Form3.Command1.Enabled = False

If Dir(App.Path & "\Photos\" & Form3.Text1.Text & ".jpg") <> "" Then
Form3.Text11.Text = App.Path & "\Photos\" & Form3.Text1.Text & ".jpg"
End If

Form3.Show
End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure You want to delete this?", vbYesNo + vbDefaultButton2 + vbExclamation, "Warning!") = vbYes Then
Ado1.Recordset.Delete
End If
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
Ado1.Recordset.Update
DataGrid1.ReBind
DataGrid1.Refresh
End Sub


Private Sub Form_Load()
Ado1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dodo.mdb;Persist Security Info=False"
Ado1.RecordSource = "Select * From Employee"
Ado1.Refresh
Ado1.Recordset.Sort = "EmployeeID"
Set DataGrid1.DataSource = Ado1
    With DataGrid1
    .Columns(0).Width = 1100
    .Columns(2).Width = 950
    .Columns(3).Width = 1300
    .Columns(5).Width = 1600
    End With
End Sub

