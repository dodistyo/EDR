VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Register Or Edit Member"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   6720
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4440
      TabIndex        =   26
      Top             =   6720
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   25
      Text            =   "Choose..."
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   24
      Text            =   "Choose..."
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   1560
      TabIndex        =   23
      Top             =   6000
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   1560
      TabIndex        =   20
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1057
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Height/Weight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Birthday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Education"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "EID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

For Each Control In Me.Controls
If TypeOf Control Is TextBox Then
If Control.Text = vbNullString Then
MsgBox "Please Fill All Information", vbOKOnly + vbExclamation, "Blank Text Found"
Control.SetFocus
Exit Sub
End If
End If
Next Control

Form1.Ado1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dodo.mdb;Persist Security Info=False"
Form1.Ado1.Refresh
Form1.Ado1.Recordset.AddNew
Form1.Ado1.Recordset("EmployeeID") = Text1.Text
Form1.Ado1.Recordset("Name") = Text2.Text
Form1.Ado1.Recordset("Gender") = Combo1.Text
Form1.Ado1.Recordset("Height/Weight") = Text3.Text
Form1.Ado1.Recordset("Contact") = Text4.Text
Form1.Ado1.Recordset("Position") = Text5.Text
Form1.Ado1.Recordset("Salary") = Text6.Text
Form1.Ado1.Recordset("Address") = Text7.Text
Form1.Ado1.Recordset("Religion") = Text8.Text
Form1.Ado1.Recordset("Status") = Combo2.Text
Form1.Ado1.Recordset("Birthday") = Text9.Text
Form1.Ado1.Recordset("Education") = Text10.Text

FileCopy Text11.Text, App.Path & "\Photos\" & Text1.Text & ".jpg"

For Each Control In Form3.Controls

If TypeName(Control) = "TextBox" Then
Control.Text = ""

End If
Next
Combo1.Text = "Choose..."
Combo2.Text = "Choose..."
Form1.DataGrid1.Refresh
Form1.Ado1.Recordset.Update
Form1.Ado1.Recordset.Sort = "EmployeeID"
MsgBox "Data Have Been Added", vbInformation + vbOKOnly, "Succesful"
End Sub

Private Sub Command2_Click()
Form1.Ado1.Recordset("EmployeeID") = Text1.Text
Form1.Ado1.Recordset("Name") = Text2.Text
Form1.Ado1.Recordset("Gender") = Combo1.Text
Form1.Ado1.Recordset("Height/Weight") = Text3.Text
Form1.Ado1.Recordset("Contact") = Text4.Text
Form1.Ado1.Recordset("Position") = Text5.Text
Form1.Ado1.Recordset("Salary") = Text6.Text
Form1.Ado1.Recordset("Address") = Text7.Text
Form1.Ado1.Recordset("Religion") = Text8.Text
Form1.Ado1.Recordset("Status") = Combo2.Text
Form1.Ado1.Recordset("Birthday") = Text9.Text
Form1.Ado1.Recordset("Education") = Text10.Text

If Text11.Text = App.Path & "\Photos\" & Text1.Text & ".jpg" Then
Text11.Text = ""
End If
If Text11.Text = "" Then
Form1.Ado1.Recordset.Update
Form1.Ado1.Recordset.Sort = "EmployeeID"
Command2.Enabled = False
Command1.Enabled = True
Unload Form3
MsgBox "Data Have Been Edited", vbOKOnly, "Saved!"
Exit Sub
End If

FileCopy Text11.Text, App.Path & "\Photos\" & Text1.Text & ".jpg"
Unload Form3
MsgBox "Data Have Been Edited", vbOKOnly, "Saved!"
End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "Picture (*.jpg)|*.jpg|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "jpg"
CommonDialog1.DialogTitle = "Select Photo"
CommonDialog1.ShowOpen
Text11.Text = CommonDialog1.FileName
Text11.Enabled = False
End Sub

Private Sub Form_Load()
Command2.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyTab) Then
Else
Beep
KeyAscii = 0
End If
End Sub
