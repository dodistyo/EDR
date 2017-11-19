VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Profile"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   26
      Text            =   "Search by..."
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   255
      Left            =   8040
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find and Show"
      Default         =   -1  'True
      Height          =   315
      Left            =   6360
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox Ado1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8160
      ScaleHeight     =   270
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4995
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4995
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Address      :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   1800
      TabIndex        =   22
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1800
      TabIndex        =   18
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Height/W    :           :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Status          :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Education  :         :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday      :         :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Religion      :"
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
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EID                :"
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
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name           :"
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
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender        :"
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
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact        :"
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
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Position       :"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Photo Profil"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
Text1.Text = ""
End If
If Combo1.ListIndex = 1 Then
Text1.Text = ""
End If
End Sub

Private Sub Command1_Click()
Form1.Ado1.Refresh
Form1.DataGrid1.Refresh
Form1.Ado1.Recordset.Update
Form1.Ado1.Recordset.Sort = "EmployeeID"

Dim cari As String

If Combo1.Text = "EID" Then
cari = "EmployeeID = '" & Text1 & "'"
With Form1.Ado1.Recordset
    .Find cari
    If Not .EOF Then
    Label12.Caption = .Fields(0)
    Label13.Caption = .Fields(1)
    Label14.Caption = .Fields(2)
    Label15.Caption = .Fields(3)
    Label16.Caption = .Fields(4)
    Label17.Caption = .Fields(5)
    Label18.Caption = .Fields(7)
    Label19.Caption = .Fields(8)
    Label20.Caption = .Fields(9)
    Label21.Caption = .Fields(10)
    Label22.Caption = .Fields(11)
    
    Else
    MsgBox "Cannot Find " & " " & Text1 & " " & "Data Is Not Exist", vbCritical, "Not Found"
    Text1.Text = ""
    Text1.SetFocus
    End If
    End With
    End If

If Combo1.Text = "Name" Then
cari = "Name = '" & Text1 & "'"
With Form1.Ado1.Recordset
    .Find cari
    If Not .EOF Then
    Label12.Caption = .Fields(0)
    Label13.Caption = .Fields(1)
    Label14.Caption = .Fields(2)
    Label15.Caption = .Fields(3)
    Label16.Caption = .Fields(4)
    Label17.Caption = .Fields(5)
    Label18.Caption = .Fields(7)
    Label19.Caption = .Fields(8)
    Label20.Caption = .Fields(9)
    Label21.Caption = .Fields(10)
    Label22.Caption = .Fields(11)

    Else
    MsgBox "Cannot Find " & " " & Text1 & " " & "Data Is Not Exist", vbCritical, "Not Found"
    Text1.Text = ""
    Text1.SetFocus
    End If
    End With
    End If

If Dir(App.Path & "\" & "Photos" & "\" & Label12.Caption & ".jpg") <> "" Then
Image1.Picture = LoadPicture(App.Path & "\" & "Photos" & "\" & Label12.Caption & ".jpg")
Else
Image1.Picture = LoadPicture(App.Path & "\" & "Photos" & "\" & "dodo.jpg")
End If
End Sub

Private Sub Form_Load()
Form1.Ado1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dodo.mdb;Persist Security Info=False"
Form1.Ado1.RecordSource = "Select * From Employee"
Form1.Ado1.Refresh
Set Form1.DataGrid1.DataSource = Form1.Ado1
With Combo1
    .AddItem "EID", 0
    .AddItem "Name", 1
    End With
Image1.Picture = LoadPicture(App.Path & "\" & "Photos" & "\" & "dodo.jpg")
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If Combo1.ListIndex = 0 Then
Text1.MaxLength = 5
    If (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyTab) Then
    Else
    Beep
    KeyAscii = 0
    End If
Else
Text1.MaxLength = 30
End If
End Sub
