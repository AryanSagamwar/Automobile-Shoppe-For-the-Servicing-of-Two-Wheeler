VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form entry 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Entry"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7290
   Begin VB.TextBox txtid 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2055
      TabIndex        =   22
      Top             =   1080
      Width           =   1365
   End
   Begin VB.TextBox txtvehicle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   3720
      TabIndex        =   19
      Top             =   4200
      Width           =   1845
   End
   Begin VB.CommandButton cmdold 
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpservice 
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12640511
      CalendarTitleBackColor=   12640511
      Format          =   123994113
      CurrentDate     =   44935
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   4000
   End
   Begin VB.TextBox txtvehicle1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2400
      TabIndex        =   4
      Text            =   "MH - 34 "
      Top             =   4200
      Width           =   885
   End
   Begin VB.TextBox txtproblem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5040
      Width           =   4000
   End
   Begin VB.TextBox txtchasis 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   6000
      Width           =   4000
   End
   Begin VB.TextBox txtaddress 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2400
      TabIndex        =   1
      Top             =   2760
      Width           =   4000
   End
   Begin VB.TextBox txtcontact 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2400
      TabIndex        =   0
      Top             =   3480
      Width           =   4000
   End
   Begin VB.ComboBox cmbnumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   21
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   23
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   2310
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   105
      TabIndex        =   12
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   105
      TabIndex        =   11
      Top             =   4320
      Width           =   1665
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3720
      TabIndex        =   10
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Problems in vehicle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   100
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Chasis number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   100
      TabIndex        =   8
      Top             =   6000
      Width           =   1545
   End
   Begin VB.Label Custom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   105
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact  Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   3600
      Width           =   2415
   End
End
Attribute VB_Name = "entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbnumber_Click()
        SQL = "Select * from Entry Where Vehicle_Number='" & Trim$(cmbnumber.Text) & "'"
        Set RS = New ADODB.Recordset
        RS.Open SQL, CON, 1, 3
        txtname.Text = RS.Fields("Customer_name")
        txtaddress.Text = RS.Fields("Customer_address")
        txtcontact.Text = RS.Fields("Contact_Number")
        txtchasis.Text = RS.Fields("Chasis_number")
        RS.Close
        cmdsave.Enabled = True
End Sub

Private Sub cmdclose_Click()
    Unload Me
    mainform.mnuservice.Enabled = True
End Sub

Private Sub cmdnew_Click()
    dtpservice.Value = Date
    SQL = "Select * from Entry order by Job_id"
    Set RS = New ADODB.Recordset
    RS.Open SQL, CON, 1, 3
    If RS.RecordCount = 0 Then
        txtid.Text = 1
    Else
        RS.MoveLast
        txtid.Text = RS.Fields("Job_id") + 1
    End If
    RS.Close
    cmdsave.Enabled = True
    cmdold.Enabled = False
    cmdnew.Enabled = False
End Sub

Private Sub cmdold_Click()
dtpservice.Value = Date
SQL = "Select * from Entry order by Job_id"
    Set RS = New ADODB.Recordset
    RS.Open SQL, CON, 1, 3
    If RS.RecordCount = 0 Then
        txtid.Text = 1
    Else
        RS.MoveLast
        txtid.Text = RS.Fields("Job_id") + 1
    End If
    RS.Close
    cmdnew.Enabled = False
    cmbnumber.Visible = True
    txtvehicle.Visible = False
    Call filldata
End Sub
Private Sub filldata()
    SQL = "Select Vehicle_Number from Entry where Status='Out'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, CON, 1, 3
     While Not RS.EOF
        cmbnumber.AddItem (RS.Fields("Vehicle_Number"))
        RS.MoveNext
     Wend
     RS.Close
End Sub
Private Sub cmdsave_Click()
    If txtname.Text = "" Then
        MsgBox " Fill all the Details"
    Else
        SQL = "Select * from Entry Where Vehicle_Number='" & Trim$(txtvehicle.Text) & "'"
        Set RS = New ADODB.Recordset
        RS.Open SQL, CON, 1, 3
        If RS.RecordCount > 0 Then
            MsgBox " Vehicle Number Repeated, Search in Old"
            Exit Sub
            RS.Close
        Else
            RS.Close
            SQL = "select * from Entry order by Job_Id"
            Set RS = New ADODB.Recordset
            RS.Open SQL, CON, 1, 3
            RS.AddNew
            RS.Fields("Job_id") = txtid.Text
            RS.Fields("Customer_name") = txtname.Text
            RS.Fields("Customer_address") = txtaddress.Text
            RS.Fields("Contact_Number") = txtcontact.Text
            RS.Fields("Vehicle_Number") = txtvehicle.Text
            RS.Fields("Date_of_servicing") = dtpservice.Value
            RS.Fields("Problem") = txtproblem.Text
            RS.Fields("Chasis_number") = txtchasis.Text
            RS.Fields("status") = "Pending"
            RS.Update
            RS.Close
            MsgBox " New Vehicle is Entered"
            Call clear
            cmdnew.Enabled = True
            cmdold.Enabled = True
            cmdsave.Enabled = False
        End If
        End If
                   
End Sub
Private Sub clear()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        End If
    Next
        
        
        
End Sub

Private Sub txtchasis_KeyPress(KeyAscii As Integer)
Call CHECKNUM(KeyAscii)
End Sub

Private Sub txtcontact_KeyPress(KeyAscii As Integer)
Call CHECKNUM(KeyAscii)
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Call CHECKTEXT(KeyAscii)
End Sub
