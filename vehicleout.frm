VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vehicleout 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Out"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7215
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   21
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   20
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   19
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox txttotal 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2400
      TabIndex        =   17
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtscharge 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2400
      TabIndex        =   15
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtfix 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2400
      TabIndex        =   13
      Text            =   "1000"
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtproblem 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1020
      Left            =   2895
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2880
      Width           =   4000
   End
   Begin VB.ComboBox cmbnumber 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtvehicle 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   3615
      TabIndex        =   4
      Top             =   2160
      Width           =   1845
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2295
      TabIndex        =   1
      Top             =   1440
      Width           =   4000
   End
   Begin VB.TextBox txtvehicle1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2295
      TabIndex        =   0
      Text            =   "MH - 34 "
      Top             =   2160
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker dtpservice 
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12640511
      CalendarTitleBackColor=   12640511
      Format          =   123666433
      UpDown          =   -1  'True
      CurrentDate     =   44935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Charges"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Charges"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Charges"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4080
      TabIndex        =   11
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Out"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Problems in vehicle"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   2490
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job ID :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   975
      TabIndex        =   6
      Top             =   840
      Width           =   1110
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
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2115
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2085
   End
End
Attribute VB_Name = "vehicleout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbnumber_Click()
    SQL = "Select * from Repair where Job_id=" & cmbnumber.Text
    Set RS = New ADODB.Recordset
    RS.Open SQL, CON, 1, 3
    txtname.Text = RS.Fields("Mechanic_name")
    txtvehicle.Text = RS.Fields("Vehicle_number")
    txtproblem.Text = RS.Fields("Problem")
    txtfix.Text = RS.Fields("Fix_Charge")
    txtscharge.Text = RS.Fields("Spare_Chnarge")
    txttotal.Text = RS.Fields("Total_Charge")
    RS.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
 SQL = "Select Job_id from Repair Where Status='In'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, CON, 1, 3
     While Not RS.EOF
        cmbnumber.AddItem (RS.Fields("Job_id"))
        RS.MoveNext
     Wend
     RS.Close
     cmdnew.Enabled = False
     cmdsave.Enabled = True
     dtpservice.Value = Date
End Sub

Private Sub cmdsave_Click()
    SQL = "select * from VehicleOut order by Job_Id"
            Set RS = New ADODB.Recordset
            RS.Open SQL, CON, 1, 3
            RS.AddNew
            RS.Fields("Job_id") = cmbnumber.Text
            RS.Fields("DDate") = dtpservice.Value
            RS.Fields("Mechanic_name") = txtname.Text
            RS.Fields("Problem") = txtproblem.Text
            RS.Fields("Vehicle_Number") = txtvehicle.Text
            RS.Fields("Fix_Charge") = txtfix.Text
            RS.Fields("Spare_Chnarge") = txtscharge.Text
            RS.Fields("Total_Charge") = txttotal.Text
            RS.Update
            RS.Close
            MsgBox "Vehicle is Released"
            SQL = "select * from Repair Where Job_id=" & cmbnumber.Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, CON, 1, 3
            RS.Fields("Status") = "Out"
            RS.Update
            RS.Close
              SQL = "select * from Entry Where Job_id=" & cmbnumber.Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, CON, 1, 3
            RS.Fields("Status") = "Out"
            RS.Update
            RS.Close
            Unload Me
End Sub

