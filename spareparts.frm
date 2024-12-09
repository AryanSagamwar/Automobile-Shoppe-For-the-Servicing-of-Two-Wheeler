VERSION 5.00
Begin VB.Form spareparts 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicing"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7380
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Spare Part"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   6375
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtcost 
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
         Height          =   360
         Left            =   2180
         TabIndex        =   20
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txtspare 
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
         Height          =   360
         Left            =   2180
         TabIndex        =   18
         Top             =   240
         Width           =   4000
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Spare Parts Used :"
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
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1980
      End
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
      Left            =   4920
      TabIndex        =   16
      Top             =   8280
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
      Left            =   3000
      TabIndex        =   15
      Top             =   8280
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
      Left            =   1080
      TabIndex        =   14
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Charges"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   6735
      Begin VB.CommandButton cmdcalculate 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   29
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtscharge 
         BackColor       =   &H00C0C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtfix 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   1680
         TabIndex        =   25
         Text            =   "1000"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   4920
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txttotal 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   5040
         TabIndex        =   12
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Spare Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fix Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF00FF&
         Caption         =   "TOTAL COST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.TextBox txtproblem 
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
      Height          =   780
      Left            =   2300
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2400
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
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtvehicle 
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
      Left            =   3615
      TabIndex        =   4
      Top             =   1800
      Width           =   1845
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
      Left            =   2295
      TabIndex        =   1
      Top             =   1200
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
      Left            =   2295
      TabIndex        =   0
      Text            =   "MH - 34 "
      Top             =   1800
      Width           =   885
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
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Servicing"
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
      Left            =   2880
      TabIndex        =   8
      Top             =   0
      Width           =   1650
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
      Left            =   855
      TabIndex        =   6
      Top             =   840
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
      Left            =   3135
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic Name"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1680
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
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1665
   End
End
Attribute VB_Name = "spareparts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbnumber_Click()
SQL = "Select * from Entry Where Job_id=" & cmbnumber.Text
        Set RS = New ADODB.Recordset
        RS.Open SQL, CON, 1, 3
        txtvehicle.Text = RS.Fields("Vehicle_Number")
        txtproblem.Text = RS.Fields("Problem")
        RS.Close
        cmdsave.Enabled = True
End Sub

Private Sub cmdadd_Click()
    List1.AddItem (txtspare.Text)
    List2.AddItem Val(txtcost.Text)
    txtspare.Text = ""
txtcost.Text = ""

End Sub

Private Sub cmdcalculate_Click()
Dim K As Integer
Dim I As Integer
For I = 0 To List2.ListCount
   If List2.List(I) <> "" Then
    K = K + Val(List2.List(I))
   End If
   
Next I
txtscharge.Text = K
txttotal.Text = Val(txtfix.Text) + K
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
 SQL = "Select Job_id from Entry where status='Pending'"
    Set RS = New ADODB.Recordset
    RS.Open SQL, CON, 1, 3
     While Not RS.EOF
        cmbnumber.AddItem (RS.Fields("Job_id"))
        RS.MoveNext
     Wend
     RS.Close
     cmdnew.Enabled = False
     cmdsave.Enabled = True
End Sub

Private Sub cmdsave_Click()
 SQL = "select * from Repair order by Job_Id"
            Set RS = New ADODB.Recordset
            RS.Open SQL, CON, 1, 3
            RS.AddNew
            RS.Fields("Job_id") = cmbnumber.Text
            RS.Fields("Mechanic_name") = txtname.Text
            RS.Fields("Problem") = txtproblem.Text
            RS.Fields("Vehicle_Number") = txtvehicle.Text
            RS.Fields("Fix_Charge") = txtfix.Text
            RS.Fields("Spare_Chnarge") = txtscharge.Text
            RS.Fields("Total_Charge") = txttotal.Text
            RS.Fields("Status") = "In"
            RS.Update
            RS.Close
            MsgBox "Vehicle is Repaired"
            SQL = "select * from Entry Where Job_id=" & cmbnumber.Text
            Set RS = New ADODB.Recordset
            RS.Open SQL, CON, 1, 3
            RS.Fields("Status") = "Repaired"
            RS.Update
            RS.Close
            Unload Me
            
End Sub

Private Sub txtcost_KeyPress(KeyAscii As Integer)
Call CHECKNUM(KeyAscii)
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Call CHECKTEXT(KeyAscii)
End Sub
