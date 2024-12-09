VERSION 5.00
Begin VB.MDIForm mainform 
   BackColor       =   &H8000000C&
   Caption         =   "Automobile"
   ClientHeight    =   10440
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   16950
   LinkTopic       =   "MDIForm1"
   Picture         =   "mainform.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuvehicle 
      Caption         =   "Vehicle"
      Begin VB.Menu mnuentry 
         Caption         =   "Entry"
      End
      Begin VB.Menu mnuview 
         Caption         =   "View"
      End
      Begin VB.Menu mnustata 
         Caption         =   "Statistics"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear Database"
      End
   End
   Begin VB.Menu mnuservice 
      Caption         =   "Servicing"
      Begin VB.Menu mnuspare 
         Caption         =   "Spare Parts"
      End
   End
   Begin VB.Menu mnuout 
      Caption         =   "Vehile Out"
      Begin VB.Menu mnurelease 
         Caption         =   "Release Vehicle"
      End
      Begin VB.Menu mnuviewout 
         Caption         =   "View"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuclear_Click()
    frmdeletedata.Show
End Sub

Private Sub mnuentry_Click()
    entry.Show
    mnuservice.Enabled = False
    
End Sub

Private Sub mnuexit_Click()
    a = MsgBox(" Do you want to Exit ?", vbQuestion + vbYesNo, "Confirmation")
    If a = vbYes Then
        End
    End If
End Sub

Private Sub mnurelease_Click()
    vehicleout.Show
End Sub

Private Sub mnuspare_Click()
    spareparts.Show
End Sub

Private Sub mnustata_Click()
    statistics.Show
End Sub

Private Sub mnuview_Click()
    viewvehicle.Show
End Sub

Private Sub mnuviewout_Click()
    viewout.Show
End Sub
