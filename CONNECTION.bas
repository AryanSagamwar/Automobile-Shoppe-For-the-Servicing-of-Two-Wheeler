Attribute VB_Name = "Module1"
Public CON As New ADODB.Connection
Public RS As New ADODB.Recordset
Public SQL As String

Sub main()
SQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\SERVICE.mdb;Persist Security Info=False"
CON.Open SQL
splash.Show
End Sub

Function CHECKTEXT(K As Integer)
    ' For back Space ascii value = 8
    Select Case K
        Case 65 To 90, 97 To 122, 8, 32
                 K = K
        Case Else
                 K = 0
    End Select
    CHECKTEXT = K
End Function
Function CHECKNUM(K As Integer)
Select Case K
        Case 48 To 57, 8
                 K = K
        Case Else
                 K = 0
End Select
CHECKNUM = K
End Function
