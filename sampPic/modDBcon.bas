Attribute VB_Name = "Module1"
Option Explicit
Dim myConnectionString As String
Public libcon As ADODB.Connection
Public profRS  As ADODB.Recordset
Public searchRs As ADODB.Recordset
Public rs As ADODB.Recordset
Public sqlstr As String



Sub dbconnect()

On Error GoTo dbconnect_Error


    Set libcon = New ADODB.Connection
          
      libcon.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
      "SERVER=localhost;" & _
      "DATABASE=samppic;" & _
      "UID=root;" & _
      "PASSWORD= root;" & _
      "OPTION=3;"
    
      libcon.CursorLocation = adUseClient
      libcon.Open


dbconnect_Exit:
Exit Sub

dbconnect_Error:
MsgBox "Unexpected error - " & Err.Number & vbCrLf & vbCrLf & Error$ & vbCrLf & vbCrLf & "Please contact Mr. Carmelo Alejo D. Bisquera." & vbCrLf & "or CALL 078 321 2262. Thank you!", vbExclamation, "NVSU Library System ver 1.1"
End
Resume dbconnect_Exit

End Sub

