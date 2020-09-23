VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter new data!"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
      Height          =   495
      Left            =   1590
      TabIndex        =   11
      Top             =   2190
      Width           =   1095
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      Height          =   495
      Left            =   4140
      TabIndex        =   6
      Top             =   2190
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   1590
      TabIndex        =   4
      Top             =   2190
      Width           =   1095
   End
   Begin VB.TextBox txtid 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   1620
      TabIndex        =   0
      Top             =   90
      Width           =   2235
   End
   Begin VB.TextBox txtmname 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   1620
      TabIndex        =   3
      Top             =   1590
      Width           =   3615
   End
   Begin VB.TextBox txtgname 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   1620
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtfname 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   1620
      TabIndex        =   1
      Top             =   570
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog dlgbrowse 
      Left            =   6390
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "Browse Picture"
      Height          =   495
      Left            =   5580
      TabIndex        =   5
      Top             =   2250
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MIDDLE NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   30
      TabIndex        =   10
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GIVEN NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   60
      TabIndex        =   9
      Top             =   1050
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FAMILY NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   570
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   30
      TabIndex        =   7
      Top             =   120
      Width           =   1545
   End
   Begin VB.Image Img 
      BorderStyle     =   1  'Fixed Single
      Height          =   1965
      Left            =   5580
      Top             =   150
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   1965
      Left            =   0
      Top             =   60
      Width           =   1605
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: CADBISQUERA
'ADDRESS: BAYOMBONG, NUEVA VIZCAYA, PHILIPPINES
'ENJOY!!!! GOODLUCK AND GOD BLESS

Option Explicit
Dim mystream As ADODB.Stream


Private Sub cmdbrowse_Click()
Dim mystream As String

dlgbrowse.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgbrowse.ShowOpen
    img.Picture = LoadPicture(dlgbrowse.FileName)
  
End Sub

Private Sub cmdnew_Click()
txtfname.Enabled = True
txtgname.Enabled = True
txtmname.Enabled = True
cmdnew.Visible = False
End Sub

Private Sub cmdsave_Click()
       
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim mystream As ADODB.Stream
        
        'start by loading an image into the database.
        'In addition to our connection object,
        'we will need a RecordSet object and a Stream object.
        'Let's begin by declaring these two objects:
        
        Set mystream = New ADODB.Stream
        mystream.Type = adTypeBinary


        'An ADO Stream object can handle both text
        'and binary data (and can therefore be used to get large text
        'fields as well as BLOB fields).
        'We have to specify which type of data we will be
        'dealing with using the adTypeBinary value in the Type parameter.
        
        'The first thing we need to do is open a blank recordset and add a new record to it.
        
        rs.Open "SELECT * FROM tblfiles WHERE 1=0", libcon, adOpenStatic, adLockOptimistic
        rs.AddNew
        
        mystream.Open
        mystream.LoadFromFile (dlgbrowse.FileName)
        
        rs!file_name = dlgbrowse.FileName
        rs!file_size = mystream.Size
        rs!file = mystream.Read
        rs!idnum = txtid.Text
        rs!fname = txtfname.Text
        rs!gname = txtgname.Text
        rs!mname = txtmname.Text
        rs.Update
        mystream.Close
        rs.Close
        libcon.Close
End Sub

Private Sub cmdsearch_Click()
 'Now that our image is in the table, we need to get it back out.
 'As we have covered them already, lets get the connection
 'and recordset objects inititalized right away
   Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim mystream As ADODB.Stream
        
        Set mystream = New ADODB.Stream
        'We have opened a connection and a recordset,
        'and also declared our stream. To get our file back out,
        'we open the stream, write to it from the recordset,
        'and then save the data to a file, as follows:
        
        mystream.Type = adTypeBinary
        sqlstr = "select * from tblfiles where idnum='" & (txtid.Text) & "'"
        rs.Open sqlstr, libcon, adOpenKeyset, adLockOptimistic
        mystream.Open
        mystream.Write rs!file
        
        'Note: make it sure that folder samppic with subfolder pics is created
        mystream.SaveToFile "c:\samppic\pics\mypic.jpg", adSaveCreateOverWrite
       
       If rs.BOF And rs.EOF Then
        MsgBox "Not exists!"
       Else
        dlgbrowse.FileName = "c:\samppic\pics\mypic.jpg"
        img.Picture = LoadPicture(dlgbrowse.FileName)
        txtid.Text = rs!idnum
        txtfname.Text = rs!fname
        txtgname.Text = rs!gname
        txtmname.Text = rs!mname
        
        'Note: it's up to you if you want to use the following code
        'mystream.Close
        'rs.Close
        'libcon.Close
      End If

End Sub

Private Sub Form_Load()
dbconnect

End Sub
