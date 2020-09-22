VERSION 5.00
Begin VB.Form frmContacts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   3300
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6045
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Contacts"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5775
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtState 
         Height          =   285
         Left            =   3600
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtZip 
         Height          =   285
         Left            =   4080
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "City/St/Zip:"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Address:"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Name Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAdd 
         Caption         =   "New Record"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Record"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Record"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Searching with combos
' Author:            Adam Lankford
' Date constructed:  9-6-01
'*********************************************************************************
'This little app demos the use of ADO, Searches, and ComboBoxes.
' [Adam T.]
'---------------------------------------------------------------------------------
Option Explicit

Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection

Private Sub cmdSearch_Click()
    Call Search.Search(cboName.Text, rs, rs.Fields("NAME"))
End Sub

Private Sub Form_Load()
  On Error GoTo ErrHandler
  
  PATH = App.PATH
  
  Set conn = New ADODB.Connection
  Set rs = New ADODB.Recordset
  
  Call Utilities.OpenADO(rs, conn, "SELECT * FROM CONTACTS")
  Call PopulateCombo
  Call BindData
  
  Exit Sub
  
ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, Err.Number
    
End Sub

Private Sub BindData()
    Set txtName.DataSource = rs
    Set txtAddress.DataSource = rs
    Set txtCity.DataSource = rs
    Set txtState.DataSource = rs
    Set txtZip.DataSource = rs
    Set txtPhone.DataSource = rs
    Call SetFields
End Sub

Private Sub SetFields()
    txtName.DataField = "NAME"
    txtAddress.DataField = "ADDRESS"
    txtCity.DataField = "CITY"
    txtState.DataField = "STATE"
    txtZip.DataField = "ZIP"
    txtPhone.DataField = "PHONE"
End Sub

Private Sub PopulateCombo()
    Dim i As Long
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        cboName.AddItem (rs.Fields("NAME"))
        rs.MoveNext
    Next i
End Sub

Private Sub mnuAdd_Click()
    rs.AddNew
End Sub

Private Sub mnuDelete_Click()
    rs.Delete
    MsgBox "Record has been Deleted!", vbInformation + vbOKOnly, "====: Record Deletion :===="
    rs.MoveFirst
    ' Refresh combo after delete
    cboName.Clear
    Call PopulateCombo
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSave_Click()
    rs.Update
    MsgBox "Record has been saved!", vbInformation + vbOKOnly, "====: Record Saved :===="
    rs.MoveFirst
    ' Refresh combo after save incase new record is added...
    cboName.Clear
    Call PopulateCombo
End Sub
