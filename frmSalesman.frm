VERSION 5.00
Begin VB.Form frmSalesMan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New SalesMan"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmSalesman.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   840
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   840
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   840
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   840
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "Last"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   840
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   840
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "Previous"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   840
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "First"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   840
   End
   Begin VB.TextBox Text2 
      DataField       =   "VendorAddress1"
      DataSource      =   "mrsVendors"
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   2475
   End
   Begin VB.TextBox Text1 
      DataField       =   "VendorName"
      DataSource      =   "mrsVendors"
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   1995
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   960
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   450
      Left            =   2400
      TabIndex        =   10
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SMan Name"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   885
   End
   Begin VB.Label lblVendor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SMan Code:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "frmSalesMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Email phil_zac@yahoo.com or philip707@hotmail.com
Private WithEvents Cn As ADODB.Connection
Attribute Cn.VB_VarHelpID = -1
Private WithEvents rsRecordSet As ADODB.Recordset
Attribute rsRecordSet.VB_VarHelpID = -1
Private WithEvents rstempcode As ADODB.Recordset
Attribute rstempcode.VB_VarHelpID = -1
Dim mblnAdd As Boolean
Dim mblnEdit As Boolean
Dim strEditRec As String


Private Sub Form_Load()
   Set Cn = New ADODB.Connection
   Cn.CursorLocation = adUseClient
    'Cn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Persist Security info = false; Data source = D:\datafiles\Accountss.mdb"
     
     Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\Accountss.mdb"

    Set rsRecordSet = New ADODB.Recordset
    rsRecordSet.Open "SELECT * From Salesman ORDER BY Smancode", Cn, adOpenStatic, adLockOptimistic

If rsRecordSet.EOF = True And rsRecordSet.BOF = True Then
 CmdEdit.Enabled = False
 CmdDelete.Enabled = False
 Exit Sub
End If
    
    SetButtons True
    mblnAdd = False
    LoadDataInControls
    
End Sub

Private Sub SetButtons(bVal As Boolean)
  CmdAdd.Visible = bVal
  CmdEdit.Visible = bVal
  CmdSave.Visible = Not bVal
  CmdCancel.Visible = Not bVal
  CmdDelete.Visible = bVal
  CmdClose.Visible = bVal
  
  CmdNext.Visible = bVal
  CmdFirst.Visible = bVal
  CmdLast.Visible = bVal
  CmdPrevious.Visible = bVal
  End Sub
Private Sub LockTbox()
Text1.Locked = True
Text2.Locked = True
End Sub
Private Sub UnLockTbox()
Text1.Locked = False
Text2.Locked = False
End Sub
Private Sub LoadDataInControls()
 If rsRecordSet.BOF = True Or rsRecordSet.EOF = True Then
        Exit Sub
     End If
    Text1.Text = rsRecordSet!SmanCode & ""
    Text2.Text = rsRecordSet!SalesMan & ""
 End Sub
Private Sub WriteData()
    rsRecordSet!SmanCode = Text1.Text
    rsRecordSet!SalesMan = Text2.Text
End Sub
Private Sub ClearControls()
    Text1.Text = ""
    Text2.Text = ""
  End Sub
 Private Sub cmdFirst_Click()
   mblnAdd = False
    
    If rsRecordSet.BOF = False Then
       rsRecordSet.MoveFirst
    
    ElseIf rsRecordSet.BOF = True And rsRecordSet.EOF = True Then
     MsgBox " There is no data in the Recordset!", , "Oops"
   
    End If
    Text1.SetFocus
    LoadDataInControls
    End Sub
Private Sub CmdPrevious_Click()
  mblnAdd = False
  If rsRecordSet.BOF = False Then
     rsRecordSet.MovePrevious
   If rsRecordSet.BOF Then rsRecordSet.MoveFirst
    Else
   If rsRecordSet.EOF = True Then
      MsgBox "There is no data in the Recordset!", , "Oops"
   Else
   rsRecordSet.MoveLast
   End If
  End If
    LoadDataInControls
    Text1.SetFocus
End Sub
Private Sub CmdNext_Click()
mblnAdd = False
If rsRecordSet.EOF = False Then
rsRecordSet.MoveNext
If rsRecordSet.EOF Then rsRecordSet.MoveLast
Else
If rsRecordSet.BOF Then
MsgBox "There is no data in the Recordset!", , "Oops"
Else
rsRecordSet.MoveLast
End If
End If
   LoadDataInControls
   Text1.SetFocus
End Sub
Private Sub CmdLast_Click()
mblnAdd = False
If rsRecordSet.EOF = False Then
rsRecordSet.MoveLast

ElseIf rsRecordSet.BOF = True And rsRecordSet.EOF = True Then
MsgBox " There is no Data in the RecordSet!", , "Oops"

 End If
     LoadDataInControls
     Text1.SetFocus
End Sub
Private Sub cmdAdd_Click()
    Label1.Visible = True
    Label1.Caption = "Please Press Enterkey After Write SalesMan Code"
    mblnAdd = True
    UnLockTbox
    ClearControls
    SetButtons False
    Text1.SetFocus
   End Sub
Private Sub cmdSave_Click()
Label1.Visible = False
On Error GoTo ErrorHandler
 If mblnEdit = True Then
    mblnEdit = False
    If strEditRec <> Text1.Text Then
    LoadDataInControls
  End If
      End If
    If ValidData Then
        If mblnAdd Then rsRecordSet.AddNew
        WriteData
        rsRecordSet.Update
        mblnAdd = False
        SetButtons True
        Text1.Enabled = True
        'Text1.SetFocus
        CmdAdd.SetFocus
        LockTbox
        'ClearControls
        CmdEdit.Enabled = True
        CmdDelete.Enabled = True
      End If
    Exit Sub
    rsRecordSet.Close
    rsRecordSet.Open
   
ErrorHandler:
    DisplayErrorMsg
    If rsRecordSet.EditMode = adEditAdd Then rsRecordSet.CancelUpdate
End Sub
Private Sub CmdEdit_Click()
    mblnEdit = True
    UnLockTbox
    Text1.Enabled = False
    strEditRec = Text1
    SetButtons False
    Text2.SetFocus
End Sub
Private Sub CmdCancel_Click()
Text1.Enabled = True
LockTbox
Label1.Visible = False
rsRecordSet.CancelUpdate
LoadDataInControls
SetButtons True
mblnAdd = False
mblnEdit = False
End Sub
Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
    If MsgBox("Are you sure you want to delete this record?", _
              vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
          On Error Resume Next
             rsRecordSet.Delete
             rsRecordSet.MoveNext
          If rsRecordSet.EOF Then
            'rsRecordSet.Requery     '(its only necessary when you are using Severside cursors)
             rsRecordSet.MoveLast
         End If
             LoadDataInControls
            
    End If
    If rsRecordSet.BOF = True Then
       Call ClearControls
    MsgBox "There is no data in the recordset!", , "Oops!"
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    End If
    SetButtons True
    
    Text1.SetFocus
    Exit Sub
ErrorHandler:
    DisplayErrorMsg
End Sub
Private Sub CmdClose_Click()
    rsRecordSet.Close
    Cn.Close
    Set rsRecordSet = Nothing
    Set Cn = Nothing
    Unload frmSalesMan
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
    SendKeys "{tab}"
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim ChkCode As Boolean
   If mblnEdit = False And mblnAdd = False Then
   MsgBox "Click Edit to edit records or Add to enter new records.", vbOKOnly + vbInformation
   Exit Sub
   End If
   If KeyCode = vbKeyReturn Then
    ChkCode = VrfyUnique
    If ChkCode = True Then
    MsgBox "  Code has already been taken ! ", vbExclamation + vbOKOnly
    Text1.SetFocus
   End If
 End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
 If mblnEdit = False And mblnAdd = False Then
   MsgBox "Click Edit to edit records or Add to enter new records.", vbOKOnly + vbInformation
   Exit Sub
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys "{tab}"
    End If
End Sub
Private Function ValidData() As Boolean
    Dim strMessage As String
    If Text1.Text = "" Then
        Text1.SetFocus
        strMessage = "You must enter a SalesMan Code."
    ElseIf Text2 = "" Then
        Text2.SetFocus
        strMessage = "You must enter SalesMan Name."
     Else
        ValidData = True
    End If
    If Not ValidData Then
        MsgBox strMessage, vbOKOnly
    End If
End Function
Public Function VrfyUnique() As Boolean
    Set rstempcode = New ADODB.Recordset
    rstempcode.Open "SELECT SManCode from Salesman where SManCode ='" & Trim$(Text1.Text) & "'", Cn, adOpenStatic, adLockOptimistic
    If rstempcode.RecordCount > 0 Then
        VrfyUnique = True
      Else
        VrfyUnique = False
    End If
    Set rstempcode = Nothing
End Function
Private Sub DisplayErrorMsg()
    MsgBox "Error Code: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description & vbCrLf & _
        "Source: " & Err.Source, vbOKOnly + vbCritical
End Sub
Private Sub Text2_GotFocus()
Text2.SelLength = Len(Text2.Text)
End Sub
