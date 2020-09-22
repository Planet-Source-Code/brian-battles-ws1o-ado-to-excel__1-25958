VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form CopyToExcel 
   Caption         =   "   Copy Database Tables to Excel"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CopyToExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "  Click on a Table to Copy to Excel  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Reselect Database"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   6000
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   4875
         TabIndex        =   2
         Top             =   5640
         Width           =   4935
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5325
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Database File"
      FileName        =   "*.mdb"
      Filter          =   "Access Files (*.mdb)"
      FilterIndex     =   1
      FontName        =   "Arial"
      InitDir         =   "."
   End
End
Attribute VB_Name = "CopyToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' original code/project posted on PSC by Ian Mitchell
'
' (Ian -- Great job on the part that does all the HARD work with Excel!)
'
' Modified Aug 7, 2001, 7 PM by Brian Battles WS1O  brianb@cmtelephone.com
'
' I decided I wanted to make this more flexible by using ADO instead of DAO;
' that way we can use this on databases other than MS Access...
'
'  just be sure to set the necessary references:
'    Microsoft ADO 2.x library
'    OLE DB Service Component 1.0 Type Library
'  etc

Dim adoConn    As ADODB.Connection
Dim RS         As ADODB.Recordset
Dim strCaption As String
Dim SN         As String
Dim I          As Single
Dim Recs       As Integer
Dim Counter    As Integer
Dim BarString  As String
Dim MdbFile    As String
Dim Junk       As String
Dim strAdoConn As String

Private Type ExlCell
    Row As Long
    Col As Long
End Type
Private Sub Form_Load()
    
    LoadForm

Exit_Form_Load:

    On Error GoTo 0
    Exit Sub
    
Err_Form_Load:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Err_Form_Unload
    
    If Not (adoConn Is Nothing) Then
        adoConn.Close
        Set adoConn = Nothing
    End If
    
Exit_Form_Unload:

    On Error GoTo 0
    Exit Sub
    
Err_Form_Unload:

    Select Case Err
        Case 0, 91, 3704
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_Form_Unload
    End Select

End Sub
Private Sub Command1_Click()
            
    On Error GoTo Err_Command1_Click
    
    ' clear the progress bar
    UpdateProgress Picture1, 0
    ' hide the frame
    Frame1.Visible = False
    ' clear the listbox
    List1.Clear
    ' rerun the routine that initially populates the listbox
    LoadForm

Exit_Command1_Click:

    On Error GoTo 0
    Exit Sub
    
Err_Command1_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            Frame1.Visible = True
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_Command1_Click
    End Select

End Sub
Private Sub List1_Click()
    
    On Error GoTo Err_List1_Click
    
    Screen.MousePointer = vbHourglass
    Junk = List1.Text
    Set RS = New ADODB.Recordset
    RS.Open Junk, adoConn, adOpenStatic, adLockReadOnly, adCmdTable
    ToExcel RS, App.Path & "\wk.xls"
    
Exit_List1_Click:

    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub
    
Err_List1_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_List1_Click
    End Select
        
End Sub
Private Sub CopyRecords(RST As ADODB.Recordset, WS As Worksheet, StartingCell As ExlCell)
    
    Dim SomeArray() As Variant
    Dim Row         As Long
    Dim Col         As Long
    Dim Fd          As ADODB.Field
    
    On Error GoTo Err_CopyRecords
    
    ' check if recordset is not empty
    If RST.EOF And RST.BOF Then Exit Sub
    RST.MoveLast
    ReDim SomeArray(RST.RecordCount + 1, RST.Fields.Count)
    ' copy column headers to array
    Col = 0
    For Each Fd In RST.Fields
        SomeArray(0, Col) = Fd.Name
        Col = Col + 1
    Next
    ' copy recordset to some array
    RST.MoveFirst
    Recs = RST.RecordCount
    Counter = 0
    For Row = 1 To RST.RecordCount - 1
        Counter = Counter + 1
        If Counter <= Recs Then I = (Counter / Recs) * 100
        UpdateProgress Picture1, I
        For Col = 0 To RST.Fields.Count - 1
            SomeArray(Row, Col) = RST.Fields(Col).Value
            If IsNull(SomeArray(Row, Col)) Then _
            SomeArray(Row, Col) = ""
        Next
        RST.MoveNext
    Next
    ' The range should have the same number of
    ' rows and cols as in the recordset
    WS.Range(WS.Cells(StartingCell.Row, StartingCell.Col), _
        WS.Cells(StartingCell.Row + RST.RecordCount + 1, _
        StartingCell.Col + RST.Fields.Count)).Value = SomeArray

Exit_CopyRecords:

    On Error GoTo 0
    Exit Sub
    
Err_CopyRecords:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_CopyRecords
    End Select
        
End Sub
Private Sub ToExcel(SN As ADODB.Recordset, strCaption As String)
    
    Dim oExcel    As Object
    Dim objExlSht As Object ' OLE automation object
    Dim stCell    As ExlCell

    On Error GoTo Err_ToExcel
    
    DoEvents
        On Error Resume Next
        Set oExcel = GetObject(, "Excel.Application")
        ' if Excel is not launched start it
        If Err = 429 Then
            Err = 0
            Set oExcel = CreateObject("Excel.Application")
            ' can't create object
            If Err = 429 Then
                MsgBox Err & ": " & Error, vbExclamation + vbOKOnly
                Exit Sub
            End If
        End If
        oExcel.Workbooks.Add
        oExcel.Worksheets("sheet1").Name = strCaption
        Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
        stCell.Row = 1
        stCell.Col = 1
        ' place the fields across the top of the spreadsheet:
        CopyRecords SN, objExlSht, stCell
        ' give the user control
        oExcel.Visible = True
        oExcel.Interactive = True
        ' clean up (I test if objects are still "alive" to avoid errors):
        If Not (objExlSht Is Nothing) Then
            Set objExlSht = Nothing ' Remove object variable
        End If
        If Not (oExcel Is Nothing) Then
            Set oExcel = Nothing    ' Remove object variable
        End If
        If Not (SN Is Nothing) Then
            Set SN = Nothing        ' Remove snapshot object
        End If
    UpdateProgress Picture1, 100
    
Exit_ToExcel:

    On Error GoTo 0
    Exit Sub
    
Err_ToExcel:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_ToExcel
    End Select

End Sub
Sub UpdateProgress(PB As Control, ByVal Percent)
    
    Dim Num As String        'use percent
    
    On Error GoTo Err_UpdateProgress
    
    If Not PB.AutoRedraw Then    'picture in memory ?
        PB.AutoRedraw = -1       'no, make one
    End If
    PB.Cls                       'clear picture in memory
    PB.ScaleWidth = 100          'new sclaemodus
    PB.DrawMode = 10             'not XOR Pen Modus
    Num = BarString & Format$(Percent, "###") + "%"
    PB.CurrentX = 50 - PB.TextWidth(Num) / 2
    PB.CurrentY = (PB.ScaleHeight - PB.TextHeight(Num)) / 2
    PB.Print Num                 'print percent
    PB.Line (0, 0)-(Percent, PB.ScaleHeight), , BF
    PB.Refresh                   'show difference

Exit_UpdateProgress:

    On Error GoTo 0
    Exit Sub
    
Err_UpdateProgress:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_UpdateProgress
    End Select

End Sub
Private Sub LoadForm()

    On Error GoTo Err_LoadForm
    
    Picture1.Visible = True
    Frame1.Caption = "  Click on a Table to Copy to Excel  "
    
    
    GoTo TECHNIQUE_2
    ' CHANGE THE LINE ABOVE TO TRY THE FOLLOWING TECHNIQUES:
    '
    ' There are 2 ways we can do this;
    '   use Technique 1 for Access 2000 databases, or
    '   use Technique 2 for any ODBC data source (more generic)
    ' depends on what your application requires
    
TECHNIQUE_1:
    
    'set blue bar color
    Picture1.ForeColor = RGB(0, 0, 255)
    'open common dialog control
    CommonDialog1.Filter = "Access Files (*.mdb)"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.FileName = "*.mdb"
    CommonDialog1.ShowOpen
    MdbFile = (CommonDialog1.FileName)
    'Set up a DSN-less connection to our MS Access database
    Set adoConn = New ADODB.Connection
    adoConn.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & MdbFile   'App.Path & "\Examples.mdb"
    GoTo OPENTHEDATABASE
    
TECHNIQUE_2:

    strAdoConn = BuildAdoConnection("")
    'Set up a DSN-less connection to our ODBC database
    Set adoConn = New ADODB.Connection
    adoConn.ConnectionString = strAdoConn
    
OPENTHEDATABASE:

    adoConn.Open
    ' now we have a recordset containing the names of all the tables and queries in the database
    Set RS = adoConn.OpenSchema(adSchemaTables)
    'Now we loop through the recordset, row-by-row until we reach the End Of File
    Do Until RS.EOF
        ' make sure we're using the names of Tables that aren't
        ' System Object Tables, or tables that start with USys, or "Views" (queries)
        If RS.Fields("TABLE_TYPE") = "TABLE" Then
            ' populate the List Box
            If LCase$(Left$(RS.Fields("TABLE_NAME"), 4)) = "usys" Then
                ' skip system tables
                DoEvents
            Else
                List1.AddItem RS.Fields("TABLE_NAME")
            End If
        End If
        ' tell ADO to move to the next record or we'll be stuck
        'on the same row forever in an infinite loop
        RS.MoveNext
    Loop
    ' close objects when we're done and set to Nothing.
    If Not (RS Is Nothing) Then
        RS.Close
        Set RS = Nothing
    End If
    Frame1.Visible = True

Exit_LoadForm:

    On Error GoTo 0
    Exit Sub
    
Err_LoadForm:

    Select Case Err
        Case 0, 91 ' user cancelled
            Resume Next
        Case 32755, -2147467259, 3704
            Frame1.Visible = True
            Picture1.Visible = False
            Frame1.Caption = "  No Database Selected  "
            Resume Exit_LoadForm
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_LoadForm
    End Select
 
End Sub
Private Function BuildAdoConnection(ByVal ConnectionString As String) As String

    ' display the ADO Connection Window (ADO DB Designer)

    Dim dlViewConnection As MSDASC.DataLinks

    On Error GoTo Err_BuildAdoConnection
    
    Set adoConn = New ADODB.Connection
    If Not (Trim$(ConnectionString) = "") Then
        Set adoConn = New ADODB.Connection
        adoConn.ConnectionString = ConnectionString
        Set dlViewConnection = New MSDASC.DataLinks
        dlViewConnection.hWnd = Me.hWnd
            If dlViewConnection.PromptEdit(adoConn) Then
                BuildAdoConnection = adoConn.ConnectionString
            Else
                BuildAdoConnection = ConnectionString
            End If
        Set dlViewConnection = Nothing
        Set adoConn = Nothing
    Else
        Set dlViewConnection = New MSDASC.DataLinks
        dlViewConnection.hWnd = Me.hWnd
        Set adoConn = dlViewConnection.PromptNew
        BuildAdoConnection = adoConn.ConnectionString
        Set dlViewConnection = Nothing
        Set adoConn = Nothing
    End If

Exit_BuildAdoConnection:

    On Error Resume Next
        If Not (adoConn Is Nothing) Then
            Set adoConn = Nothing
        End If
        If Not (dlViewConnection Is Nothing) Then
            Set dlViewConnection = Nothing
        End If
    On Error GoTo 0
    Exit Function

Err_BuildAdoConnection:

    Select Case Err
        Case 0
            Resume Next
        Case -2147217805
            adoConn.ConnectionString = ""
            Resume
        Case 91
            Resume Exit_BuildAdoConnection
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & "  -  Advisory"
            Resume Exit_BuildAdoConnection
    End Select
   
End Function
