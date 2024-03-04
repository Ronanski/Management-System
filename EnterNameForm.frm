VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnterNameForm 
   Caption         =   "Enter details"
   ClientHeight    =   2184
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5784
   OleObjectBlob   =   "EnterNameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnterNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub filetemp_save()
    
    Dim myConn As ADODB.Connection
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileTemp_query As String
    Dim fileTemp, file_Format, filePath, file_Name As String
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    fileTemp = ws.Range("J26").Value
    file_Format = ws.Range("H29").Value
    filePath = ws.Range("H30").Value
    file_Name = ws.Range("H31").Value
    
    fileTemp_query = _
    "INSERT INTO KIODB.DBO.FileTemplates" & _
    "(FileTempName,FileFormat,FilePath,[FileName])" & _
    "Values ('" & fileTemp & "','" & file_Format & "','" & filePath & "','" & file_Name & "')"
    
    myConn.Execute fileTemp_query
    myConn.Close
End Sub
Sub geninfo_save()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim genTitle_query As String
    Dim addGen_query As String
    Dim incterm As String
    Dim excterm As String
    Dim varSet As Variant
    Dim incSet As Variant
    Dim excSet As Variant
    Dim termRow As Integer
    Dim myConn As ADODB.Connection
    Dim conn As String
    Dim genTempName As String
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    genTempName = ws.Range("J27").Value
    
    varSet = Array("CompanyName", "AttentionTo", "Currency", "BankAccount", _
            "AccountName", "Delivery", "DPTerms", "ProgressTerms", "Completion", _
            "ModeofPayment", "Cancellation", "Note1", "Note2", "Note3")
    
        genTitle_query = "INSERT INTO KIODB.DBO.Templates (TemplateName) values ('" & genTempName & "')"
        myConn.Execute genTitle_query
        
        For genRow = 0 To 13
            addGen_query = _
            "UPDATE KIODB.DBO.Templates " & _
            "Set " & varSet(genRow) & " = '" & ws.Cells(genRow + 14, 15).Value & "'" & _
            "WHERE TemplateName = '" & genTempName & "'"
            
            myConn.Execute addGen_query
        Next genRow
        
    incSet = Array("Inclusion1", "Inclusion2", "Inclusion3", "Inclusion4", "Inclusion5", _
            "Inclusion6", "Inclusion7", "Inclusion8", "Inclusion9", "Inclusion10", _
            "Inclusion11", "Inclusion12", "Inclusion13", "Inclusion14", "Inclusion15", _
            "Inclusion16", "Inclusion17", "Inclusion18", "Inclusion19", "Inclusion20", _
            "Inclusion21", "Inclusion22", "Inclusion23", "Inclusion24", "Inclusion25")
    
    excSet = Array("Exclusion1", "Exclusion2", "Exclusion3", "Exclusion4", "Exclusion5", _
            "Exclusion6", "Exclusion7", "Exclusion8", "Exclusion9", "Exclusion10", _
            "Exclusion11", "Exclusion12", "Exclusion13", "Exclusion14", "Exclusion15", _
            "Exclusion16", "Exclusion17", "Exclusion18", "Exclusion19", "Exclusion20", _
            "Exclusion21", "Exclusion22", "Exclusion23", "Exclusion24", "Exclusion25")
            
        For termRow = 0 To 24
            incterm = _
            "UPDATE KIODB.DBO.Templates " & _
            "Set " & incSet(termRow) & " = '" & ws.Cells(termRow + 35, 6).Value & "'" & _
            "WHERE TemplateName = '" & genTempName & "'"
            
            excterm = _
            "UPDATE KIODB.DBO.Templates " & _
            "Set " & excSet(termRow) & " = '" & ws.Cells(termRow + 35, 8).Value & "'" & _
            "WHERE TemplateName = '" & genTempName & "'"
            
            myConn.Execute incterm
            myConn.Execute excterm

        Next termRow
    
    Call RefreshGenList
    myConn.Close
End Sub
Private Sub savebtn_Click()
    
    user_response = MsgBox("Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Action")
    If user_response = vbYes Then
            
        If TB1.Value = "" And TB2.Value = "" Then
            MsgBox "Please provide the required information.", vbInformation, "Input Required"
        ElseIf TB1.Enabled = False Then
            Call geninfo_save
            MsgBox "General information template has been saved to the database!", vbInformation, "Success"
        ElseIf TB2.Enabled = False Then
            Call filetemp_save
            MsgBox "File information template has been saved to the database!", vbInformation, "Success"
        Else
            Call geninfo_save
            Call filetemp_save
            MsgBox "All necessary informations has been saved to the database!", vbInformation, "Success"
        End If
    End If
End Sub

Private Sub TB1_Change()
    ActiveWorkbook.Sheets("Registry").Range("J26").ClearContents
    ActiveWorkbook.Sheets("Registry").Range("J26").Value = TB1.Value
End Sub
Private Sub TB2_Change()
    ActiveWorkbook.Sheets("Registry").Range("J27").ClearContents
    ActiveWorkbook.Sheets("Registry").Range("J27").Value = TB2.Value
End Sub
Private Sub UserForm_Initialize()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    fileLoc = ws.Range("J14").Value
    geninfo = ws.Range("J15").Value
    
    If fileLoc = True And geninfo = False Then
        TB1.Enabled = True
        TB2.Enabled = False
    ElseIf geninfo = True And fileLoc = False Then
        TB2.Enabled = True
        TB1.Enabled = False
    Else
        TB1.Enabled = True
        TB2.Enabled = True
        TB1.SetFocus
    End If
        
End Sub
Private Sub UserForm_Terminate()
    
    
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub
Sub RefreshGenList()
    
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim gen_query As String
    
    QuotationForm.CB1.Clear
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
     
    gen_query = _
        "Select [TemplateName] from KIODB.dbo.Templates"
        
        myRecord.Open gen_query
        Do Until myRecord.EOF
            genTemp = myRecord.Fields(0).Value
            QuotationForm.CB1.AddItem genTemp
            myRecord.MoveNext
        Loop
        myRecord.Close
        myRecord.Open gen_query
            QuotationForm.CB1.Value = myRecord.Fields(0).Value
        myRecord.Close
    myConn.Close
    Set myConn = Nothing
    Set myRecord = Nothing
    
End Sub
