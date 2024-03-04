VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuotationForm 
   Caption         =   "Quotation Form"
   ClientHeight    =   11532
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   16512
   OleObjectBlob   =   "QuotationForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "QuotationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '
    ' THIS VBA PROGRAM IS NOW ON MY GITHUB
    ' All changes will be publish on my repository for easy track
    '

Sub RefreshGenList()
    
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim gen_query As String
    
    CB1.Clear
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
            CB1.AddItem genTemp
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
Sub DeleteTemplate()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim myConn As ADODB.Connection
    Dim selectedTemp As String
    Dim delgen As String
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
       
    selectedTemp = ws.Range("F14").Value
    
    delgen = _
    "DELETE FROM KIODB.DBO.Templates WHERE TemplateName = '" & selectedTemp & "'"
    
    If CB1.Value = "None" Then
        MsgBox "Selected template cannot be deleted.", vbExclamation, "Try Again"
    Else
        'Confirm action
        user_response = MsgBox("Delete the selected Template?" _
        & Chr(10) & "Selected Template: " & selectedTemp, vbInformation _
        + vbYesNo, "Confirm Action")
        
        If user_response = vbYes Then
            myConn.Execute delgen
            MsgBox "Template deleted successfully!", vbInformation, "Success"
            Call RefreshGenList
        End If
    End If
    
    myConn.Close
    Set myConn = Nothing
End Sub
Private Sub applycreate_Click()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim missingFields As String
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    fileLoc = ws.Range("J14").Value
    geninfo = ws.Range("J15").Value
    
    
    'Check first for missing required user input field
    For rowNum = 15 To 30
    If rowNum < 23 Or rowNum > 24 Then 'This skips Letter(Optional)
        If IsEmpty(Cells(rowNum, 6)) Then
            If missingFields = "" Then 'Initial run, empty pa then get first value at column 5
                missingFields = Chr(10) & "- " & Cells(rowNum, 5)
            Else
                missingFields = missingFields & Chr(10) & "- " & Cells(rowNum, 5) 'dito may value na, add nalang ng add
            End If
        End If
    End If
    Next rowNum
    
    If CB4.Value = "Customer Format" Then
        c = 0
        For rowNum = 14 To 23
            If Cells(rowNum, 8).Value = True Then
                c = c + 1
            End If
        Next rowNum

        If c = 0 Then
             missingFields = missingFields & Chr(10) & "- " & "Quotation Format Option"
        End If

    End If

    If Range("H26").Value = "None" Then
    missingFields = missingFields & Chr(10) & "- " & "Bank details"
    End If

    For rowNum = 30 To 31
        If IsEmpty(Cells(rowNum, 8)) Then
            If missingFields = "" Then
                missingFields = Cells(rowNum, 7)
            Else
                missingFields = missingFields & Chr(10) & "- " & Cells(rowNum, 7)
            End If
        End If
    Next rowNum
    
    'Terms and Condition
    For p = 1 To 9
        If Controls("cx" & p).Value = True Then
            If Cells(p + 15, 10) = "None" Then
            missingFields = missingFields & Chr(10) & "- " & Cells(p + 15, 9)
            End If
        End If
    Next p
    
    'If there are missing fields
    If missingFields <> "" Then
        MsgBox "Some of the required fields are not met: " & missingFields 'idisplay na nito lahat
    Else
        If fileLoc = False And geninfo = False Then
            MsgBox "Please choose saving option.", _
            vbInformation, "Checkbox Required"
        Else
            EnterNameForm.Show
        End If
    End If
    
End Sub

Private Sub B2_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim myConn As ADODB.Connection
    Dim selectedTemp As String
    Dim delgen As String
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
       
    selectedTemp = ws.Range("F14").Value
    
    delgen = _
    "DELETE FROM KIODB.DBO.Templates WHERE TemplateName = '" & selectedTemp & "'"
    
    If CB1.Value = "None" Then
        MsgBox "Selected template cannot be deleted.", vbExclamation, "Try Again"
    Else
        'Confirm action
        user_response = MsgBox("Delete the selected Template?" _
        & Chr(10) & "Selected Template: " & selectedTemp, vbInformation _
        + vbYesNo, "Confirm Action")
        
        If user_response = vbYes Then
            myConn.Execute delgen
            MsgBox "Template deleted successfully!", vbInformation, "Success"
            Call RefreshGenList
        End If
    End If
    
    myConn.Close
    Set myConn = Nothing
    
End Sub
Private Sub UserForm_Initialize()

    'Declaration of Registry
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim myConn As ADODB.Connection
        Dim myRecord As ADODB.Recordset
        Dim sql_query As String
        Dim i As Integer
        
        Set wb = ActiveWorkbook
        Set ws = wb.Sheets("Registry")
        
        conn = _
            "Driver={SQL Server};" & _
            "Server=DCS;" & _
            "Database=KIODB;" & _
            "Trusted_Connection=Yes;"
        
        Set myConn = New ADODB.Connection
        myConn.Open conn
        
        Set myRecord = New ADODB.Recordset
        myRecord.ActiveConnection = myConn
        
    ws.Range("F35:F59").ClearContents
    ws.Range("H35:H59").ClearContents
    ws.Range("F14:F30").ClearContents
    ws.Range("H14:H31").ClearContents
    ws.Range("J14:J25").ClearContents
    
    'Populate BANK LIST COMBOBOX
        bankset = Array("Banko De Oro", "Security Bank", _
        "Bank of the Philippine Islands", "MetroBank")
        
        For counter = LBound(bankset) To UBound(bankset)
            CB5.AddItem bankset(counter)
        Next counter
        
        LB1.ColumnWidths = "143"
            
    'Fill Labels on BANK DETAILS
        selected_index = LB1.ListIndex
    
        If selected_index = -1 Then
            For c = 1 To 3
                Controls("L" & c).Caption = "None"
            Next c
        End If
    
    'Initial values for WARRANTY COVERAGE from Registry Sheet
        For coverage = 21 To 27
            WC.AddItem ws.Cells(coverage, 1).Value
        Next coverage
    
    'SQL Query for INCLUSION ComboBox
        inc_query = _
        "Select [Desc] From KIODB.DBO.Terms where TermName = 'Inclusion'"
    
    'Initial Value for INCLUSION ComboBox
        myRecord.Open inc_query
            pbx1.Value = myRecord.Fields(0).Value
        myRecord.Close
        
        myRecord.Open inc_query
            Do Until myRecord.EOF
                inc_val = myRecord.Fields(0).Value
                pbx1.AddItem inc_val
                myRecord.MoveNext
            Loop
        myRecord.Close
        
    'SQL Query for EXCLUSION ComboBox
        exc_query = _
        "Select [Desc] From KIODB.DBO.Terms where TermName = 'Exclusion'"
        
    'Initial Value for EXCLUSION ComboBox
        myRecord.Open exc_query
        pbx2.Value = myRecord.Fields(0).Value
        myRecord.Close
        
        myRecord.Open exc_query
            Do Until myRecord.EOF
                exc_val = myRecord.Fields(0).Value
                pbx2.AddItem exc_val
                myRecord.MoveNext
            Loop
        myRecord.Close
    
    'Update CBX1-CBX9
        For i = 1 To 9
        
            'SQL Query
            cbx_query = _
            "Select [Description] From KIODB.DBO.Warranty Where WarrantyType = '" & _
            Choose(i, "Delivery Terms", "Down Payment Terms", "Progress Billing Terms", _
            "Completion Terms", "Mode of Payment", "Cancellation Terms", _
            "Note 1", "Note 2", "Note 3") & "'"
            
            'Updating ComboBoxes
                myRecord.Open cbx_query
                Do Until myRecord.EOF
                    QuotationForm.Controls("cbx" & i).AddItem myRecord.Fields(0).Value
                    myRecord.MoveNext
                Loop
                myRecord.Close
             'Initial Values for CBX1 to CBX9
                myRecord.Open cbx_query
                    QuotationForm.Controls("cbx" & i).Value = myRecord.Fields(0).Value
                myRecord.Close
        Next i
    
    'Populate File Template
            ft_query = _
            "Select [FileTempName] from KIODB.dbo.FileTemplates"
            myRecord.Open ft_query
            'Default value
            FT1.Value = myRecord.Fields(0).Value
            'Populate
            Do Until myRecord.EOF
                ftName = myRecord.Fields(0).Value
                FT1.AddItem ftName
                myRecord.MoveNext
            Loop
            myRecord.Close

    'Populate General Template
            gen_query = _
            "Select [TemplateName] from KIODB.dbo.Templates"
            myRecord.Open gen_query
            CB1.Value = myRecord.Fields(0).Value
            Do Until myRecord.EOF
                genTemp = myRecord.Fields(0).Value
                CB1.AddItem genTemp
                myRecord.MoveNext
            Loop
            myRecord.Close
    'Initial setting for Terms & Condition
    
        For q = 1 To 9
            Controls("cbx" & q).Enabled = False
        Next q
    
    'Populate Quotation Format Template
        With CB4
        .AddItem "Company Format"
        .AddItem "Customer Format"
        .Value = "Company Format"
        End With
    'Populate Currency & Unit of Time
        
        For rowNum = 2 To 5
            cur.AddItem ws.Cells(rowNum, 1).Value
        Next rowNum
        
        uot_set = Array("Day", "Week", "Month", "Year")
        For a = 2 To 3
            For b = LBound(uot_set) To UBound(uot_set)
                Controls("CB" & a).AddItem uot_set(b)
            Next b
        Next a
        
    'Populate Filetype
        With FT2
        .AddItem ".pdf"
        .AddItem ".xlsx"
        End With
        
    'Quotation Format Initial Values
        For v = 1 To 10
            ws.Cells(v + 13, 8).Value = Controls("C" & v).Value
        Next v
        
    'File Location & Database Setup Initial Values
        ws.Range("J14").Value = fl1.Value
        ws.Range("J15").Value = fl2.Value
        
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub
Private Sub bankreg_Click()
    BankRegisterForm.Show
End Sub
Private Sub CB4_Change()
    If CB4.Value = "Company Format" Then
        For i = 1 To 10
            Controls("C" & i).Enabled = False
            Controls("C" & i).Value = False
        Next i
    Else
        For i = 1 To 10
            Controls("C" & i).Enabled = True
            Controls("C" & i).Value = True
        Next i
    End If
    
    ActiveWorkbook.Sheets("Registry").Range("F27").Value = CB4.Value
    
End Sub

Private Sub CB5_Change()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim sql_query As String
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    LB1.Clear
    LB1.Enabled = True
    
    'Registry Sheet Values
    ActiveWorkbook.Sheets("Registry").Range("H24").Value = CB5.Value
    
    'Refresh Labels
        selected_index = LB1.ListIndex
    
        If selected_index = -1 Then
            For c = 1 To 3
                Controls("L" & c).Caption = "None"
            Next c
        End If
        
    ActiveWorkbook.Sheets("Registry").Range("H25").Value = L1.Caption
    ActiveWorkbook.Sheets("Registry").Range("H26").Value = L2.Caption
    ActiveWorkbook.Sheets("Registry").Range("H27").Value = L3.Caption
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    
    bankName = CB5.Value
    
    sql_query = _
    "Select AccountName, BankAccount From KIODB.DBO.BankDetails Where BankName = '" & bankName & "'"
    
    If bankName <> "" Then
    
        'Populate the ListBox
        myRecord.Open sql_query
        myRecord.MoveNext
        
            Do Until myRecord.EOF
                combinedValue = myRecord.Fields(1).Value & "-" & myRecord.Fields(0).Value
                LB1.AddItem combinedValue
                myRecord.MoveNext
            Loop
         myRecord.Close
         list_count = LB1.ListCount
         
         If list_count = 0 Then
         LB1.AddItem "No Accounts Found."
         LB1.Enabled = False
         GoTo x:
         End If
    End If
     
x:
    
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub

Private Sub delbank_Click()
        
    'Routine for Delete Bank Account Button
    selected_index = LB1.ListIndex
    
    If CB5.Value = "" Then
        MsgBox "Please select bank first.", vbInformation, "Information"
    ElseIf selected_index = -1 Then
        MsgBox "Please select account to be deleted.", vbInformation, "Information"
    Else
        
        Dim myConn As ADODB.Connection
        Dim myRecord As ADODB.Recordset
        Dim del_query As String
        
        conn = _
            "Driver={SQL Server};" & _
            "Server=DCS;" & _
            "Database=KIODB;" & _
            "Trusted_Connection=Yes;"
        
        Set myConn = New ADODB.Connection
        myConn.Open conn
        
        index_value = LB1.List(selected_index)
        get_index = InStr(1, index_value, "-")
        acctNum = Mid(index_value, 1, get_index - 1)
        
        del_query = _
        "Delete From KIODB.DBO.BankDetails Where BankAccount = '" & acctNum & "'"
        
        'Confirm Action before deleting
            user_response = MsgBox("Do you want to delete selected account?", vbQuestion + vbYesNo, "Confirm Action")
            If user_response = vbYes Then
                myConn.Execute del_query
                Call CB5_Change
                MsgBox "Delete successful!", vbInformation, "Information"
                myConn.Close
            Else
                GoTo x
            End If
    End If
x:
    
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub

Private Sub LB1_Click()
    
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim bank_query As String
    Dim get_num As String
    Dim get_name As String
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    
    'Updating label values when selecting a bank account
    selected_index = LB1.ListIndex
    If selected_index > -1 Then
        index_value = LB1.List(selected_index)
        get_index = InStr(1, index_value, "-")
        get_name = Mid(index_value, get_index + 1, 20)
        get_num = Mid(index_value, 1, get_index - 1)
        
        L2.Caption = get_num
        L3.Caption = get_name
        
'        bankset = Array("Banko De Oro", "Security Bank", _
'        "Bank of the Philippine Islands", "MetroBank")
'
'        swiftset = Array("BNORPHMMXXX", "SETCPHMMXXX", "BOPIPHMMXXX", "MBTCPHMMXXX")
'
'        For counter = LBound(bankset) To UBound(bankset)
'            bank = bankset(counter)
'            If bank = CB5.Value Then
'                swiftcode = swiftset(counter)
'                L1.Caption = swiftcode
'            End If
'        Next counter
'    End If
        
        bank_query = _
        "Select Swiftcode from KIODB.dbo.BankDetails Where BankAccount = '" & get_num & "'"
        
        myRecord.Open bank_query
        L1.Caption = myRecord.Fields(0).Value
        myRecord.Close
        End If
        
    ActiveWorkbook.Sheets("Registry").Range("H25").Value = L1.Caption
    ActiveWorkbook.Sheets("Registry").Range("H26").Value = L2.Caption
    ActiveWorkbook.Sheets("Registry").Range("H27").Value = L3.Caption
    
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub

Private Sub removecontent_Click()
    
    'Routing for remove content button removing terms for Inclusion/Exclusion
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rowNum1 As Integer
    Dim rownum2 As Integer
    
    Set wb = Workbooks("Management System")
    Set ws = wb.Sheets("Registry")
    
    'Removing Selected Items
    
    If ILB.ListIndex < 0 And ELB.ListIndex < 0 Then
        MsgBox "No terms to remove.", vbInformation, "Information"
        
    'Inclusion selected only
    ElseIf ILB.ListIndex > -1 And ELB.ListIndex = -1 Then
    
        'Remove term on ILB
        ilb_rem = ILB.List(ILB.ListIndex, 0)
        ILB.RemoveItem ILB.ListIndex
        
        'Remove term on Inclusion Registry
        For rowNum1 = 35 To 59
            If ws.Cells(rowNum1, 6).Value = ilb_rem Then
                ws.Cells(rowNum1, 6).ClearContents
                ws.Range(Cells(rowNum1 + 1, 6), Cells(59, 6)).Select
                Selection.Copy
                ws.Cells(rowNum1, 6).Select
                ws.Paste
            End If
        Next rowNum1
        
    'Exclusion selected only
    ElseIf ELB.ListIndex > -1 And ILB.ListIndex = -1 Then
    
        'Remove term on ELB
        elb_rem = ELB.List(ELB.ListIndex, 0)
        ELB.RemoveItem ELB.ListIndex
        
        'Remove term on Exclusion Registry
        For rownum2 = 35 To 59
            If ws.Cells(rownum2, 8).Value = elb_rem Then
                ws.Cells(rownum2, 8).ClearContents
                ws.Range(Cells(rownum2 + 1, 8), Cells(59, 8)).Select
                Selection.Copy
                ws.Cells(rownum2, 8).Select
                ws.Paste
            End If
        Next rownum2
    
    'Both exclusion and inclusion has selected terms
    ElseIf ILB.ListIndex > -1 And ELB.ListIndex > -1 Then
    
        'Remove term on ILB
        ilb_rem = ILB.List(ILB.ListIndex, 0)
        ILB.RemoveItem ILB.ListIndex
        
        'Remove term on Inclusion Registry
        For rowNum1 = 35 To 59
            If ws.Cells(rowNum1, 6).Value = ilb_rem Then
                ws.Cells(rowNum1, 6).ClearContents
                ws.Range(Cells(rowNum1 + 1, 6), Cells(59, 6)).Select
                Selection.Copy
                ws.Cells(rowNum1, 6).Select
                ws.Paste
            End If
        Next rowNum1
            
        'Remove term on ILB
        elb_rem = ELB.List(ELB.ListIndex, 0)
        ELB.RemoveItem ELB.ListIndex
        
        'Remove term on Exclusion Registry
        For rownum2 = 35 To 59
            If ws.Cells(rownum2, 8).Value = elb_rem Then
                ws.Cells(rownum2, 8).ClearContents
                ws.Range(Cells(rownum2 + 1, 8), Cells(59, 8)).Select
                Selection.Copy
                ws.Cells(rownum2, 8).Select
                ws.Paste
                GoTo x
            End If
        Next rownum2
    End If
        rowNum1 = 35
        rownum2 = 35
x:
        ILB.ListIndex = -1
        ELB.ListIndex = -1
    
End Sub
Private Sub addcontent_Click()
    
    'Adding content to Inclusion/Exclusion Listbox
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks("Management System")
    Set ws = wb.Sheets("Registry")
    
    'Add Content Button
    inc_count = ILB.ListCount - 1
    exc_count = ELB.ListCount - 1
    
    For counter1 = 0 To inc_count
        x = ILB.List(counter1, 0)
        If pbx1.Value = x Then
        GoTo exc
        End If
    Next counter1
    
    'For InclusionLB & Registry
    If inc_count < 25 Then
        ILB.AddItem pbx1.Value
        ws.Cells(counter1 + 35, 6).Value = ILB.List(counter1, 0)
        pbx1.Value = "None"
    Else
        MsgBox "Maximum Inclusion Terms has been reached."
    End If
    
exc:
    
    For counter2 = 0 To exc_count
        y = ELB.List(counter2, 0)
        If pbx2.Value = y Then
            Exit Sub
        End If
    Next counter2
    
    'For ExclusionLB & Registry
    If exc_count < 25 Then
        ELB.AddItem pbx2.Value
        ws.Cells(counter2 + 35, 8).Value = ELB.List(counter2, 0)
        pbx2.Value = "None"
    Else
        MsgBox "Maximum Exclusion Terms has been reached."
    End If
End Sub

Private Sub UserForm_Terminate()
    
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub
Private Sub FT1_Change()
    
    'Routine for FileTemplate section
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim file_query As String
    Dim i As Integer
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
        
    ws.Range("H28").Value = FT1.Value
    ftName = ws.Range("H28").Value
    file_query = _
    "Select * from KIODB.dbo.FileTemplates where FileTempName = '" & ftName & "'"

    myRecord.Open file_query
    If Not myRecord.EOF Then
        For i = 0 To 3
            ws.Cells(28 + i, 8).Value = myRecord.Fields(i).Value
        Next i
    End If
    myRecord.Close

    FT2.Value = ws.Range("H29").Value
    txb1.Text = ws.Range("H30").Value
    txb2.Text = ws.Range("H31").Value
    
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub
Private Sub CB1_Change()

    'Routine for Template section
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Registry")
    
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim termVal As String
    Dim termCat As String
    Dim termRow As Integer
    Dim inc_val As String
    Dim exc_val As String
    Dim missingValues As String
    Dim checker As Integer
    Dim incMiss As Boolean
    Dim excMiss As Boolean
    
    Call ClearVal
    
    conn = _
        "Driver={SQL Server};" & _
        "Server=DCS;" & _
        "Database=KIODB;" & _
        "Trusted_Connection=Yes;"

    Set myConn = New ADODB.Connection
    myConn.Open conn

    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn

    ws.Range("F14").Value = CB1.Value
    gtName = ws.Range("F14").Value
    gen_query = _
    "Select * from KIODB.dbo.Templates where TemplateName = '" & gtName & "'"
    
    'DATABASE to EXCEL SHEET REGISTRY
    
    myRecord.Open gen_query
    If Not myRecord.EOF Then
        'for QUOTATION CACHE UPDATE
        ws.Range("F22").Value = myRecord.Fields(1).Value
        ws.Range("F23").Value = myRecord.Fields(2).Value
        ws.Range("F25").Value = myRecord.Fields(3).Value
        ws.Range("H26").Value = myRecord.Fields(4).Value
        ws.Range("H27").Value = myRecord.Fields(5).Value
        'for TERMS & CONDITIONS
        For a = 1 To 9
            ws.Cells(15 + a, 10).Value = myRecord.Fields(5 + a).Value
        Next a
        'for INCLUSION
        For b = 1 To 25
            ws.Cells(34 + b, 6).Value = myRecord.Fields(14 + b).Value
        Next b
        'for EXCLUSION
        For c = 1 To 25
            ws.Cells(34 + c, 8).Value = myRecord.Fields(39 + c).Value
        Next c
    End If
    myRecord.Close
    
    'REGISTRY SHEET to USERFORM
    T4 = ws.Range("F22").Value 'COMPANY NAME
    T5 = ws.Range("F23").Value 'ATTENTION TO
    cur = Trim(ws.Range("F25").Value) 'CURRENCY
    
    'BANKDETAILS CHECK IF NO VALUE
    bankAcct = ws.Range("H26").Value
    If Not bankAcct = "" Then
        'ADD TO LIST IF HAVE VALUE
        LB1.AddItem ws.Range("H26").Value & " - " & ws.Range("H27").Value
    End If
    
    checker = 0
    missingValues = ""
    
    ' for Warranty Terms, Inclusions & Exclusions
    For ct = 1 To 9
        termVal = ws.Cells(18 + ct, 15).Value
        termCat = ws.Cells(18 + ct, 14).Value
        
        If termVal = "None" Then
            Controls("cx" & ct).Value = False
        Else
            termQuery = "Select * From KIODB.dbo.Warranty where [Description] = '" & termVal & "'"
            myRecord.Open termQuery
            
            If myRecord.EOF Then
                ' Value not found in the database
                Controls("cbx" & ct).Value = "None"
                Controls("cx" & ct).Value = False
                checker = checker + 1
                missingValues = missingValues & "- " & termCat & ": " & termVal & Chr(10)
            Else
                ' Value found in the database
                Controls("cbx" & ct).Value = ws.Cells(18 + ct, 15).Value
                Controls("cx" & ct).Value = True
            End If
            myRecord.Close
        End If
    Next ct
    
    ' for INC
    If Not ws.Range("F35").Value = "" Then
        termRow = 35
        Do Until IsEmpty(ws.Cells(termRow, 6).Value)
            inc_val = ws.Cells(termRow, 6).Value
            checkIncQuery = "Select * From KIODB.dbo.Terms where [Desc] = '" & inc_val & "'"
            myRecord.Open checkIncQuery
            If myRecord.EOF Then
                ' Value not found in the database
                missingValues = missingValues & "- " & "Inclusion Term: " & inc_val & Chr(10)
                incMiss = True
            Else
                ' Value found in the database
                ILB.AddItem myRecord.Fields(1).Value
            End If
            myRecord.Close
            termRow = termRow + 1
        Loop
    End If
    
    'for EXC
    If Not ws.Range("H35").Value = "" Then
        termRow = 35
        Do Until IsEmpty(ws.Cells(termRow, 8).Value)
            exc_val = ws.Cells(termRow, 8).Value
            checkExcQuery = "Select * From KIODB.dbo.Terms where [Desc] = '" & exc_val & "'"
            myRecord.Open checkExcQuery
            If myRecord.EOF Then
                ' Value not found in the database
                missingValues = missingValues & "- " & "Exclusion Term: " & exc_val & Chr(10)
                excMiss = True
            Else
                ' Value found in the database
                ELB.AddItem myRecord.Fields(1).Value
            End If
            myRecord.Close
            termRow = termRow + 1
        Loop
    End If
    
    ' Overall Message
    If checker > 0 Or excMiss Or incMiss Then
        userResponse = MsgBox("Some of the following terms are not in the database:" _
                        & Chr(10) & missingValues & Chr(10) & "Delete selected Template?", _
                               vbQuestion + vbYesNo, "Warning")
        
        If userResponse = vbYes Then
            Call DeleteTemplate
        Else
            If incMiss = True Or excMiss = True Then
            ILB.Clear
            ELB.Clear
            End If
        End If
        
    End If
    
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub
Private Sub WC_Change()
    ActiveWorkbook.Sheets("Registry").Range("J25").Value = WC.Text
End Sub
Private Sub configtermcondi_Click()
    TermConditionForm.Show
End Sub
Private Sub configpolicy_Click()
    ConfigurePolicy.Show
End Sub
Private Sub cx1_Click()
    x = cx1.Value
    
    If x = True Then
        cbx1.Enabled = True
    Else
        cbx1.Value = "None"
        cbx1.Enabled = False
    End If
End Sub
Private Sub cx2_Click()
    x = cx2.Value
    
    If x = True Then
        cbx2.Enabled = True
    Else
        cbx2.Value = "None"
        cbx2.Enabled = False
    End If

End Sub
Private Sub cx3_Click()
    x = cx3.Value
    
    If x = True Then
        cbx3.Enabled = True
    Else
        cbx3.Value = "None"
        cbx3.Enabled = False
    End If
End Sub
Private Sub cx4_Click()
    x = cx4.Value
    
    If x = True Then
        cbx4.Enabled = True
    Else
        cbx4.Value = "None"
        cbx4.Enabled = False
    End If
End Sub
Private Sub cx5_Click()
    x = cx5.Value
    
    If x = True Then
        cbx5.Enabled = True
    Else
        cbx5.Value = "None"
        cbx5.Enabled = False
    End If
End Sub
Private Sub cx6_Click()
    x = cx6.Value
    
    If x = True Then
        cbx6.Enabled = True
    Else
        cbx6.Value = "None"
        cbx6.Enabled = False
    End If
End Sub
Private Sub cx7_Click()
    x = cx7.Value
    
    If x = True Then
        cbx7.Enabled = True
    Else
        cbx7.Value = "None"
        cbx7.Enabled = False
    End If
End Sub
Private Sub cx8_Click()
    x = cx8.Value
    
    If x = True Then
        cbx8.Enabled = True
    Else
        cbx8.Value = "None"
        cbx8.Enabled = False
    End If
End Sub
Private Sub cx9_Click()
    x = cx9.Value
    
    If x = True Then
        cbx9.Enabled = True
    Else
        cbx9.Value = "None"
        cbx9.Enabled = False
    End If
End Sub
Private Sub pbx1_Change()
    ILB.ListIndex = -1
End Sub
Private Sub pbx2_Change()
    ELB.ListIndex = -1
End Sub
Private Sub O1_Change()
    ActiveWorkbook.Sheets("Registry").Range("F15").Value = O1.Caption
End Sub
Private Sub O2_Change()
    ActiveWorkbook.Sheets("Registry").Range("F15").Value = O2.Caption
End Sub
Private Sub O3_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O3.Caption
End Sub
Private Sub O4_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O4.Caption
End Sub
Private Sub O5_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O5.Caption
End Sub
Private Sub O6_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O6.Caption
End Sub
Private Sub O7_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O7.Caption
End Sub
Private Sub O8_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O8.Caption
End Sub
Private Sub O9_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O9.Caption
End Sub
Private Sub O10_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O10.Caption
End Sub
Private Sub O11_Change()
    ActiveWorkbook.Sheets("Registry").Range("F16").Value = O11.Caption
End Sub

Private Sub CB2_Change()
    ActiveWorkbook.Sheets("Registry").Range("F19").Value = CB2.Value
End Sub
Private Sub CB3_Change()
    ActiveWorkbook.Sheets("Registry").Range("F21").Value = CB3.Value
End Sub
Private Sub T1_Change()
    ActiveWorkbook.Sheets("Registry").Range("F17").Value = T1.Text
End Sub
Private Sub T2_Change()
    ActiveWorkbook.Sheets("Registry").Range("F18").Value = T2.Text
    t2_val = ActiveWorkbook.Sheets("Registry").Range("F18").Value
    If Not IsNumeric(t2_val) Then
        MsgBox "Please put number values only.", vbCritical, "Input error"
        ActiveWorkbook.Sheets("Registry").Range("F18").ClearContents
        T2.Value = ActiveWorkbook.Sheets("Registry").Range("F18").Value
    End If
End Sub
Private Sub T3_Change()
    ActiveWorkbook.Sheets("Registry").Range("F20").Value = T3.Text
    t3_val = ActiveWorkbook.Sheets("Registry").Range("F20").Value
    If Not IsNumeric(t3_val) Then
        MsgBox "Please put number values only.", vbCritical, "Input error"
        ActiveWorkbook.Sheets("Registry").Range("F20").ClearContents
        T3.Value = ActiveWorkbook.Sheets("Registry").Range("F20").Value
    End If
End Sub
Private Sub T4_Change()
    ActiveWorkbook.Sheets("Registry").Range("F22").Value = T4.Text
End Sub
Private Sub T5_Change()
    ActiveWorkbook.Sheets("Registry").Range("F23").Value = T5.Text
End Sub
Private Sub T6_Change()
    ActiveWorkbook.Sheets("Registry").Range("F24").Value = T6.Text
End Sub
Private Sub T7_Change()
    ActiveWorkbook.Sheets("Registry").Range("F26").Value = T7.Text
End Sub
Private Sub Cur_Change()
    ActiveWorkbook.Sheets("Registry").Range("F25").Value = cur.Value
End Sub
Private Sub FNT1_Change()
    ActiveWorkbook.Sheets("Registry").Range("F28").Value = FNT1.Value
End Sub
Private Sub PNT1_Change()
    ActiveWorkbook.Sheets("Registry").Range("F29").Value = PNT1.Text
    pnt_val = ActiveWorkbook.Sheets("Registry").Range("F29").Value
    If Not IsNumeric(pnt_val) Then
        MsgBox "Please put number values only.", vbCritical, "Input error"
        ActiveWorkbook.Sheets("Registry").Range("F29").ClearContents
        PNT1.Value = ActiveWorkbook.Sheets("Registry").Range("F29").Value
    End If
End Sub
Private Sub EAT1_Change()
    ActiveWorkbook.Sheets("Registry").Range("F30").Value = EAT1.Value
End Sub
Private Sub C1_Change()
    ActiveWorkbook.Sheets("Registry").Range("H14").Value = C1.Value
End Sub
Private Sub C2_Change()
    ActiveWorkbook.Sheets("Registry").Range("H15").Value = C2.Value
End Sub
Private Sub C3_Change()
    ActiveWorkbook.Sheets("Registry").Range("H16").Value = C3.Value
End Sub
Private Sub C4_Change()
    ActiveWorkbook.Sheets("Registry").Range("H17").Value = C4.Value
End Sub
Private Sub C5_Change()
    ActiveWorkbook.Sheets("Registry").Range("H18").Value = C5.Value
End Sub
Private Sub C6_Change()
    ActiveWorkbook.Sheets("Registry").Range("H19").Value = C6.Value
End Sub
Private Sub C7_Change()
    ActiveWorkbook.Sheets("Registry").Range("H20").Value = C7.Value
End Sub
Private Sub C8_Change()
    ActiveWorkbook.Sheets("Registry").Range("H21").Value = C8.Value
End Sub
Private Sub C9_Change()
    ActiveWorkbook.Sheets("Registry").Range("H22").Value = C9.Value
End Sub
Private Sub C10_Change()
    ActiveWorkbook.Sheets("Registry").Range("H23").Value = C10.Value
End Sub
Private Sub FT2_Change()
    ActiveWorkbook.Sheets("Registry").Range("H29").ClearContents
    ActiveWorkbook.Sheets("Registry").Range("H29").Value = FT2.Value
End Sub
Private Sub txb1_Change()
    ActiveWorkbook.Sheets("Registry").Range("H30").ClearContents
    ActiveWorkbook.Sheets("Registry").Range("H30").Value = txb1.Text
End Sub
Private Sub txb2_Change()
    ActiveWorkbook.Sheets("Registry").Range("H31").ClearContents
    ActiveWorkbook.Sheets("Registry").Range("H31").Value = txb2.Text
End Sub
Private Sub CBX1_Change()
    ActiveWorkbook.Sheets("Registry").Range("J16").ClearContents
    ActiveWorkbook.Sheets("Registry").Range("J16").Value = cbx1.Value
End Sub
Private Sub CBX2_Change()
    ActiveWorkbook.Sheets("Registry").Range("J17").Value = cbx2.Value
End Sub
Private Sub CBX3_Change()
    ActiveWorkbook.Sheets("Registry").Range("J18").Value = cbx3.Value
End Sub
Private Sub CBX4_Change()
    ActiveWorkbook.Sheets("Registry").Range("J19").Value = cbx4.Value
End Sub
Private Sub CBX5_Change()
    ActiveWorkbook.Sheets("Registry").Range("J20").Value = cbx5.Value
End Sub
Private Sub CBX6_Change()
    ActiveWorkbook.Sheets("Registry").Range("J21").Value = cbx6.Value
End Sub
Private Sub CBX7_Change()
    ActiveWorkbook.Sheets("Registry").Range("J22").Value = cbx7.Value
End Sub
Private Sub CBX8_Change()
    ActiveWorkbook.Sheets("Registry").Range("J23").Value = cbx8.Value
End Sub
Private Sub CBX9_Change()
    ActiveWorkbook.Sheets("Registry").Range("J24").Value = cbx9.Value
End Sub
Private Sub fl1_Change()
    ActiveWorkbook.Sheets("Registry").Range("J14").Value = fl1.Value
End Sub
Private Sub fl2_Change()
    ActiveWorkbook.Sheets("Registry").Range("J15").Value = fl2.Value
End Sub
Sub ClearVal()
    LB1.Clear
    ILB.Clear
    ELB.Clear
    CB2.Value = ""
    CB3.Value = ""
    CB5.Value = ""
    CB4.Value = "Company Format"
    PNT1.Text = ""
    FNT1.Text = ""
    EAT1.Text = ""
    WC.Value = ""
    pbx1.Value = "None"
    pbx2.Value = "None"
    
    For counter = 1 To 7
        Controls("T" & counter).Text = ""
    Next counter
    
    For counter = 1 To 11
        Controls("O" & counter).Value = False
    Next counter
    
    selected_index = LB1.ListIndex
    
        If selected_index = -1 Then
            For c = 1 To 3
                Controls("L" & c).Caption = "None"
            Next c
        End If
End Sub











