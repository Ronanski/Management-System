VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TermConditionForm 
   Caption         =   "Configure Terms & Conditions"
   ClientHeight    =   6060
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5412
   OleObjectBlob   =   "TermConditionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TermConditionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADD_Click()
    
    'Adding the term to database & listbox
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    
    conn = _
                "Driver={SQL Server};" & _
                "Server=DCS;" & _
                "Database=KIODB;" & _
                "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    
    termName = CB1.Value
    descValue = TB1.Value
    
    'SQL Query
    myRecord.Open _
    "Select [Description] From KIODB.DBO.Warranty Where WarrantyType = '" & termName & "'"
    
    'If TextBox is Empty
    If descValue = "" Then
        L4.Caption = "Term Description must not be empty."
    Else
        'Checking if term is existing
        Do Until myRecord.EOF
            If descValue = myRecord.Fields(0).Value Then
             L4.Caption = descValue & " is existing."
             TB1.Text = ""
             Exit Sub
            End If
            myRecord.MoveNext
        Loop
        
            'Confirm Action
            user_response = MsgBox("Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Action")
            If user_response = vbYes Then
                
                'Saving to Database
                myConn.Execute _
                "Insert Into KIODB.DBO.Warranty (WarrantyType,[Description])" & _
                "Values ('" & termName & "','" & descValue & "')"
                
                'Refresh ListBox
                Call CB1_Change
                
                'Updating Status
                L4.Caption = descValue & " has been added."
            Else
            GoTo x
            End If
            
    End If
x:
    myRecord.Close
    myConn.Close
    
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub
Private Sub DEL_Click()
     
    'Delete in Listbox/Database
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    
    conn = _
                "Driver={SQL Server};" & _
                "Server=DCS;" & _
                "Database=KIODB;" & _
                "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    termName = CB1.Value
    
    'Checking if may selected term
    If LB1.ListIndex > -1 Then
        selected_Desc = LB1.List(LB1.ListIndex)
        
        'Confirm Action
        user_response = MsgBox("Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Action")
        If user_response = vbYes Then
            
            'Delete execution
            myConn.Execute _
            "Delete From KIODB.DBO.Warranty " & _
            "WHERE [Description] = '" & selected_Desc & "' AND WarrantyType = '" & termName & "'"
            Call CB1_Change
            L4.Caption = selected_Desc & " has been deleted!"
            TB1.Text = ""
        Else
            Exit Sub
        End If
    Else
        L4.Caption = "Please select an item to be deleted."
    End If
    
    myConn.Close
    
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub
Private Sub CB1_Change()

    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    
    'Refresh ListBox
    LB1.Clear
    L4.Caption = ""
    
    conn = _
                "Driver={SQL Server};" & _
                "Server=DCS;" & _
                "Database=KIODB;" & _
                "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    termName = CB1.Value
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    myRecord.Open "Select * From KIODB.DBO.Warranty WHERE WarrantyType = '" & termName & "' "
    
    'Skipping Description = NONE
        myRecord.MoveNext
        
    'Populate Listbox
    
    Do Until myRecord.EOF
        LB1.AddItem myRecord.Fields(1).Value
        myRecord.MoveNext
    Loop
    
    'If no existing terms found
    
    list_count = LB1.ListCount
    
    If list_count = 0 Then
        L4.Caption = ""
        With LB1
        .Enabled = False
        .AddItem "No existing terms found."
        End With
    Else
        LB1.Enabled = True
    End If
    
    myRecord.Close
    myConn.Close
    
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub
Private Sub UserForm_Initialize()

    u = Array _
    ("Delivery Terms", "Down Payment Terms", "Progress Billing Terms", "Completion Terms", _
    "Mode of Payment", "Cancellation Terms", "Note 1", "Note 2", "Note 3")

    For i = LBound(u) To UBound(u)
        CB1.AddItem u(i)
    Next i
    
End Sub
Private Sub LB1_Click()

    selected_Desc = LB1.List(LB1.ListIndex)
    L4.Caption = selected_Desc & " has been selected."
    
End Sub
Private Sub UserForm_Terminate()
    
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim sql_query As String
    Dim i As Integer
    
    'Clear ComboBoxes First
    For i = 1 To 9
        QuotationForm.Controls("CBX" & i).Clear
    Next i
    
    conn = _
                "Driver={SQL Server};" & _
                "Server=DCS;" & _
                "Database=KIODB;" & _
                "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    
    'Update CBX1-CBX9
    For i = 1 To 9
    
        'SQL Query
        sql_query = _
        "Select [Description] From KIODB.DBO.Warranty Where WarrantyType = '" & _
        Choose(i, "Delivery Terms", "Down Payment Terms", "Progress Billing Terms", _
        "Completion Terms", "Mode of Payment", "Cancellation Terms", _
        "Note 1", "Note 2", "Note 3") & "'"
        
        'Updating ComboBoxes
            myRecord.Open sql_query
            Do Until myRecord.EOF
                QuotationForm.Controls("cbx" & i).AddItem myRecord.Fields(0).Value
                myRecord.MoveNext
            Loop
            myRecord.Close
         'Initial Values for CBX1 to CBX9
            myRecord.Open sql_query
                QuotationForm.Controls("cbx" & i).Value = myRecord.Fields(0).Value
            myRecord.Close
            
    Next i
    
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub


