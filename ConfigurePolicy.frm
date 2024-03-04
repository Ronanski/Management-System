VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigurePolicy 
   Caption         =   "Configure Policy"
   ClientHeight    =   5844
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4860
   OleObjectBlob   =   "ConfigurePolicy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConfigurePolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN1_Click()

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
    myRecord.CursorType = adOpenStatic
    myRecord.ActiveConnection = myConn
    
    termName = CB1.Value
    descValue = TB1.Value
    
    'SQL Query
    myRecord.Open _
    "Select [Desc] From KIODB.DBO.Terms Where TermName = '" & termName & "'"
    
    'If TextBox is Empty
    If descValue = "" Then
        L4.Caption = "Term Description must not be empty."
    Else
        'Checking if term is existing
        Do Until myRecord.EOF
            If descValue = myRecord.Fields(0).Value Then
             L4.Caption = descValue & " is existing."
             Exit Sub
            End If
            myRecord.MoveNext
        Loop
        
            'Confirm Action
            user_response = MsgBox("Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Action")
            If user_response = vbYes Then
            
                'Saving to Database
                myConn.Execute _
                "Insert Into KIODB.DBO.Terms (TermName,[Desc])" & _
                "Values ('" & termName & "','" & descValue & "')"
                
                'Refresh ListBox
                Call CB1_Change
                
                'Updating Status
                L4.Caption = descValue & " has been added to list."
            Else
            L4.Caption = ""
            TB1.Text = ""
            GoTo x
            End If
    End If
x:
    myRecord.Close
    myConn.Close
    
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub

Private Sub BTN2_Click()

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
    
    'Checking if a term is selected
    If LB1.ListIndex > -1 Then
        selected_Desc = LB1.List(LB1.ListIndex)
        
        'Confirm Action
            user_response = MsgBox("Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Action")
            If user_response = vbYes Then
                myConn.Execute _
                "Delete From KIODB.DBO.Terms " & _
                "WHERE [Desc] = '" & selected_Desc & "' AND TermName = '" & termName & "'"
                Call CB1_Change
                L4.Caption = selected_Desc & " has been deleted!"
            Else
                LB1.ListIndex = -1
                L4.Caption = ""
                GoTo x
            End If
    Else
        L4.Caption = "Please select an item to be deleted."
    End If
x:
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
    myRecord.Open "Select * From KIODB.DBO.Terms WHERE TermName = '" & termName & "' "
        myRecord.MoveNext
        
    Do Until myRecord.EOF
        LB1.AddItem myRecord.Fields(1).Value
        myRecord.MoveNext
    Loop
    
    list_count = LB1.ListCount
    
    If list_count = 0 Then
        L4.Caption = ""
        LB1.Enabled = False
        With LB1
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
Private Sub LB1_Click()

    selected_Desc = LB1.List(LB1.ListIndex)
    L4.Caption = selected_Desc & " has been selected."
    
End Sub
Private Sub UserForm_Initialize()
    
    CB1.AddItem "Inclusion"
    CB1.AddItem "Exclusion"
    
End Sub
Private Sub UserForm_Terminate()
    
    'Refresh Quotationform Inclusion/Exclusion ComboBoxes after Exit
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim sql_query As String
    
    'Clear ComboBoxes First
    QuotationForm.pbx1.Clear
    QuotationForm.pbx2.Clear
    
    conn = _
                "Driver={SQL Server};" & _
                "Server=DCS;" & _
                "Database=KIODB;" & _
                "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    
    'Update both Inclusion/Exclusion
    For i = 1 To 2
    
        'SQL Query
        sql_query = _
        "Select [Desc] From KIODB.DBO.Terms Where TermName = '" & _
        Choose(i, "Inclusion", "Exclusion") & "'"
        
        'Updating both Inclusion/Exclusion
            myRecord.Open sql_query
            Do Until myRecord.EOF
                QuotationForm.Controls("pbx" & i).AddItem myRecord.Fields(0).Value
                myRecord.MoveNext
            Loop
            myRecord.Close
            
         'Initial Values for both Inclusion/Exclusion
            myRecord.Open sql_query
                QuotationForm.Controls("pbx" & i).Value = myRecord.Fields(0).Value
            myRecord.Close
            
    Next i
    
    myConn.Close
    Set myRecord = Nothing
    Set myConn = Nothing
    
End Sub
