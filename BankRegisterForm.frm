VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BankRegisterForm 
   Caption         =   "Register Bank Account"
   ClientHeight    =   2796
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5052
   OleObjectBlob   =   "BankRegisterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BankRegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Change()
    L4.Caption = ""
    TB1.Text = ""
End Sub

Private Sub regbtn_Click()
    
    'Registration
    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim sql_query As String
    Dim counter As Integer
    
    conn = _
            "Driver={SQL Server};" & _
            "Server=DCS;" & _
            "Database=KIODB;" & _
            "Trusted_Connection=Yes;"
    
    Set myConn = New ADODB.Connection
    myConn.Open conn
    
    Set myRecord = New ADODB.Recordset
    myRecord.ActiveConnection = myConn
    
    bankName = CB1.Value
    bankNum = TB1.Value
    acctName = TB2.Value
    
    'Check if there is a selected bank
    If bankName = "" Then
        L4.Caption = "Please select a bank."
        
    'Check if the Account Name is empty
    ElseIf acctName = "" Then
        L4.Caption = "Account name must not be empty."
    
    Else
        'Check if Account Number is valid.
        If IsNumeric(bankNum) Then
        
        bankset = Array("Banko De Oro", "Security Bank", _
        "Bank of the Philippine Islands", "MetroBank")
        
        swiftset = Array("BNORPHMMXXX", "SETCPHMMXXX", "BOPIPHMMXXX", "MBTCPHMMXXX")
        
            'Setting corresponding SWIFTCODE for selected bank
            For counter = LBound(bankset) To UBound(bankset)
                bank = bankset(counter)
                If bank = CB1.Value Then
                    swiftcode = swiftset(counter)
                End If
            Next counter
        
            'Convert bankNum into String
            
            index_value = bankNum
            get_index = InStr(1, index_value)
            bankNum_int = CDbl(Mid(index_value, get_index + 1, 20))
            
            
            'SQL QUERIES
            sql_query = _
            "Insert Into KIODB.DBO.BankDetails (BankName,Swiftcode,BankAccount,AccountName)" & _
            "Values ('" & bankName & "','" & swiftcode & "','" & bankNum_int & "','" & acctName & "')"
            
            exist_query = _
            "Select BankAccount From KIODB.DBO.BankDetails Where BankName = '" & bankName & "'"
            
            'Check for existing bank account
            myRecord.Open exist_query
            Do Until myRecord.EOF
                If bankNum_int = myRecord.Fields(0).Value Then
                    L4.Caption = "Account already existing."
                    myRecord.Close
                    GoTo x
                End If
                myRecord.MoveNext
            Loop
            
            
            'Confirm Action before registration
            user_response = MsgBox("Do you want to proceed?", vbQuestion + vbYesNo, "Confirm Action")
            If user_response = vbYes Then
                myConn.Execute sql_query
                L4.Caption = "Registration successful!"
                myConn.Close
                RefreshListBox
                GoTo x
            Else
                GoTo x
            End If
        Else
            L4.Caption = "Registration unsuccessful. Please enter a valid Account Number."
        End If
    End If
x:
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub
Private Sub TB1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    L4.Caption = ""
    check_num = TB1.Value
    If Not IsNumeric(check_num) Then
        MsgBox "Please input number values only.", vbExclamation, "Input Error"
        TB1.Value = ""
        
    End If
    
End Sub
Private Sub UserForm_Initialize()
    
    'Populate Bank List ComboBox
        bankset = Array("Banko De Oro", "Security Bank", _
        "Bank of the Philippine Islands", "MetroBank")
        
        For counter = LBound(bankset) To UBound(bankset)
            CB1.AddItem bankset(counter)
        Next counter
    
End Sub
Private Sub UserForm_Terminate()
    
    Set myRecord = Nothing
    Set myConn = Nothing
End Sub
Sub RefreshListBox()

    Dim myConn As ADODB.Connection
    Dim myRecord As ADODB.Recordset
    Dim sql_query As String
    
    If QuotationForm.CB5.Value <> "" Then
    
        QuotationForm.LB1.Clear
        QuotationForm.LB1.Enabled = True
        
        'Refresh Labels
            selected_index = QuotationForm.LB1.ListIndex
        
            If selected_index = -1 Then
                For c = 1 To 3
                    QuotationForm.Controls("L" & c).Caption = "None"
                Next c
            End If
            
        conn = _
            "Driver={SQL Server};" & _
            "Server=DCS;" & _
            "Database=KIODB;" & _
            "Trusted_Connection=Yes;"
        
        Set myConn = New ADODB.Connection
        myConn.Open conn
        
        Set myRecord = New ADODB.Recordset
        myRecord.ActiveConnection = myConn
        
        bankName = CB1.Value
        acctName = TB2.Value
        
        If bankName = "" Or acctName = "" Then
            Exit Sub
        Else
        
            sql_query = _
            "Select AccountName, BankAccount From KIODB.DBO.BankDetails Where BankName = '" & bankName & "'"
            
            'Populate the ListBox
            myRecord.Open sql_query
            myRecord.MoveNext
            
                Do Until myRecord.EOF
                    combinedValue = myRecord.Fields(1).Value & "-" & myRecord.Fields(0).Value
                    QuotationForm.LB1.AddItem combinedValue
                    myRecord.MoveNext
                Loop
             myRecord.Close
             
             list_count = QuotationForm.LB1.ListCount
             
             If list_count = 0 Then
             QuotationForm.LB1.AddItem "No Accounts Found."
             QuotationForm.LB1.Enabled = False
             End If
        myConn.Close
        End If
    End If
    
    Set myConn = Nothing
    Set myRecord = Nothing
End Sub
