Option Compare Database
Private Sub cmdAdd_Click()

'when we click Add there are 2 options:
'1. for insert
'2. for update
 counter = Int(DCount("*", "jscbb_dir"))
If Me.Textidd.Tag & "" = "" Then
ssql = "INSERT INTO jscbb_dir(ID,LastName,FirstName,CardNum, PrimA, Area,LabNum,OfficeNum, OfficeAccess, OfficePhone, Notes, BuffOne, Email,LabPhone, EEID, Status, Mailbox, ActivityLevel)" & _
" VALUES(" & Me.Textidd & ",'" & _
Me.TextLast & "','" & Me.TextFirst & "','" & Me.Textcardnum & "','" & Me.Textprima & _
"','" & Me.Textarea & "','" & Me.Textlabnum & _
"','" & Me.Textofficenum & "','" & Me.Textofficeaccess & "','" & Me.Textofficephone & "','" & Me.Textnotes & "','" & Me.Textbuffone & _
"','" & Me.Textemail & "','" & Me.Textlabphone & "','" & Me.Texteeid & "','" & Me.Textstatus & "','" & Me.Textmailbox & "','" & Me.Textactivitylevel & "')"
If (Int(Me.Textidd.Value) > Int(counter + 1)) Then

        Set Db = CurrentDb()
        Db.Execute ssql
            MsgBox "New Record Added! " & Me.TextFirst & " " & Me.TextLast & " " & "At ID: " & Me.Textidd.Value
        Else
            MsgBox "The record you are trying to add " & Me.Textidd & " already exists"

       End If



Else
CurrentDb.Execute "UPDATE jscbb_dir " & " SET ID=" & Me.Textidd & _
    ", LastName='" & Me.TextLast & "'" & _
    ", FirstName='" & Me.TextFirst & "'" & _
    ", CardNum='" & Me.Textcardnum & "'" & _
    ", PrimA='" & Me.Textprima & "'" & _
    ", Area='" & Me.Textarea & "'" & _
    ", LabNum='" & Me.Textlabnum & "'" & _
    ", OfficeNum='" & Me.Textofficenum & "'" & _
    ", OfficeAccess='" & Me.Textofficeaccess & "'" & _
    ", OfficePhone='" & Me.Textofficephone & "'" & _
    ", Notes='" & Me.Textnotes & "'" & _
    ", BuffOne='" & Me.Textbuffone & "'" & _
    ", Email='" & Me.Textemail & "'" & _
    ", LabPhone='" & Me.Textlabphone & "'" & _
    ", EEID='" & Me.Texteeid & "'" & _
    ", Status ='" & Me.Textstatus & "'" & _
    ", Mailbox ='" & Me.Textmailbox & "'" & _
    ", ActivityLevel ='" & Me.Textactivitylevel & "'" & _
    "WHERE ID=" & Me.Textidd.Tag
End If

'clear form
cmdClear_Click
'refresh data is list on focus
jscbb_dirsub.Form.Requery


End Sub

Private Sub cmdClear_Click()
Me.Textidd = ""
Me.TextLast = ""
Me.TextFirst = ""
Me.Textcardnum = ""
Me.Textprima = ""
Me.Textarea = ""
Me.Textlabnum = ""
Me.Textofficenum = ""
Me.Textofficeaccess = ""
Me.Textofficephone = ""
Me.Textnotes = ""
Me.Textbuffone = ""
Me.Textemail = ""
Me.Textlabphone = ""
Me.Texteeid = ""
Me.Textstatus = ""
Me.Textmailbox = ""
Me.Textactivitylevel = ""
'focus on text id box
Me.Textidd.SetFocus
'set edit button to enabled
Me.cmdEdit.Enabled = True
'change add to Add
Me.cmdAdd.Caption = "Add"
'clear tag on Textidd for reset new
Me.Textidd.Tag = ""
End Sub

Private Sub cmdClose_Click()
DoCmd.Close
End Sub






Private Sub cmdDelete_Click()

'delete record
'check existing seleced record
If Not (Me.jscbb_dirsub.Form.Recordset.EOF And Me.jscbb_dirsub.Form.Recordset.BOF) Then
    'confirm delete
    If MsgBox("Are you sure you want to delete this record? " & Me.TextFirst.Value & " " & Me.TextLast.Value, vbYesNo) = vbYes Then
        'delete now
        CurrentDb.Execute "DELETE FROM jscbb_dir " & _
            " WHERE ID=" & Me.jscbb_dirsub.Form.Recordset.Fields("ID")
    'refresh data
    Me.jscbb_dirsub.Form.Requery
End If
End If

End Sub

Private Sub cmdEdit_Click()
'check whether there exists data in set
If Not (Me.jscbb_dirsub.Form.Recordset.EOF And Me.jscbb_dirsub.Form.Recordset.BOF) Then
'get data from record set to box
    With Me.jscbb_dirsub.Form.Recordset
        Me.Textidd = .Fields("ID")
        Me.TextLast = .Fields("Lastname")
        Me.TextFirst = .Fields("FirstName")
        Me.Textcardnum = .Fields("CardNum")
        Me.Textprima = .Fields("PrimA")
        Me.Textarea = .Fields("Area")
        Me.Textlabnum = .Fields("LabNum")
        Me.Textofficenum = .Fields("OfficeNum")
        Me.Textofficeaccess = .Fields("OfficeAccess")
        Me.Textofficephone = .Fields("OfficePhone")
        Me.Textnotes = .Fields("Notes")
        Me.Textbuffone = .Fields("BuffOne")
        Me.Textemail = .Fields("Email")
        Me.Textlabphone = .Fields("LabPhone")
        Me.Texteeid = .Fields("EEID")
        Me.Textstatus = .Fields("Status")
        Me.Textmailbox = .Fields("Mailbox")
        Me.Textactivitylevel = .Fields("ActivityLevel")
        'store id of Textidd
        Me.Textidd.Tag = ("ID")
        'change add button to update
        Me.cmdAdd.Caption = "Update"
        'disable edit button
        Me.cmdEdit.Enabled = False
       
    End With
End If
End Sub
