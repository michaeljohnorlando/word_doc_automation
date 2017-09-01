Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub create_Click()
    Dim name As Range
    Set name = ActiveDocument.Bookmarks("name").Range
    name.Text = Me.TextBox1.Value
    ActiveDocument.Bookmarks.Add "name", name
    
    Dim street As Range
    Set street = ActiveDocument.Bookmarks("street").Range
    street.Text = Me.TextBox2.Value
    ActiveDocument.Bookmarks.Add "street", street
    
    Dim number As Range
    Set number = ActiveDocument.Bookmarks("number").Range
    number.Text = Me.TextBox3.Value
    ActiveDocument.Bookmarks.Add "number", number
    
    ThisDocument.Fields.Update
    Me.Repaint
    change_box.Hide
    ActiveDocument.PrintPreview
End Sub

Private Sub quit_Click()
change_box.Hide
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub
