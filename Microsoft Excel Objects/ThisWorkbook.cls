VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_Open()
     Sheet1.TextBox1.Text = "Click Browse and Select .xml file"
     Sheet1.ListBox1.Clear
     Sheet1.ListBox2.Clear
     Sheet1.TextBox2.Text = "no file selected"
End Sub
