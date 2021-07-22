VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim oXMLFile As MSXML2.DOMDocument60
Dim xRootNode As MSXML2.IXMLDOMNode
Dim xRootNodeList As IXMLDOMNodeList
Dim xTypeList As IXMLDOMNodeList

                                                                                     
Private Sub ListBox1_Click()
    
    Dim SelectedValue As String
    Dim SelectedItem As Variant
    Dim statusBit As Integer
    
    Sheet1.ListBox2.Clear

    Set oXMLFile = New MSXML2.DOMDocument60
    Set mainWorkBook = ActiveWorkbook
    
    XMLFileName = Sheet1.TextBox1.Text
    
If oXMLFile.Load(XMLFileName) Then
        statusBit = 1
Else
        statusBit = 2
End If

If statusBit = 1 Then
        Sheet1.TextBox2.Text = "Status Good"
End If
        
If statusBit = 2 Then
        Sheet1.TextBox2.Text = "Status Bad (Reload File)"
End If
    
'Sheet1.ListBox2.Clear

If statusBit = 1 Then

    SelectedValue = Sheet1.ListBox1.Value
    Set xTypeList = oXMLFile.DocumentElement.ChildNodes.Item(1).ChildNodes 'Finds all the types of variables in the XML.  Finds Data Types
    
    For Each Item In xTypeList '
     If Item.BaseName = "TypeUserDef" Then
            For Each IECName In Item.Attributes ' This will go through all the items and looks for the node that has the name iecname
                If IECName.BaseName = "iecname" Then
                        If IECName.nodeTypedValue = SelectedValue Then
                           For Each ElementItem In Item.ChildNodes
                                For Each ElementProperty In ElementItem.Attributes
                                    If ElementProperty.BaseName = "iecname" Then
                                        With Sheet1.ListBox2
                                           .AddItem ElementProperty.nodeTypedValue
                                        End With
                                   End If
                                Next
                            Next
                        End If
                End If
            Next
        End If
    Next
End If

       
End Sub


