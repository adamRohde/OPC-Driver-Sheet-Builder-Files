Attribute VB_Name = "CreateOPC"

Dim sDataType As String
Dim vaFiles As String
Dim mainWorkBook As Workbook
Dim OPCBrowsePathCells As Variant
Dim ApplicationVarCell As Variant
Dim ApplicationVarType As Variant
Dim VarDeclaredIn As String
Dim TestVar As Integer
Dim AppVars As Variant
Dim strXML As String
Dim sLoadStatus As String
Dim ListBoxValue As Variant


Dim sTemp As String
Dim sTemp1 As String

'Public oXMLFile As MSXML2.DOMDocument
Public oXMLFile As MSXML2.DOMDocument60
Public xRootNode As MSXML2.IXMLDOMNode
Public xNodeList As IXMLDOMNodeList
Public xDataStructures As IXMLDOMNodeList
Public statusBit As Integer
 
Dim bAllInOne As Boolean
Dim xClass1 As New Class1
Dim xtester As Boolean


'########################### Create OPC Data  ###########################

'Public Sub OptionButton1_Click()

   'If xClass1.OPCFormatConfig(False) = True Then
  '      bAllInOne = xClass1.OPCFormatConfig(False)
 '   End If
    
'End Sub

'Public Sub OptionButton2_Click()
 
   ' If xClass1.OPCFormatConfig(True) = True Then
  '      bAllInOne = xClass1.OPCFormatConfig(True)
 '   End If
    
'End Sub

       
Public Sub CreateOPCData()
      
If Sheet1.OptionButton1.Value = True Then

        Set oXMLFile = New MSXML2.DOMDocument60
        Set mainWorkBook = ActiveWorkbook
            
        XMLFileName = Sheet1.TextBox1.Text
        
        If oXMLFile.Load(XMLFileName) Then
          
            statusBit = 1
            Set xTypeList = oXMLFile.DocumentElement.ChildNodes.Item(1).ChildNodes
            Set xDataStructures = xTypeList
            Set xNodeList = oXMLFile.DocumentElement.ChildNodes.Item(2).ChildNodes.Item(0).ChildNodes
        Else
            statusBit = 2
        End If
        
        If statusBit = 1 Then
                Sheet1.TextBox2.Text = "Status Good"
        End If
            
        If statusBit = 2 Then
                Sheet1.TextBox2.Text = "Status Bad (Reload File)"
        End If
                   
        Sheet = 1
        
        
    If statusBit = 1 Then
        For I = 0 To (xNodeList.Length - 1) 'Indexes through POUs
                Sheet = I + 1
                Col = 1
                Row = 1
                   SheetName = xNodeList.Item(I).Attributes.Item(0).nodeTypedValue
                   Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SheetName 'xNodeList.Item(i).Attributes.Item(0).nodeTypedValue 'Add worksheet with name of POU
                   
                        mainWorkBook.Sheets(SheetName).Cells(1, 1).Value = "Tag Name"
                        mainWorkBook.Sheets(SheetName).Cells(1, 2).Value = "Node ID"
                        mainWorkBook.Sheets(SheetName).Cells(1, 3).Value = "Scan"
                        mainWorkBook.Sheets(SheetName).Cells(1, 4).Value = "Div"
                        mainWorkBook.Sheets(SheetName).Cells(1, 5).Value = "Add"
                        mainWorkBook.Sheets(SheetName).Cells(1, 6).Value = "Browse Path"
                        
                        Col = Col + 1
                     
                For j = 0 To (xNodeList.Item(I).ChildNodes.Length - 1) 'Indexes through variables
                                                       
                    sDataType = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(1).nodeTypedValue 'Looks at the attribute to see what the data type of the PLC variable is
                                   
                    Set ApplicationVarType = mainWorkBook.Sheets(SheetName).Cells(Col, Row + 1)
                    Set ApplicationVarName = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 1)
                    Set ApplicationVarNodeID = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 2)
                    Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                    
                    Dim ReturnType As Variant
                    Dim sStructFound As Variant
                    Dim xGetType As New GetType
                    Set xGetType = New GetType
                    
                    ReturnType = xGetType.SelectDataType(sDataType)
                    
                    sStructFound = ReturnType(2)
                    
                    VarDeclaredIn = xNodeList.Item(I).Attributes.Item(0).NodeValue 'Gets POU Name
                     
    '##################  Structures  ##################
                    
                    If sStructFound = "1" Then  'You got a structure on your hands
                        
                        Dim m As Integer
                        Dim VarType As String
                        Dim VarName As String
                        Dim ElementName As String
                        
                        VarType = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(1).nodeTypedValue
                        VarName = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue
                        
                        For Each Item In xDataStructures
                            If VarType = Item.Attributes.Item(0).NodeValue Then
                                For m = 0 To (Item.ChildNodes.Length - 1)
                                
                                    Set ApplicationVarName = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 1)
                                    Set ApplicationVarNodeID = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 2)
                                    Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                                    
                                    ElementName = Item.ChildNodes.Item(m).Attributes.Item(0).nodeTypedValue
                                    ApplicationVarName.Value = VarName & "." & ElementName
                                    OPCApplicationName = VarName & "/2:" & ElementName
                                    ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value
                                    OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & OPCApplicationName
                                    Col = Col + 1
                                Next
                            End If
                        Next
                           
                    End If
                    
    '##################  TypeSimple  ##################
                    
                   If sStructFound = "0" Then
                        If InStr(sDataType, "T_ARRAY") = 0 Then 'Makes sure it isn't array and if it isn't then we start to populate
                                    Col = Col + 1
                                    ApplicationVarName.Value = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue
                                    ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value 'Gets OPC Browse Path
                                    OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & ApplicationVarName.Value 'Gets OPC Browse Path
                         End If
                   End If
                                        
    '##################  ARRAYS  ##################
                    If InStr(sDataType, "T_ARRAY") = 1 Then ' end if 1
                          ApplicationVarType = "Array"
                            
                                For Each Item In xTypeList 'This will go through the entire typelist looking for "TypeArray" next 1
                                    If Item.BaseName = "TypeArray" Then
                                        a = 0
                                       For Each PropertyListItem In Item.Attributes
                                            a = a + 1
                                            If PropertyListItem.BaseName = "basetype" Then
                                                    Dim sType As String
                                                    sType = Item.Attributes.Item(a - 1).nodeTypedValue
                                                    ReturnType = xGetType.SelectDataType(sType)
                                                    
                                                        If sDataType = Item.Attributes(0).nodeTypedValue Then 'Compares TypeList basename to the variable type if they are arrays end if 4
                                                            iCount = iCount + 1
                                                                                       
                                                                    If iCount > 1 Then
                                                                        j = j + 1
                                                                        iCount = 1
                                                                    End If
                                                        If ReturnType(2) = 1 Then
                                                                        VarType = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(1).nodeTypedValue
                                                                        VarName = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue
                                                                  
                                                                        For Each ItemMatch In xTypeList
                                                                            If sType = ItemMatch.Attributes.Item(0).nodeTypedValue Then
                                                                                    For b = 0 To Item.ChildNodes.Item(0).Attributes.Item(1).nodeTypedValue 'Find max range of array struct
                                                                                        For Each ChildNodeStructElement In ItemMatch.ChildNodes
                                                                                            Dim ElementName2 As String
                                                                                            ElementName2 = ChildNodeStructElement.Attributes.Item(1).nodeTypedValue
                                                                                            Set ApplicationVarName = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 1)
                                                                                            ApplicationVarName.Value = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "[" & b & "]" & "." & ElementName2 'TAG NAME
                                                                                
                                                                                            Set ApplicationVarNodeID = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 2)
                                                                                            ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value '& ElementName2  'NODE ID
                                                                                             
                                                                                            Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                                                                                            OPCBrowsePathVariable = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "&" & "[" & b & "&" & "]"
                                                                                            OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & OPCBrowsePathVariable & "/2:" & ElementName2 'BROWSED PATH
                                                                                            Col = Col + 1
                                                                                        Next
                                                                                    Next
                                                                            End If
                                                                        Next
                                                        Else
                                                        
                                                                For C = 0 To Item.ChildNodes.Item(0).Attributes.Item(1).nodeTypedValue
                                                                
                                                                        Set ApplicationVarName = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 1)
                                                                        ApplicationVarName.Value = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "[" & C & "]"
                                                                
                                                                        Set ApplicationVarNodeID = mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Cells(Col, 2)
                                                                        ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value 'Gets OPC Browse Path
                                                                
                                                                        Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                                                                        OPCBrowsePathVariable = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "&" & "[" & C & "&" & "]"
                                                                        OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & OPCBrowsePathVariable 'Gets OPC Browse Path
                                                                        Col = Col + 1
                                                                        a = 0
                                                                Next
                                                                
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                                iCount = 0
                    Next
                    mainWorkBook.Sheets(xNodeList.Item(I).Attributes.Item(0).nodeTypedValue).Columns("A:F").AutoFit
              Next
        End If
        
    Else
    
        '################################################################################################################################################################################

        Set oXMLFile = New MSXML2.DOMDocument60
        Set mainWorkBook = ActiveWorkbook
            
        XMLFileName = Sheet1.TextBox1.Text
        
        If oXMLFile.Load(XMLFileName) Then
          
            statusBit = 1
            Set xTypeList = oXMLFile.DocumentElement.ChildNodes.Item(1).ChildNodes
            Set xDataStructures = xTypeList
            Set xNodeList = oXMLFile.DocumentElement.ChildNodes.Item(2).ChildNodes.Item(0).ChildNodes
        Else
            statusBit = 2
        End If
        
        If statusBit = 1 Then
                Sheet1.TextBox2.Text = "Status Good"
        End If
            
        If statusBit = 2 Then
                Sheet1.TextBox2.Text = "Status Bad (Reload File)"
        End If
                   
        Sheet = 1
        
        
    If statusBit = 1 Then
    
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "New" 'xNodeList.Item(i).Attributes.Item(0).nodeTypedValue 'Add worksheet with name of POU
        SheetName = "New"
        For I = 0 To (xNodeList.Length - 1) 'Indexes through POUs
                Sheet = 1
                Col = 1
                Row = 1
                   'SheetName = xNodeList.Item(I).Attributes.Item(0).nodeTypedValue
                   'Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SheetName 'xNodeList.Item(i).Attributes.Item(0).nodeTypedValue 'Add worksheet with name of POU
                   
                        mainWorkBook.Sheets(SheetName).Cells(1, 1).Value = "Tag Name"
                        mainWorkBook.Sheets(SheetName).Cells(1, 2).Value = "Node ID"
                        mainWorkBook.Sheets(SheetName).Cells(1, 3).Value = "Scan"
                        mainWorkBook.Sheets(SheetName).Cells(1, 4).Value = "Div"
                        mainWorkBook.Sheets(SheetName).Cells(1, 5).Value = "Add"
                        mainWorkBook.Sheets(SheetName).Cells(1, 6).Value = "Browse Path"
                        
                        Col = Col + 1
                     
                For j = 0 To (xNodeList.Item(I).ChildNodes.Length - 1) 'Indexes through variables
                                                       
                    sDataType = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(1).nodeTypedValue 'Looks at the attribute to see what the data type of the PLC variable is
                                   
                    Set ApplicationVarType = mainWorkBook.Sheets(SheetName).Cells(Col, Row + 1)
                    Set ApplicationVarName = mainWorkBook.Sheets(SheetName).Cells(Col, 1)
                    Set ApplicationVarNodeID = mainWorkBook.Sheets(SheetName).Cells(Col, 2)
                    Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                    
                    Set xGetType = New GetType
                    
                    ReturnType = xGetType.SelectDataType(sDataType)
                    
                    sStructFound = ReturnType(2)
                    
                    VarDeclaredIn = xNodeList.Item(I).Attributes.Item(0).NodeValue 'Gets POU Name
                     
    '##################  Structures  ##################
                    
                    If sStructFound = "1" Then  'You got a structure on your hands
                        
                        VarType = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(1).nodeTypedValue
                        VarName = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue
                        
                        For Each Item In xDataStructures
                            If VarType = Item.Attributes.Item(0).NodeValue Then
                                For m = 0 To (Item.ChildNodes.Length - 1)
                                
                                    Set ApplicationVarName = mainWorkBook.Sheets(SheetName).Cells(Col, 1)
                                    Set ApplicationVarNodeID = mainWorkBook.Sheets(SheetName).Cells(Col, 2)
                                    Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                                    
                                    ElementName = Item.ChildNodes.Item(m).Attributes.Item(0).nodeTypedValue
                                    ApplicationVarName.Value = VarName & "." & ElementName
                                    OPCApplicationName = VarName & "/2:" & ElementName
                                    ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value
                                    OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & OPCApplicationName
                                    Col = Col + 1
                                Next
                            End If
                        Next
                           
                    End If
                    
    '##################  TypeSimple  ##################
                    
                   If sStructFound = "0" Then
                        If InStr(sDataType, "T_ARRAY") = 0 Then 'Makes sure it isn't array and if it isn't then we start to populate
                                    Col = Col + 1
                                    ApplicationVarName.Value = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue
                                    ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value 'Gets OPC Browse Path
                                    OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & ApplicationVarName.Value 'Gets OPC Browse Path
                         End If
                   End If
                                        
    '##################  ARRAYS  ##################
                    If InStr(sDataType, "T_ARRAY") = 1 Then ' end if 1
                          ApplicationVarType = "Array"
                            
                                For Each Item In xTypeList 'This will go through the entire typelist looking for "TypeArray" next 1
                                    If Item.BaseName = "TypeArray" Then
                                        a = 0
                                       For Each PropertyListItem In Item.Attributes
                                            a = a + 1
                                            If PropertyListItem.BaseName = "basetype" Then
                                                    sType = Item.Attributes.Item(a - 1).nodeTypedValue
                                                    ReturnType = xGetType.SelectDataType(sType)
                                                    
                                                        If sDataType = Item.Attributes(0).nodeTypedValue Then 'Compares TypeList basename to the variable type if they are arrays end if 4
                                                            iCount = iCount + 1
                                                                                       
                                                                    If iCount > 1 Then
                                                                        j = j + 1
                                                                        iCount = 1
                                                                    End If
                                                        If ReturnType(2) = 1 Then
                                                                        VarType = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(1).nodeTypedValue
                                                                        VarName = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue
                                                                  
                                                                        For Each ItemMatch In xTypeList
                                                                            If sType = ItemMatch.Attributes.Item(0).nodeTypedValue Then
                                                                                    For b = 0 To Item.ChildNodes.Item(0).Attributes.Item(1).nodeTypedValue 'Find max range of array struct
                                                                                        For Each ChildNodeStructElement In ItemMatch.ChildNodes
                                                                                            ElementName2 = ChildNodeStructElement.Attributes.Item(1).nodeTypedValue
                                                                                            Set ApplicationVarName = mainWorkBook.Sheets(SheetName).Cells(Col, 1)
                                                                                            ApplicationVarName.Value = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "[" & b & "]" & "." & ElementName2 'TAG NAME
                                                                                
                                                                                            Set ApplicationVarNodeID = mainWorkBook.Sheets(SheetName).Cells(Col, 2)
                                                                                            ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value '& ElementName2  'NODE ID
                                                                                             
                                                                                            Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                                                                                            OPCBrowsePathVariable = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "&" & "[" & b & "&" & "]"
                                                                                            OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & OPCBrowsePathVariable & "/2:" & ElementName2 'BROWSED PATH
                                                                                            Col = Col + 1
                                                                                        Next
                                                                                    Next
                                                                            End If
                                                                        Next
                                                        Else
                                                        
                                                                For C = 0 To Item.ChildNodes.Item(0).Attributes.Item(1).nodeTypedValue
                                                                
                                                                        Set ApplicationVarName = mainWorkBook.Sheets(SheetName).Cells(Col, 1)
                                                                        ApplicationVarName.Value = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "[" & C & "]"
                                                                
                                                                        Set ApplicationVarNodeID = mainWorkBook.Sheets(SheetName).Cells(Col, 2)
                                                                        ApplicationVarNodeID.Value = "ns=2;s=Application." & VarDeclaredIn & "." & ApplicationVarName.Value 'Gets OPC Browse Path
                                                                
                                                                        Set OPCBrowsePathCells = mainWorkBook.Sheets(SheetName).Cells(Col, 6)
                                                                        OPCBrowsePathVariable = xNodeList.Item(I).ChildNodes.Item(j).Attributes.Item(0).nodeTypedValue & "&" & "[" & C & "&" & "]"
                                                                        OPCBrowsePathCells.Value = "/0:Objects/2:Logic/2:Application/2:" & VarDeclaredIn & "/2:" & OPCBrowsePathVariable 'Gets OPC Browse Path
                                                                        Col = Col + 1
                                                                        a = 0
                                                                Next
                                                                
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            iCount = 0
                    Next
                    mainWorkBook.Sheets(SheetName).Columns("A:F").AutoFit
              Next
        End If
        
        '############################################################################################
    End If
End Sub

