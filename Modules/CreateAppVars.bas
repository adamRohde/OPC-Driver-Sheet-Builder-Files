Attribute VB_Name = "CreateAppVars"

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
Dim sDataType As String
Dim ListBoxValue As Variant

Dim oXMLFile As MSXML2.DOMDocument60
Dim xNodeList As IXMLDOMNodeList
Dim xTypeList As IXMLDOMNodeList
Public statusBit As Integer


'########################### Gets XML File ###########################
Sub Button1_Click()

    vaFiles = Application.GetOpenFilename
    Sheet1.TextBox1.Text = vaFiles

    Set oXMLFile = New MSXML2.DOMDocument60
    Set mainWorkBook = ActiveWorkbook
    
    XMLFileName = Sheet1.TextBox1.Text
   
    If oXMLFile.Load(XMLFileName) Then
        statusBit = 1
        Sheet1.ListBox1.Clear
        Set xTypeList = oXMLFile.DocumentElement.ChildNodes.Item(1).ChildNodes
        Set xNodeList = oXMLFile.DocumentElement.ChildNodes.Item(2).ChildNodes
    Else
        statusBit = 2
    End If
    
        If statusBit = 1 Then
            For Each Item In xTypeList
              If Item.BaseName = "TypeUserDef" Then
                For Each IECName In Item.Attributes
                    If IECName.BaseName = "iecname" Then
                        With Sheet1.ListBox1
                            .AddItem IECName.nodeTypedValue
                        End With
                    End If
                Next
              End If
            Next
        End If
        
        If statusBit = 1 Then
            Sheet1.TextBox2.Text = "Status Good"
        End If
        
        If statusBit = 2 Then
            Sheet1.TextBox2.Text = "Status Bad (Reload File)"
        End If
        

End Sub
 
'########################### Create Application Vars ###########################
Public Sub CreateApplicationVars()

If statusBit = 0 Then  '9/17/2020 this is right.  For some reason you had to change to look for 0 not 1.  Status bit must not be working right

    XMLFileName = Sheet1.TextBox1.Text
    
    '### Sets ###
    Set oXMLFile = New MSXML2.DOMDocument60
    oXMLFile.Load (XMLFileName)
    Set mainWorkBook = ActiveWorkbook
    
    Set xTypeList = oXMLFile.DocumentElement.ChildNodes.Item(1).ChildNodes
    Set xNodeList = oXMLFile.DocumentElement.ChildNodes.Item(2).ChildNodes.Item(0).ChildNodes
     
    'This looks to see if the "Application Vars" sheets has already been created or not
    Dim ws As Worksheet
    Dim iCounter As Integer
    
    For Each ws In ThisWorkbook.Sheets
        If InStr(1, ws.Name, "Application Vars") Then
            iCounter = iCounter + 1
        Else
            iCounter = iCounter
        End If
    Next
    
    If iCounter = 0 Then 'If button was pressed the sheets doesn't already exist (knows via counter) then add the sheet "Application Vars"
       Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Application Vars"
    Else
       Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Application Vars" & iCounter
    End If
             
   'Gets the variables from the PLC and adds them to the main sheet.  Seperates the POUs/GVLs via columns
    Sheets("Application Vars").Cells.Clear
   
    mainWorkBook.Sheets("Application Vars").Cells(1, 1).Value = "Name"
    mainWorkBook.Sheets("Application Vars").Cells(1, 2).Value = "Array"
    mainWorkBook.Sheets("Application Vars").Cells(1, 3).Value = "Type"
     
    Row = 1
    Col = 3
    
            For Each Item In xNodeList
            
                    Set ApplicationVarPOU = mainWorkBook.Sheets("Application Vars").Cells(Col, 4)
                    ApplicationVarPOU.Value = Item.Attributes.Item(0).nodeTypedValue  'This gets "POU" name (e.g. GVL, PLCProg, myPOU)
                                     
                      For Each VarObject In Item.ChildNodes 'Indexes Variables in individual POUs
                      
                            Set ApplicationVarName = mainWorkBook.Sheets("Application Vars").Cells(Col, 1)
                            Set ApplicationVarType = mainWorkBook.Sheets("Application Vars").Cells(Col, 3)
                            Set ApplicationVarSize = mainWorkBook.Sheets("Application Vars").Cells(Col, 2)
                            
                            
                            ApplicationVarName.Value = VarObject.Attributes.Item(0).nodeTypedValue
                            sDataType = VarObject.Attributes.Item(1).nodeTypedValue 'Gets the variable type (e.g. BOOL)
                
                            Col = Col + 1
                                                     
                            Dim xGetType As New GetType
                            Set xGetType = New GetType

                            ApplicationVarType.Value = xGetType.SelectDataType(sDataType) 'Calls function that gets the data types
                                                                       
                            '##################  ARRAYS  ##################
                
                            If InStr(sDataType, "T_ARRAY") > 0 Then
                                For Each TypeObject In xTypeList 'This will go through the entire typelist looking for "TypeArray"
                                        If TypeObject.nodeName = "TypeArray" Then 'If type if found to be of type array then we go inside the loop to find the specific array that matches the current var type
                                                    If TypeObject.Attributes.Item(0).nodeTypedValue = sDataType Then
                                                        ApplicationVarSize.Value = TypeObject.ChildNodes.Item(0).Attributes.Item(1).nodeTypedValue 'Grabs the "maxrange" item and displays
                                                        ApplicationVarType.Value = xGetType.SelectDataType(TypeObject.Attributes.Item(5).nodeTypedValue)
                                                    End If
                                        End If
                                  Next
                            End If
                      
                            mainWorkBook.Sheets("Application Vars").Columns("A:C").AutoFit
                 
                      Next
            Next
     End If
End Sub


