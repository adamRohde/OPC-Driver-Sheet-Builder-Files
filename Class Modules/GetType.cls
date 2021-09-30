VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function SelectDataType(DataType As String)
                   
                    Dim Element(1 To 2) As String
                    
                    If InStr(DataType, "T_ARRAY") = 1 Then 'Does serve much of a
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_BOOL" Then
                        Element(1) = "Boolean"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    '##################  BYTE  ##################
                     If DataType = "T_BYTE" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    '##################  WORD  ##################
                     If DataType = "T_WORD" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_DWORD" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_LWORD" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    '##################  INTS  ##################
                     If DataType = "T_SINT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_INT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                     If DataType = "T_DINT" Then
                       Element(1) = "Integer"
                       Element(2) = "0"
                       SelectDataType = Element
                       Exit Function
                    End If
                    
                     If DataType = "T_LINT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                     If DataType = "T_USINT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                     If DataType = "T_UINT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                     If DataType = "T_UDINT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                     If DataType = "T_ULINT" Then
                        Element(1) = "Integer"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                     
                    '##################  REALS  ##################
                    If DataType = "T_REAL" Then
                        Element(1) = "Real"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_LREAL" Then
                        Element(1) = "Real"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    '##################  STRINGS  ##################
                    If InStr(DataType, "T_STRING") Then
                        Element(1) = "String"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_WSTRING" Then
                        Element(1) = "String"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    If DataType = "T_TIME" Then
                        Element(1) = "String"
                        Element(2) = "0"
                        SelectDataType = Element
                        Exit Function
                    End If
                    
                    '##################  User Defined Data Structures  ##################
                    Dim I As Integer
                    For I = 0 To (Sheet1.ListBox1.ListCount - 1)
                        If DataType = "T_" & Sheet1.ListBox1.List(I) Then
                            Element(1) = Sheet1.ListBox1.List(I)
                            Element(2) = "1"
                            SelectDataType = Element
                            Exit Function
                        Else
                            Element(2) = "0"
                            SelectDataType = Element
                        End If
                    Next
                                        
End Function


