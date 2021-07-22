VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function OPCFormatConfig(Selection As Boolean)

    If Selection = True Then
        OPCFormatConfig = True
    End If
    
    If Selection = False Then
        OPCFormatConfig = False
    End If
    
End Function
