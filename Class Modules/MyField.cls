VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by Mehmet Gürevin :)

Option Explicit

Private local_FieldType                              As VbVarType
Private local_FieldLength                            As Long
Private local_FieldName                              As String
Private local_Value                                  As Variant
Private local_Tag                                    As String
Private local_Index                                  As Long
Private local_MySQLTypeNumber                        As Long
Private local_RealLength                             As Long

Public Property Get FieldType() As VbVarType
    FieldType = local_FieldType
End Property

Public Property Let FieldType(ByVal Value As VbVarType)
    local_FieldType = Value
End Property

Public Property Get FieldLength() As Long
    FieldLength = local_FieldLength
End Property

Public Property Let FieldLength(ByVal Value As Long)
    local_FieldLength = Value
End Property

Public Property Get FieldName() As String
    FieldName = local_FieldName
End Property

Public Property Let FieldName(ByVal Value As String)
    local_FieldName = Value
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = local_Value
End Property

Public Property Let Value(ByVal Value As Variant)
    local_Value = Value
End Property

Public Property Get Tag() As String
    Tag = local_Tag
End Property

Public Property Let Tag(ByVal Value As String)
    local_Tag = Value
End Property

Public Property Get Index() As Long
    Index = local_Index
End Property

Public Property Let Index(ByVal Value As Long)
    local_Index = Value
End Property

Public Property Get MySQLTypeNumber() As Long
    MySQLTypeNumber = local_MySQLTypeNumber
End Property

Public Property Let MySQLTypeNumber(ByVal Value As Long)
    local_MySQLTypeNumber = Value
End Property

Public Property Get RealLength() As Long
    RealLength = local_RealLength
End Property

Public Property Let RealLength(ByVal Value As Long)
    local_RealLength = Value
End Property
