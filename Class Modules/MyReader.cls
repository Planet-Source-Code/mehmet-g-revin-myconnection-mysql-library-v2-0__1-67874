VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by Mehmet Gürevin :)

Option Explicit

Private Const SIZE_OF_CHAR = 4

Private Declare Function mysql_store_result Lib "libmysql.dll" (ByVal hMysql As Long) As Long
Private Declare Function mysql_free_result Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long
Private Declare Function mysql_fetch_row Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                                              'Return the row struct pointer
Private Declare Function mysql_field_count Lib "libmysql.dll" (ByVal hMysql As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function mysql_fetch_field Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                                            'Return the field struct pointer
Private Declare Function mysql_fetch_lengths Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                              'returns * unsigned long
Private Declare Function mysql_num_rows Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                                               'Return the 8 bit lenght pointer, myulonglong

Private lResult                         As Long
Private mCnnPtr                         As Long
Private lRes                            As Long
Private bClosed                         As Boolean
Private mCNN                            As MyConnection

Private Values()                        As MyField

Public Property Set CnnObject(ByRef lObject As MyConnection)
    Set mCNN = lObject
End Property

Public Property Get ResultPointer() As Long
    ResultPointer = lResult
End Property

Public Property Let ResultPointer(ByVal Value As Long)
    lResult = Value
    lRes = mysql_store_result(mCnnPtr)
    If lRes Then
        bClosed = False
        
        Dim i As Long
        Dim fType                           As Long
        Dim lRet                            As Long
        Dim fLen                            As Long
        Dim fName                           As Long
        Dim fCount                          As Long
        
        fCount = mysql_field_count(mCnnPtr)
        
        ReDim Values(fCount - 1) As MyField
        
        For i = 0 To fCount - 1
            Set Values(i) = New MyField
        Next i
        
        For i = 1 To fCount
            lRet = mysql_fetch_field(lRes)
            If lRet Then
                CopyMemory fName, ByVal lRet, 4
                CopyMemory fLen, ByVal lRet + 28, 4
                CopyMemory fType, ByVal lRet + 76, 4
            End If
            
            Select Case fType
                Case 16, 1, 2, 9, 3, 8 '16: BIT, 1:TINYINT, 2:SMALLINT, 9:MEDIUMINT, 3:INT, 8:BIGINT
                    'Long
                    Values(i - 1).FieldType = vbLong
                Case 4   'FLOAT
                    'Currency
                    Values(i - 1).FieldType = vbCurrency
                Case 5   'DOUBLE
                    'Double
                    Values(i - 1).FieldType = vbDouble
                Case 246, 254, 253 '[246:DECIMAL][245:CHAR,BINARY,ENUM,SET][253:VARCHAR , VARBINARY]
                    'String
                    Values(i - 1).FieldType = vbString
                    If fType = 253 Then
                        Values(i - 1).Tag = "Trim"
                    End If
                Case 10, 12, 7, 11 '10:DATE, 12:DATETIME, 7:TIMSTAMP, 11:TIME
                    'Date
                    Values(i - 1).FieldType = vbDate
                    If fType = 12 Or fType = 7 Then
                        Values(i - 1).Tag = "Date + Time"
                    ElseIf fType = 11 Then
                        Values(i - 1).Tag = "Time"
                    ElseIf fType = 10 Then
                        Values(i - 1).Tag = "Date"
                    End If
                Case 13  'YEAR
                    'Integer
                    Values(i - 1).FieldType = vbInteger
                Case 252 'TINYTEXT , TEXT, MEDIUMTEXT, LONGTEXT, TINYBLOB, BLOB, MEDIUMBLOB, LONGBLOB
                    'Byte Array
                    Values(i - 1).FieldType = vbArray
                Case 255 'GEOMETRY , POINT, LINESTRING, POLYGON, MULTIPOINT, MULTILINESTRING, MULTIPOLYGON, GEOMETRYCOLLECTION
                    'Tanýmsýz
                    Values(i - 1).FieldType = vbNull
            End Select
            
            Values(i - 1).FieldName = Ptr2Str(fName)
            Values(i - 1).FieldLength = fLen
            Values(i - 1).MySQLTypeNumber = fType
            Values(i - 1).Index = i - 1
        Next i
    Else
        Err.Raise vbObjectError, "MyReader:ResultPointer[" & CStr(Value) & "]", "Kaynak:[mysql_store_result]" & vbCrLf & "Bilinmeyen bir hata oluþtu."
    End If
End Property

Public Property Get ConnectionPointer() As Long
    ConnectionPointer = mCnnPtr
End Property

Public Property Let ConnectionPointer(ByVal Value As Long)
    mCnnPtr = Value
End Property

Public Property Get RowCount() As Long
    If bClosed Then
        Err.Raise vbObjectError, "MyReader:Read()", "Nesne kapalý."
        Exit Property
    End If
    
    RowCount = mysql_num_rows(lRes)
End Property
    
Public Property Get GetValue(ByVal Index As Variant) As MyField
    If bClosed Then
        Err.Raise vbObjectError, "MyReader:Read()", "Nesne kapalý."
        Exit Property
    End If
    
    Dim lValue                  As Long
    Dim sValue                  As String
    Dim i                       As Long
    Dim vFlag                   As Boolean
       
    Select Case VarType(Index)
        Case vbSingle, vbByte, vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong
            Set GetValue = Values(Index)
        Case vbString
            For i = 0 To UBound(Values)
                If Values(i).FieldName = Trim(Index) Then
                    GetValue = Values(i)
                    vFlag = True
                End If
            Next i
            If Not vFlag Then
                Err.Raise vbObjectError, "MyReader:GetValue", "[" & Index & "] Alaný Tabloda Bulunamadý..."
            End If
    End Select
End Property

Public Function Read() As Boolean
    Dim Row                             As Long
    Dim i                               As Long
    Dim mStr                            As String
    
    If bClosed Then
        Err.Raise vbObjectError, "MyReader:Read()", "Nesne kapalý."
        Exit Function
    End If
    
    Row = mysql_fetch_row(lRes)
    
    If Row Then
        Read = True
        Dim ColCount                As Long
        Dim FieldNames()            As Long
        
        ColCount = mysql_field_count(mCnnPtr)
        ReDim FieldNames(1 To ColCount) As Long
        
        Dim FieldsLenghts()     As Long
        ReDim FieldsLenghts(ColCount) As Long
        
        CopyMemory FieldsLenghts(0), ByVal mysql_fetch_lengths(lRes), (ColCount * SIZE_OF_CHAR)
        
        For i = 1 To ColCount
            CopyMemory FieldNames(1), ByVal Row, SIZE_OF_CHAR * ColCount
            
            Values(i - 1).RealLength = FieldsLenghts(i - 1)
            
            Select Case Values(i - 1).FieldType
                Case VbVarType.vbLong
                    On Local Error Resume Next
                        Values(i - 1).Value = CLng(Val(Trim(Ptr2Str(FieldNames(i) & " "))))
                    If Err.Number = 6 Then
                        Values(i - 1).Value = Ptr2Str(FieldNames(i) & " ")
                    End If
                    On Local Error GoTo 0
                Case VbVarType.vbCurrency
                    Values(i - 1).Value = CCur(IIf(Trim(Ptr2Str(FieldNames(i) & " ")) = "", "0" & Mid(CStr((3 / 2)), 2, 1) & "00", Trim(Ptr2Str(FieldNames(i) & " "))))
                Case VbVarType.vbDouble
                    Values(i - 1).Value = CDbl(Val(Trim(Ptr2Str(FieldNames(i) & " "))))
                Case VbVarType.vbString
                    If Values(i - 1).Tag = "Trim" Then
                        Values(i - 1).Value = Ptr2Str(FieldNames(i))
                    Else
                        Values(i - 1).Value = Trim(Ptr2Str(FieldNames(i) & " "))
                    End If
                Case VbVarType.vbInteger
                    Values(i - 1).Value = CInt(Val(Trim(Ptr2Str(FieldNames(i) & " "))))
                Case VbVarType.vbArray
                    Dim DataLength          As Long
                    Dim Data()              As Byte

                    ReDim Data(Values(i - 1).RealLength) As Byte
                    CopyMemory Data(0), ByVal FieldNames(i), Values(i - 1).RealLength

                    Values(i - 1).Value = Data
                    Erase Data
                Case VbVarType.vbDate
                    If Trim(Ptr2Str(FieldNames(i) & " ")) = "" Then
                        Values(i - 1).Value = Null
                    Else
                        mStr = Trim(Ptr2Str(FieldNames(i) & " "))
                        If Values(i - 1).Tag = "Date + Time" Then
                            mStr = Split(mStr, " ")(0)
                            Values(i - 1).Value = CDate(CStr(DateSerial(CInt(Left(mStr, 4)), CInt(Mid(mStr, 6, 2)), CInt(Right(mStr, 2)))) & " " & Split(Trim(Ptr2Str(FieldNames(i) & " ")), " ")(1))
                        ElseIf Values(i - 1).Tag = "Date" Then
                            Values(i - 1).Value = DateSerial(CInt(Left(mStr, 4)), CInt(Mid(mStr, 6, 2)), CInt(Right(mStr, 2)))
                        ElseIf Values(i - 1).Tag = "Time" Then
                            Values(i - 1).Value = CDate(mStr)
                        End If
                    End If
                Case VbVarType.vbNull
                    DoEvents
            End Select
        Next i
    Else
        Call mysql_free_result(lRes)
        Call mCNN.ExecuteReader("#closereader#")
        bClosed = True
        Read = False
    End If
End Function

Public Sub CloseReader()
    If Not bClosed Then
        Call mysql_free_result(lRes)
        bClosed = True
        Call mCNN.ExecuteReader("#closereader#")
    End If
End Sub

Private Sub Class_Initialize()
    bClosed = True
End Sub

Private Function Ptr2Str(ByVal lPtr As Long) As String
    On Local Error Resume Next

    Dim lTmp As Long

    If lPtr = 0 Then Exit Function

    Dim bTmp As Byte
    Dim aBytes() As Byte
    Dim lChars As Long

    lChars = lstrlen(lPtr)
    If lChars = 0 Then Exit Function

    ReDim aBytes(1 To lChars) As Byte
    aBytes = String(lChars, " ")
    CopyMemory aBytes(1), ByVal (lPtr), lChars
    Ptr2Str = StrConv(aBytes, vbUnicode)

    lTmp = InStr(Ptr2Str, vbNullChar)
    If lTmp > 0 Then
        Ptr2Str = Trim(Left$(Ptr2Str, lTmp - 1))
    Else
        Ptr2Str = Ptr2Str
    End If

    Erase aBytes
End Function

Private Sub Class_Terminate()
    If Not bClosed Then
        Call mysql_free_result(lRes)
        Call mCNN.ExecuteReader("#closereader#")
    End If
End Sub

