VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written by Mehmet Gürevin :)

Option Explicit

Public Enum BATHPROGRESSTYPE
    MEMORY_PROGRESS
    FILE_PROGRESS
    RESOURCE_PROGRESS
End Enum

Public Enum MYSTATE
    MY_CONNECTED
    MY_NOT_CONNECTED
End Enum

Private Const SIZE_OF_CHAR = 4

Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare Function CreateReadFileHandle Lib "kernel32.dll" Alias "CreateFileA" ( _
                         ByVal lpFileName As String, _
                         Optional ByVal dwDesiredAccess As Long = &H80000000, _
                         Optional ByVal dwShareMode As Long = &H1, _
                         Optional ByVal lpSecurityAttributes As Long = 0&, _
                         Optional ByVal dwCreationDisposition As Long = 3, _
                         Optional ByVal dwFlagsAndAttributes As Long = 0&, _
                         Optional ByVal hTemplateFile As Long = 0&) As Long
Private Declare Function CreateWriteFileHandle Lib "kernel32" Alias "CreateFileA" ( _
                         ByVal lpFileName As String, _
                         Optional ByVal dwDesiredAccess As Long = &H2, _
                         Optional ByVal dwShareMode As Long = &H2, _
                         Optional ByVal lpSecurityAttributes As Long = 0, _
                         Optional ByVal dwCreationDisposition As Long = 2, _
                         Optional ByVal dwFlagsAndAttributes As Long = 0, _
                         Optional ByVal hTemplateFile As Long = 0) As Long
Private Declare Function APIWriteFile Lib "kernel32" Alias "WriteFile" ( _
                         ByVal hFile As Long, _
                         ByVal lStringPointer As Long, _
                         ByVal lStringLength As Long, _
                         ByRef ReturnToWriteNumberOfBytes As Long, _
                         Optional ByVal DO_NOT_USE As Long = 0) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
                         ByRef lpBuffer As Any, _
                         ByVal nNumberOfBytesToRead As Long, _
                         ByRef lpNumberOfBytesRead As Long, _
                         ByVal lpOverlapped As Any) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Function mysql_init Lib "libmysql.dll" (ByVal hMysql As Long) As Long

Private Declare Function mysql_real_connect Lib "libmysql.dll" (ByVal hMysql As Long, ByVal Host As Long, ByVal User As Long, ByVal Passwd As Long, ByVal dB As Long, ByVal Port As Long, ByVal Unix_Socket As Long, ByVal clientflag As Long) As Long
Private Declare Function mysql_close Lib "libmysql.dll" (ByVal hMysql As Long) As Long

Private Declare Function mysql_errno Lib "libmysql.dll" (ByVal hMysql As Long) As Long
Private Declare Function mysql_error Lib "libmysql.dll" (ByVal hMysql As Long) As Long

Private Declare Function mysql_select_db Lib "libmysql.dll" (ByVal hMysql As Long, ByVal dB As Long) As Long
Private Declare Function mysql_options Lib "libmysql.dll" (ByVal hMysql As Long, ByVal PropID As Long, ByVal arg As Long) As Long

Private Declare Function mysql_real_query Lib "libmysql.dll" (ByVal hMysql As Long, ByVal q As Long, ByVal Length As Long) As Long
Private Declare Function mysql_send_query Lib "libmysql.dll" (ByVal hMysql As Long, ByVal q As Long, ByVal Length As Long) As Long
Private Declare Function mysql_escape_string Lib "libmysql.dll" (ByVal to_ As String, ByVal from_ As String, ByVal Length As Long) As Long

Private Declare Function mysql_store_result Lib "libmysql.dll" (ByVal hMysql As Long) As Long
Private Declare Function mysql_free_result Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long

Private Declare Function mysql_field_count Lib "libmysql.dll" (ByVal hMysql As Long) As Long
Private Declare Function mysql_fetch_field Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                                            'Return the field struct pointer
Private Declare Function mysql_num_rows Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                                               'Return the 8 bit lenght pointer, myulonglong
Private Declare Function mysql_fetch_row Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                                              'Return the row struct pointer

Private Declare Function mysql_get_server_info Lib "libmysql.dll" (ByVal hMysql As Long) As Long
Private Declare Function mysql_stat Lib "libmysql.dll" (ByVal hMysql As Long) As Long

Private Declare Function mysql_list_dbs Lib "libmysql.dll" (ByVal hMysql As Long, ByVal wild As Long) As Long                                                 'Return the True or False
Private Declare Function mysql_list_tables Lib "libmysql.dll" (ByVal hMysql As Long, ByVal wild As Long) As Long                                                 'Return the True or False
Private Declare Function mysql_list_fields Lib "libmysql.dll" (ByVal hMysql As Long, ByVal table As Long, ByVal wild As Long) As Long                                                'Return the True or False
Private Declare Function mysql_fetch_lengths Lib "libmysql.dll" (ByVal lResultPointer As Long) As Long                              'returns * unsigned long

Private Declare Function GetSystemDirectoryA Lib "kernel32.dll" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathNameA Lib "kernel32.dll" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Event Connected(ByVal lAPIHandle As Long, _
                       ByVal sHost As String, _
                       ByVal sUser As String, _
                       ByVal sPass As String, _
                       ByVal sDatabase As String, _
                       ByVal lPort As String, _
                       ByVal sUnixSocket As String)
                       
Public Event Disconnected(ByVal lAPIHandle As Long, _
                          ByVal sHost As String)
                          
Public Event BatchQueryProgress(ByVal lTotal As Long, _
                                ByVal lCurrent As Long, _
                                ByVal eProgressType As BATHPROGRESSTYPE)
                                
Public Event OnError(ByVal ErrCode As Long, ByVal ErrDescription As String)

Private lRet                                    As Long
Private bRet                                    As Boolean

Private APIHandle                               As Long
Private lResult                                 As Long

Private IsConnected                             As Boolean
Private OpenIsReader                            As Boolean

Private i                                       As Long
Private j                                       As Long

Private mUsername                               As String
Private mPassword                               As String
Private mServerAddress                          As String
Private mServerPort                             As Long
Private mDatabaseName                           As String
Private mUnixSocket                             As String
Private mConnectionTimeout                      As Long
Private mCharset                                As String
Private mGiveError                              As Boolean

Public Property Get State() As MYSTATE
    If (IsConnected) Then
        State = MY_CONNECTED
    Else
        State = MY_NOT_CONNECTED
    End If
End Property

Public Property Get ObjectPointer() As Long
    ObjectPointer = ObjPtr(Me)
End Property

Public Property Get GiveError() As Boolean
    GiveError = mGiveError
End Property

Public Property Let GiveError(ByVal Value As Boolean)
    mGiveError = Value
End Property

Public Property Get ConnectionHandle() As Long
    ConnectionHandle = APIHandle
End Property

Public Property Let ConnectionHandle(ByVal Value As Long)
    APIHandle = Value
    IsConnected = True
End Property

Public Property Get ConnectionTimeout() As Long
    ConnectionTimeout = mConnectionTimeout
End Property

Public Property Let ConnectionTimeout(ByVal Value As Long)
    If Not (IsConnected) Then
        mConnectionTimeout = Value
        lRet = mysql_options(APIHandle, 0, StrPtr(StrConv(CStr(Value), vbFromUnicode)))
        If lRet <> 0 Then
            If (mGiveError) Then
                Err.Raise vbObjectError + GetErrorCode, GetErrorCode & ":" & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4:" & " Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Property

Public Property Get Username() As String
    Username = mUsername
End Property

Public Property Let Username(ByVal Value As String)
    If Not (IsConnected) Then
        If (Len(Trim(Value)) = 0) Then Value = "root"
        mUsername = Value
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4: Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Property

Public Property Get Charset() As String
    Charset = mCharset
End Property

Public Property Let Charset(ByVal Value As String)
    If (IsConnected) Then
        If (CLng(Left(Ptr2Str(mysql_get_server_info(APIHandle)), 1)) < 5) Then
            If (GiveError) Then
                'Err.Raise vbObjectError - 5, , "-5: Veritabaný sunucu versiyonu 5.0'dan düþük olduðu için 'SET NAMES' komutu iþletilemiyor. Lütfen sunucu versiyonunuzu güncelleyiniz"
                Call MsgBox("MySQL server version required 5.0 or higher.")
            Else
                RaiseErrorEvent -5
            End If
        Else
            If (Len(Trim(Value))) Then
                mCharset = Value
                lRet = mysql_send_query(APIHandle, StrPtr(StrConv("SET NAMES '" & mCharset & "'", vbFromUnicode)), Len("SET NAMES '" & mCharset & "'"))
                If lRet = 0 Then
                    If (GiveError) Then
                        Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
                    Else
                        RaiseErrorEvent
                    End If
                End If
            End If
        End If
    Else
        If (Len(Trim(Value))) Then
            mCharset = Value
        End If
    End If
End Property

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal Value As String)
    If Not (IsConnected) Then
        mPassword = Value
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4:" & " Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Property

Public Property Get ServerAddress() As String
    ServerAddress = mServerAddress
End Property

Public Property Let ServerAddress(ByVal Value As String)
    If Not (IsConnected) Then
        If (Len(Trim(Value)) = 0) Then Value = "localhost"
        mServerAddress = Value
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4:" & " Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Property

Public Property Get ServerPort() As Long
    ServerPort = mServerPort
End Property

Public Property Let ServerPort(ByVal Value As Long)
    If Not (IsConnected) Then
        If (Value = 0) Then Value = 3306
        mServerPort = Value
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4:" & " Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Property

Public Property Get DatabaseName() As String
    DatabaseName = mDatabaseName
End Property

Public Property Let DatabaseName(ByVal Value As String)
    If Not (IsConnected) Then
        mDatabaseName = Value
    Else
        lRet = mysql_select_db(APIHandle, StrPtr(StrConv(Value, vbFromUnicode)))
        If lRet <> 0 Then
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    End If
End Property

Public Property Get UnixSocket() As String
    UnixSocket = mUnixSocket
End Property

Public Property Let UnixSocket(ByVal Value As String)
    If Not (IsConnected) Then
        If (Value = 0) Then Value = 3306
        mUnixSocket = Value
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4:" & " Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Property

Public Function Connect(Optional ByVal sServerAddress As String = vbNullString, _
                        Optional ByVal sUsername As String = vbNullString, _
                        Optional ByVal sPassword As String = vbNullString, _
                        Optional ByVal sDatabase As String = vbNullString, _
                        Optional ByVal lServerPort As Long = -1, _
                        Optional ByVal sUnixSocket As String = vbNullString) As Boolean
    
   
    If Not (IsConnected) Then
        APIHandle = mysql_init(APIHandle)
        If APIHandle = 0 Then
            If (GiveError) Then
                Err.Raise vbObjectError - 1, , "-1: MySQL API Arabirimi baþlatýlamadý."
            Else
                RaiseErrorEvent -1
            End If
            Exit Function
        End If
        
        If Len(Trim(sServerAddress)) > 0 Then mServerAddress = sServerAddress
        If Len(Trim(sUsername)) > 0 Then mUsername = sUsername
        If Len(Trim(sPassword)) > 0 Then mPassword = sPassword
        If Len(Trim(sDatabase)) > 0 Then mDatabaseName = sDatabase
        If lServerPort > 0 Then mServerPort = lServerPort
        If Len(Trim(sUnixSocket)) > 0 Then mUnixSocket = sUnixSocket
        
        lRet = mysql_real_connect(APIHandle, _
                                  StrPtr(StrConv(mServerAddress, vbFromUnicode)), _
                                  StrPtr(StrConv(mUsername, vbFromUnicode)), _
                                  StrPtr(StrConv(mPassword, vbFromUnicode)), _
                                  StrPtr(StrConv(mDatabaseName, vbFromUnicode)), _
                                  mServerPort, _
                                  StrPtr(StrConv(mUnixSocket, vbFromUnicode)), _
                                  0)

        If (lRet <> 0) Then
            Connect = True
            IsConnected = True
'Automatic Setting: Turkish Charset
'            If CLng(Left(Ptr2Str(mysql_get_server_info(APIHandle)), 1)) < 5 Then
'                If (GiveError) Then
'                    Err.Raise vbObjectError - 5, , "-5: Veritabaný sunucu versiyonu 5.0'dan düþük olduðu için 'SET NAMES' komutu iþletilemiyor. Lütfen sunucu versiyonunuzu güncelleyiniz."
'                Else
'                    RaiseErrorEvent -5
'                End If
'            Else
'                lRet = mysql_send_query(APIHandle, StrPtr(StrConv("SET NAMES '" & mCharset & "'", vbFromUnicode)), Len("SET NAMES '" & mCharset & "'"))
'                If lRet <> 0 Then
'                    If (GiveError) Then
'                        Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
'                    Else
'                        RaiseErrorEvent
'                    End If
'                End If
'            End If
            
            RaiseEvent Connected(APIHandle, mServerAddress, mUsername, mPassword, mDatabaseName, mServerPort, mUnixSocket)
        Else
            Connect = False
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 4, , "-4:" & " Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir."
        Else
            RaiseErrorEvent -4
        End If
    End If
End Function

Public Property Get GetServerVersion() As String
    If (IsConnected) Then
        GetServerVersion = Ptr2Str(mysql_get_server_info(APIHandle))
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Property

Public Property Get GetServerState() As String
    If (IsConnected) Then
        GetServerState = Ptr2Str(mysql_stat(APIHandle))
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Property

Public Function SelectDatabase(ByVal sDatabaseName As String) As Boolean
    If Len(Trim(sDatabaseName)) = 0 Then Exit Function
    
    If (IsConnected) Then
        lRet = mysql_select_db(APIHandle, StrPtr(StrConv(sDatabaseName, vbFromUnicode)))
        If lRet = 0 Then
            SelectDatabase = True
        Else
            SelectDatabase = False
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function ExecuteReader(ByVal sQuery As String) As MyReader
    If sQuery = "#closereader#" Then OpenIsReader = False: Exit Function
    
    If Len(Trim(sQuery)) = 0 Then Exit Function
    
    If OpenIsReader Then
        Err.Raise vbObjectError, "MyConnection:ExecuteReader", "Zaten açýk bir MyReader nesnesi var."
    End If
    
    Dim lExe                    As Long
    Dim mReader                 As MyReader
    
    lExe = -1
    
    If (IsConnected) Then
        lRet = mysql_real_query(APIHandle, StrPtr(StrConv(sQuery, vbFromUnicode)), Len(sQuery))
        If lRet = 0 Then
            Set ExecuteReader = New MyReader
            OpenIsReader = True
            Set ExecuteReader.CnnObject = Me
            ExecuteReader.ConnectionPointer = APIHandle
            ExecuteReader.ResultPointer = lResult
        Else
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function Execute(ByVal sQuery As String) As Long
    If Len(Trim(sQuery)) = 0 Then Exit Function
    
    If (IsConnected) Then
        lRet = mysql_real_query(APIHandle, StrPtr(StrConv(sQuery, vbFromUnicode)), Len(sQuery))
        If lRet = 0 Then
            lResult = mysql_store_result(APIHandle)
            If lResult Then
                Execute = lResult
            End If
        Else
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function ExecuteNonQuery(ByVal sQuery As String) As Long
    If Len(Trim(sQuery)) = 0 Then Exit Function
    
    If (IsConnected) Then
        lRet = mysql_real_query(APIHandle, StrPtr(StrConv(sQuery, vbFromUnicode)), Len(sQuery))
        If lRet <> 0 Then
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
            Exit Function
        End If
        ExecuteNonQuery = lRet
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function BathRealQuery(ByVal sBathQuery As String) As Boolean
    If Len(Trim(sBathQuery)) = 0 Then Exit Function
    
    Dim lArray()            As String
    Dim lMax                As Long
    
    If (IsConnected) Then
        lArray = SQLParse(sBathQuery)
        lMax = UBound(lArray)
        
        For i = 0 To UBound(lArray)
            lRet = mysql_real_query(APIHandle, StrPtr(StrConv(lArray(i), vbFromUnicode)), Len(lArray(i)))
            If lRet <> 0 Then
                If (GiveError) Then
                    Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
                Else
                    RaiseErrorEvent
                End If
                Exit Function
            End If
            DoEvents
            RaiseEvent BatchQueryProgress(lMax, i, MEMORY_PROGRESS)
        Next i
        
        BathRealQuery = True
        Erase lArray
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function BathRealQueryFromFile(ByVal sScriptFilePath As String) As Boolean
    If Len(Trim(sScriptFilePath)) = 0 Then Exit Function
    
    Dim hFile               As Long
    Dim fLen                As Long
    Dim fBytes()            As Byte
    Dim lMax                As Long
    
    If Not CBool(PathFileExists(sScriptFilePath)) Then
        If (mGiveError) Then
            Err.Raise vbObjectError - 2, , "-2:" & " Dosya bulunamadý."
        Else
            RaiseErrorEvent -2
        End If
        Exit Function
    End If
    
    hFile = CreateReadFileHandle(sScriptFilePath)
        fLen = GetFileSize(hFile, ByVal 0&)
        ReDim fBytes(fLen) As Byte
    
        ReadFile hFile, fBytes(0), fLen, lRet, ByVal 0&
        sScriptFilePath = StrConv(fBytes, vbUnicode)
    CloseHandle hFile
    Erase fBytes

    Dim lArray()            As String
    
    If (IsConnected) Then
        lArray = SQLParse(sScriptFilePath)
        lMax = UBound(lArray)
        
        For i = 0 To UBound(lArray)
            lRet = mysql_real_query(APIHandle, StrPtr(StrConv(lArray(i), vbFromUnicode)), Len(lArray(i)))
            If lRet <> 0 Then
                If (GiveError) Then
                    Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
                Else
                    RaiseErrorEvent
                End If
                Exit Function
            End If
            DoEvents
            RaiseEvent BatchQueryProgress(lMax, i, FILE_PROGRESS)
        Next i
        
        BathRealQueryFromFile = True
        Erase lArray
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
    
    sScriptFilePath = Empty
End Function

Public Function BathRealQueryFromResource(ByVal lResurceID As Long, Optional ByVal sResurceType As String = "CUSTOM") As Boolean
    Dim sBathQuery          As String
    sBathQuery = StrConv(LoadResData(lResurceID, sResurceType), vbUnicode)
    
    If Len(Trim(sBathQuery)) = 0 Then Exit Function
    
    Dim lArray()            As String
    Dim lMax                As Long
    
    If (IsConnected) Then
        lArray = SQLParse(sBathQuery)
        lMax = UBound(lArray)
        
        For i = 0 To UBound(lArray)
            lRet = mysql_real_query(APIHandle, StrPtr(StrConv(lArray(i), vbFromUnicode)), Len(lArray(i)))
            If lRet <> 0 Then
                If (GiveError) Then
                    Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
                Else
                    RaiseErrorEvent
                End If
                Exit Function
            End If
            DoEvents
            RaiseEvent BatchQueryProgress(lMax, i, RESOURCE_PROGRESS)
        Next i
        
        BathRealQueryFromResource = True
        Erase lArray
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function SendQuery(ByVal sQuery As String) As Boolean
   
    If Len(Trim(sQuery)) = 0 Then Exit Function
    
    If (IsConnected) Then
        lRet = mysql_send_query(APIHandle, StrPtr(StrConv(sQuery, vbFromUnicode)), Len(sQuery))
        If lRet = 0 Then
            SendQuery = True
        Else
            SendQuery = False
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function DatabaseExists(ByVal sDatabaseName As String) As Boolean
    If Len(Trim(sDatabaseName)) = 0 Then Exit Function
    
    Dim lName                   As Long
    
    If (IsConnected) Then
        
        
        lResult = mysql_list_dbs(APIHandle, StrPtr(StrConv(vbNullString, vbFromUnicode)))
        If lResult Then
            For i = 1 To mysql_num_rows(lResult)
                lRet = mysql_fetch_row(lResult)
                If lRet Then
                    CopyMemory lName, ByVal lRet, SIZE_OF_CHAR * 1
                    If lstrcmpi(lName, sDatabaseName) = 0 Then
                        DatabaseExists = True
                        Exit For
                    End If
                End If
            Next i
            lRet = mysql_free_result(lResult)
        End If
    Else
        If (mGiveError) Then
            Err.Raise vbObjectError - 3, , "-3:" & " Aktif baðlantý yok."
        Else
            RaiseErrorEvent -3
        End If
    End If
End Function

Public Function ListDatabases() As Long
    If (IsConnected) Then
        lResult = mysql_list_dbs(APIHandle, StrPtr(StrConv(vbNullString, vbFromUnicode)))
        If lResult Then
            ListDatabases = lResult
        Else
            If (GiveError) Then
                Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
            Else
                RaiseErrorEvent
            End If
        End If
    Else
        RaiseErrorEvent -3
    End If
End Function

Public Function TableExists(ByVal sDatabaseName As String, ByVal sTableName As String) As Boolean
    If Len(Trim(sDatabaseName)) = 0 Then Exit Function
    If Len(Trim(sTableName)) = 0 Then Exit Function

    If DatabaseExists(sDatabaseName) Then
        SelectDatabase sDatabaseName
    Else
        TableExists = False
        RaiseErrorEvent -6
        Exit Function
    End If

    lResult = mysql_list_tables(APIHandle, StrPtr(StrConv(sTableName, vbFromUnicode)))
    If lResult Then
        TableExists = Not (mysql_num_rows(lResult) = 0)
        lRet = mysql_free_result(lResult)
    Else
        If (GiveError) Then
            Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
        Else
            RaiseErrorEvent
        End If
    End If
End Function

Public Function ListTables(ByVal sDatabaseName As String) As Long
    If Len(Trim(sDatabaseName)) = 0 Then Exit Function

    If DatabaseExists(sDatabaseName) Then
        SelectDatabase sDatabaseName
    Else
        RaiseErrorEvent -6
        Exit Function
    End If

    lResult = mysql_list_tables(APIHandle, StrPtr(StrConv(vbNullString, vbFromUnicode)))
    If lResult Then
        ListTables = lResult
    Else
        If (GiveError) Then
            Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
        Else
            RaiseErrorEvent
        End If
    End If
End Function

Public Function EscapeString(ByVal sString As String) As String
    Dim sTmp            As String
    Dim sLen            As Long

    sLen = Len(sString)
    sTmp = Space(sLen * 2 + 1)

    sLen = mysql_escape_string(sTmp, sString, sLen)
    EscapeString = Left(sTmp, sLen)
End Function

Public Function FieldExists(ByVal sDatabaseName As String, _
                            ByVal sTableName As String, _
                            ByVal sFieldName As String) As Boolean

    If Len(Trim(sDatabaseName)) = 0 Then Exit Function
    If Len(Trim(sTableName)) = 0 Then Exit Function
    If Len(Trim(sFieldName)) = 0 Then Exit Function

    If TableExists(sDatabaseName, sTableName) Then
        SelectDatabase sDatabaseName
    Else
        FieldExists = False
        RaiseErrorEvent -7
        Exit Function
    End If

    lResult = mysql_list_fields(APIHandle, StrPtr(StrConv(sTableName, vbFromUnicode)), StrPtr(StrConv(sFieldName, vbFromUnicode)))
    If lResult Then
        FieldExists = (mysql_num_rows(lResult) = 0)
        lRet = mysql_free_result(lResult)
    Else
        If (GiveError) Then
            Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
        Else
            RaiseErrorEvent
        End If
    End If
End Function

Public Function ListFields(ByVal sDatabaseName As String, ByVal sTableName As String) As Long
    If Len(Trim(sDatabaseName)) = 0 Then Exit Function
    If Len(Trim(sTableName)) = 0 Then Exit Function
    
    If TableExists(sDatabaseName, sTableName) Then
        SelectDatabase sDatabaseName
    Else
        RaiseErrorEvent -7
        Exit Function
    End If

    lResult = mysql_list_fields(APIHandle, StrPtr(StrConv(sTableName, vbFromUnicode)), StrPtr(StrConv(vbNullString, vbFromUnicode)))
    If lResult Then
        ListFields = lResult
    Else
        If (GiveError) Then
            Err.Raise vbObjectError + GetErrorCode, , GetErrorCode & ": " & GetErrorDescription
        Else
            RaiseErrorEvent
        End If
    End If
End Function

Public Sub FreeResult(ByVal ResultPointer As Long)
    lRet = mysql_free_result(ResultPointer)
End Sub

Public Function Disconnect() As Boolean
    If (IsConnected) Then
        lRet = mysql_close(APIHandle)
        IsConnected = False
        Disconnect = True
        RaiseEvent Disconnected(APIHandle, mServerAddress)
        APIHandle = 0
        lRet = 0
        bRet = False
    Else
        Disconnect = False
        RaiseEvent OnError(-3, " Aktif baðlantý yok.")
    End If
End Function

Private Function GetErrorCode() As Long
    GetErrorCode = mysql_errno(APIHandle)
End Function

Private Function GetErrorDescription() As String
    GetErrorDescription = Ptr2Str(mysql_error(APIHandle))
End Function

Public Function Ptr2Str(ByVal lPtr As Long) As String
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

Private Sub Class_Initialize()
    mUsername = "root"
    mPassword = ""
    mServerAddress = "localhost"
    mServerPort = 3306
    mDatabaseName = ""
    mUnixSocket = ""
    mConnectionTimeout = 15
    mCharset = "latin5"
    mGiveError = True
End Sub

Private Sub Class_Terminate()
    bRet = Disconnect
End Sub

Public Sub ConvertADORS(ByVal ResultPointer As Long, ByRef RSObject As Object)
    Dim FieldCount              As Long
    Dim RowCount                As Long
    Dim FieldNames()            As Long

    Dim FieldName               As Long
    Dim FieldLenght             As Long
    Dim FieldType               As Long

    Dim BlobIndex               As Long

    FieldCount = mysql_field_count(APIHandle)
    RowCount = mysql_num_rows(ResultPointer)
    ReDim FieldNames(1 To FieldCount) As Long

    Dim sl As Long
    With RSObject

        For i = 1 To FieldCount
            lRet = mysql_fetch_field(ResultPointer)
            If lRet Then
                CopyMemory FieldName, ByVal lRet, 4
                CopyMemory FieldLenght, ByVal lRet + 28, 4
                CopyMemory FieldType, ByVal lRet + 76, 4
                
                If FieldType = 252 Then
                    .Fields.Append Ptr2Str(FieldName), 12, 65535
                    BlobIndex = i
                ElseIf FieldType = 246 Then 'Is Decimal Field To adBSTR
                    .Fields.Append Ptr2Str(FieldName), 8, 127
                Else
                    .Fields.Append Ptr2Str(FieldName), MyFieldTypeToADOFieldType(FieldType), FieldLenght + 1, 32
                End If
            End If
        Next i

        .Open

        For i = 1 To RowCount
            lRet = mysql_fetch_row(ResultPointer)
            If lRet Then
                CopyMemory FieldNames(1), ByVal lRet, SIZE_OF_CHAR * FieldCount
                .AddNew
                 For j = 1 To FieldCount
                    If .Fields(j - 1).Type <> 7 Then
                        If j = BlobIndex Then 'Is Blob Field
                            Dim DataLength          As Long
                            Dim Data()              As Byte
                            
                            CopyMemory DataLength, ByVal mysql_fetch_lengths(ResultPointer), IIf((BlobIndex - 1) = 0, 1, BlobIndex - 1) * SIZE_OF_CHAR
                            
                            ReDim Data(DataLength) As Byte
                            
                            CopyMemory Data(0), ByVal FieldNames(j), DataLength

                            .Fields(j - 1) = Data
                            Erase Data
                        Else
                            .Fields(j - 1) = Ptr2Str(FieldNames(j))
                        End If
                    Else
                        Dim sTmp As String
                        sTmp = Ptr2Str(FieldNames(j))
                        If Len(Trim(sTmp)) > 1 Then
                        .Fields(j - 1) = DateSerial(CInt(Left(sTmp, 4)), CInt(Mid(sTmp, 6, 2)), CInt(Right(sTmp, 2)))
                            Else
                        .Fields(j - 1) = Null
                        End If
                    End If
                 Next j
            End If
        Next i
        .UpdateBatch 4
    End With

    Erase FieldNames

    If Not (RSObject.EOF And RSObject.BOF) Then RSObject.MoveFirst
    Call Me.FreeResult(ResultPointer)
End Sub

Private Function MyFieldTypeToADOFieldType(ByVal MyField As Long) As Long
    Select Case MyField
        Case 3
            MyFieldTypeToADOFieldType = 20
        Case 8
            MyFieldTypeToADOFieldType = 20
        Case 254
            MyFieldTypeToADOFieldType = 129
        Case 4
            MyFieldTypeToADOFieldType = 6
        Case 10
            MyFieldTypeToADOFieldType = 7
        Case 11
            MyFieldTypeToADOFieldType = 8
        Case 7
            MyFieldTypeToADOFieldType = 135
        Case 0, 246
            MyFieldTypeToADOFieldType = 14
        Case 5
            MyFieldTypeToADOFieldType = 5
        Case 6
            MyFieldTypeToADOFieldType = 0
        Case 9
            MyFieldTypeToADOFieldType = 3
        Case 2
            MyFieldTypeToADOFieldType = 2
        Case 1
            MyFieldTypeToADOFieldType = 16
        Case 253
            MyFieldTypeToADOFieldType = 200
        Case 255, 252, 250, 251, 248, 247, 16, 14
            MyFieldTypeToADOFieldType = 12
        End Select
End Function

Private Sub RaiseErrorEvent(Optional ByVal lErrorCode As Long = -21565)
    Select Case lErrorCode
        Case -21565
            RaiseEvent OnError(GetErrorCode, GetErrorDescription)
        Case -1
            RaiseEvent OnError(lErrorCode, "MySQL API Arabirimi baþlatýlamadý.")
        Case -2
            RaiseEvent OnError(lErrorCode, "Dosya bulunamadý.")
        Case -3
            RaiseEvent OnError(lErrorCode, "Aktif baðlantý yok.")
        Case -4
            RaiseEvent OnError(lErrorCode, "Aktif baðlantý var. Öncelikle aktif baðlantýyý kapatmanýz gerekir.")
        Case -5
            RaiseEvent OnError(lErrorCode, "Veritabaný sunucu versiyonu 5.0'dan düþük olduðu için 'SET NAMES' komutu iþletilemiyor. Lütfen sunucu versiyonunuzu güncelleyiniz.")
        Case -6
            RaiseEvent OnError(lErrorCode, "Veritabaný Bulunamadý!")
        Case -7
            RaiseEvent OnError(lErrorCode, "Tablo Bulunamadý!")
    End Select
End Sub

Private Function SQLParse(ByVal sText As String) As String()
    Dim Dizi()              As String
    Dim dC                  As Long
    Dim SrcStr              As String
    Dim LastPos             As Long
    Dim Flag                As Boolean
    Dim i                   As Long
    
    dC = -1
    LastPos = 1
    SrcStr = sText
    Flag = True
    
    For i = 1 To Len(SrcStr)
        If InStr(1, Mid(SrcStr, i, 5), "BEGIN") Then
            Flag = False
        End If
        If InStr(1, Mid(SrcStr, i, 4), "END;") Then
            Flag = True
        End If
        If Mid(SrcStr, i, 1) = "'" Then
            If Mid(SrcStr, i - 1, 1) <> "\" Then
                Flag = Not Flag
            End If
        End If
        
        If Flag Then
            If Mid(SrcStr, i, 1) = ";" Then
                i = i + 1
                dC = dC + 1
                ReDim Preserve Dizi(dC) As String
                Dizi(dC) = Mid(SrcStr, LastPos, i - LastPos)
                Dizi(dC) = SqlTemizle(Dizi(dC))
                LastPos = i
            End If
        End If
    Next i
    SQLParse = Dizi()
End Function

Private Function SqlTemizle(ByVal sText As String) As String
    Dim i                   As Long
    Dim Dizi()              As String
    
    If InStr(1, sText, vbCrLf) Then
        Dizi = Split(sText, vbCrLf)
        For i = 0 To UBound(Dizi)
            If Left(Dizi(i), 1) = "#" Then
                Dizi(i) = ""
            End If
        Next i
    
        SqlTemizle = Join(Dizi, vbCrLf)
        SqlTemizle = Replace(SqlTemizle, vbCrLf & vbCrLf, "")
        If Left(SqlTemizle, 2) = vbCrLf Then
            SqlTemizle = Right(SqlTemizle, Len(SqlTemizle) - 2)
        End If
        If Right(SqlTemizle, 2) = vbCrLf Then
            SqlTemizle = Left(SqlTemizle, Len(SqlTemizle) - 2)
        End If
    Else
        SqlTemizle = sText
    End If
    SqlTemizle = Trim(SqlTemizle)
End Function




