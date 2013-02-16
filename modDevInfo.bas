Attribute VB_Name = "modDevInfo"
'获取磁盘信息，用于判断程序是否从可移动磁盘启动，并做一些优化
Option Explicit

'获取特定磁盘信息：GetDeviceInf

Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal cb As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const OPEN_EXISTING                 As Long = 3&
Private Const FILE_SHARE_READ               As Long = &H1&
Private Const FILE_SHARE_WRITE              As Long = &H2&
Private Const GENERIC_READ                  As Long = &H80000000
Private Const IOCTL_STORAGE_QUERY_PROPERTY  As Long = &H2D1400

Private Type STORAGE_PROPERTY_QUERY
    PropertyId                              As STORAGE_PROPERTY_ID
    QueryType                               As STORAGE_QUERY_TYPE
    AdditionalParameters                    As Byte
End Type

Public Type DEVICE_INFORMATION
    Valid                                   As Boolean
    BusType                                 As STORAGE_BUS_TYPE
    Removable                               As Boolean
    VendorID                                As String
    ProductID                               As String
    ProductRevision                         As String
End Type

Private Type STORAGE_DEVICE_DESCRIPTOR
    Version                                 As Long
    Size                                    As Long
    DeviceType                              As Byte
    DeviceTypeModifier                      As Byte
    RemovableMedia                          As Byte
    CommandQueueing                         As Byte
    VendorIdOffset                          As Long
    ProductIdOffset                         As Long
    ProductRevisionOffset                   As Long
    SerialNumberOffset                      As Long
    BusType                                 As Integer
    RawPropertiesLength                     As Long
    RawDeviceProperties                     As Byte
End Type

Public Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum

Private Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0
    StorageAdapterProperty
    StorageDeviceIdProperty
End Enum

Private Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0
    PropertyExistsQuery
    PropertyMaskQuery
    PropertyQueryMaxDefined
End Enum

Public Function GetDevInfo(ByVal strDrive As String) As DEVICE_INFORMATION
    Dim hDrive          As Long
    Dim udtQuery        As STORAGE_PROPERTY_QUERY
    Dim dwOutBytes      As Long
    Dim lngResult       As Long
    Dim btBuffer(9999)  As Byte
    Dim udtOut          As STORAGE_DEVICE_DESCRIPTOR
    
    hDrive = CreateFile("\\.\" & Left$(strDrive, 1) & ":", 0, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        ByVal 0&, OPEN_EXISTING, 0, 0)

    If hDrive = -1 Then Exit Function
    
    With udtQuery
        .PropertyId = StorageDeviceProperty
        .QueryType = PropertyStandardQuery
    End With
    
    lngResult = DeviceIoControl(hDrive, IOCTL_STORAGE_QUERY_PROPERTY, _
                                udtQuery, LenB(udtQuery), _
                                btBuffer(0), UBound(btBuffer) + 1, _
                                dwOutBytes, ByVal 0&)
        
    If lngResult Then
        CpyMem udtOut, btBuffer(0), Len(udtOut)
        
        With GetDevInfo
            .Valid = True
            .BusType = udtOut.BusType
            .Removable = CBool(udtOut.RemovableMedia)
            
            If udtOut.ProductIdOffset > 0 Then _
                .ProductID = StringCopy(VarPtr(btBuffer(udtOut.ProductIdOffset)))
            If udtOut.ProductRevisionOffset > 0 Then _
                .ProductRevision = StringCopy(VarPtr(btBuffer(udtOut.ProductRevisionOffset)))
            If udtOut.VendorIdOffset > 0 Then
                .VendorID = StringCopy(VarPtr(btBuffer(udtOut.VendorIdOffset)))
            End If
        End With
    Else
        GetDevInfo.Valid = False
    End If
    
    CloseHandle hDrive
End Function

Private Function StringCopy(ByVal pBuffer As Long) As String
    Dim tmp As String
    
    tmp = Space(lstrlen(ByVal pBuffer))
    lstrcpy ByVal tmp, ByVal pBuffer
    StringCopy = Trim$(tmp)
End Function

Public Sub GetDeviceInf(ByVal Drive As String, Optional ByRef DriveName As String, Optional ByRef IsRemovable As Boolean, Optional ByRef BusType As Long, Optional ByRef BusTypeString As String, Optional ByRef VendorID As String, Optional ByRef ProductID As String, Optional ByRef ProductRevision As String)
    Dim strDriveBuffer  As String
    Dim strDrives()     As String
    Dim i               As Long
    Dim udtInfo         As DEVICE_INFORMATION
    Dim iDriveName As String
    Dim uBusTypeString
    iDriveName = Left(Trim(Drive), 1) & ":\"
    strDriveBuffer = Space(240)
    strDriveBuffer = Left$(strDriveBuffer, GetLogicalDriveStrings(Len(strDriveBuffer), strDriveBuffer))
    strDrives = Split(strDriveBuffer, Chr$(0))
    For i = 0 To UBound(strDrives) - 1
        If iDriveName = strDrives(i) Then
            udtInfo = GetDevInfo(strDrives(i))
            If udtInfo.Valid Then
                Select Case udtInfo.BusType
                    Case BusTypeUsb:        uBusTypeString = "USB"
                    Case BusType1394:       uBusTypeString = "1394"
                    Case BusTypeAta:        uBusTypeString = "ATA"
                    Case BusTypeAtapi:      uBusTypeString = "ATAPI"
                    Case BusTypeFibre:      uBusTypeString = "Fibre"
                    Case BusTypeRAID:       uBusTypeString = "RAID"
                    Case BusTypeScsi:       uBusTypeString = "SCSI"
                    Case BusTypeSsa:        uBusTypeString = "SSA"
                    Case BusTypeUnknown:    uBusTypeString = "未知"
                End Select
                If Not (IsMissing(DriveName)) Then DriveName = iDriveName
                If Not (IsMissing(IsRemovable)) Then IsRemovable = udtInfo.Removable Or uBusTypeString = "USB" Or uBusTypeString = "1394"
                If Not (IsMissing(BusType)) Then BusType = udtInfo.BusType
                If Not (IsMissing(BusTypeString)) Then BusTypeString = uBusTypeString
                If Not (IsMissing(VendorID)) Then VendorID = udtInfo.VendorID
                If Not (IsMissing(ProductID)) Then ProductID = udtInfo.ProductID
                If Not (IsMissing(ProductRevision)) Then ProductRevision = udtInfo.ProductRevision
            End If
        Exit For
        End If
    Next
End Sub
