VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpcServiciosyWea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////////
'//UpcServiciosyWea
'//
'// Fecha:        Wednesday, June 20, 2007 17:01:27
'// Autor:        Mauricio Díaz
'// Organización: SleepLess Software
'// Descripción:  Clase para crear, arrancar, detener, pausar y eliminar servicios
'/////////////////////////////////////////////////////////////////////////////
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Private Const SERVICE_DISABLED As Long = &H4
Private Const SERVICE_DEMAND_START As Long = &H3
Private Const SERVICE_AUTO_START  As Long = &H2
Private Const SERVICE_SYSTEM_START As Long = &H1
Private Const SERVICE_BOOT_START As Long = &H0

Private Const SERVICE_WIN32_OWN_PROCESS = &H10&
Private Const SERVICE_WIN32_SHARE_PROCESS = &H20&
Private Const SERVICE_WIN32 = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS

Private Const SERVICES_ACTIVE_DATABASE = "ServicesActive"

Private Const SERVICE_CONTROL_STOP = &H1
Private Const SERVICE_CONTROL_PAUSE = &H2
Private Const SERVICE_CONTROL_CONTINUE = &H3

Private Const SERVICE_STOPPED = &H1
Private Const SERVICE_START_PENDING = &H2
Private Const SERVICE_STOP_PENDING = &H3
Private Const SERVICE_RUNNING = &H4
Private Const SERVICE_CONTINUE_PENDING = &H5
Private Const SERVICE_PAUSE_PENDING = &H6
Private Const SERVICE_PAUSED = &H7
'Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS



Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SC_MANAGER_CONNECT = &H1
Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Private Const SC_MANAGER_LOCK = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Private Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
'Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4

Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_ACTIVE = &H1
Private Const SERVICE_INACTIVE = &H2

Private Const GENERIC_READ = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER = 122
Private Const ERROR_SERVICE_ALREADY_RUNNING As Long = 1056&
Private Const ERROR_SERVICE_MARKED_FOR_DELETE As Long = 1072&
Private Const ERROR_SERVICE_EXISTS As Long = 1073&
Private Const ERROR_SERVICE_DATABASE_LOCKED As Long = 1055&
Private Const ERROR_SERVICE_DEPENDENCY_DELETED As Long = 1075&
Private Const ERROR_SERVICE_DEPENDENCY_FAIL As Long = 1068&
Private Const ERROR_SERVICE_DISABLED As Long = 1058&
Private Const ERROR_SERVICE_LOGON_FAILED As Long = 1069&
Private Const ERROR_SERVICE_NO_THREAD As Long = 1054&
Private Const ERROR_SERVICE_REQUEST_TIMEOUT As Long = 1053&
Private Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL As Long = 1061&
Private Const ERROR_SERVICE_NOT_ACTIVE As Long = 1062&
Private Const ERROR_MORE_DATA = 234

Private Const ERROR_INVALID_HANDLE As Long = 6&
Private Const ERROR_INVALID_NAME As Long = 123&
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_INVALID_SERVICE_ACCOUNT As Long = 1057&
Private Const ERROR_INVALID_SERVICE_CONTROL As Long = 1052&

Private Const ERROR_PATH_NOT_FOUND As Long = 3&
Private Const ERROR_ACCESS_DENIED As Long = 5&
Private Const ERROR_CIRCULAR_DEPENDENCY As Long = 1059&
Private Const ERROR_DUP_NAME As Long = 52&
Private Const ERROR_DEPENDENT_SERVICES_RUNNING As Long = 1051&
Private Const ERROR_SHUTDOWN_IN_PROGRESS As Long = 1115&

Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Public Enum e_ServiceModeStart
    e_ServiceType_Disabled = 4
    e_ServiceType_Manual = 3
    e_ServiceType_Automatic = 2
    e_ServiceType_SystemStart = 1
    e_ServiceType_BootTime = 0
End Enum

Private Type ENUM_SERVICE_STATUS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS
End Type



Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal HService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal HService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal HService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal HService As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (szDest As String, szcSource As Long) As Long

Private strServiceName As String
Private StrInstallRemoveError As String

Public Property Let ServiceName(nValue As String)
    strServiceName = nValue
End Property
Public Property Get ServiceName() As String
    ServiceName = strServiceName
End Property
Public Property Get InstallRemoveServiceError() As String

    InstallRemoveServiceError = StrInstallRemoveError
End Property



Public Function ServiceStatus(Optional ByVal Computer As String = vbNullString) As String
    Dim Status As SERVICE_STATUS
    Dim hSManager As Long
    Dim HService As Long
    Dim hServiceStatus As Long
    ServiceStatus = ""
    If Trim(strServiceName) = "" Then Exit Function
    

    
    hSManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    
    If hSManager <> 0 Then
        HService = OpenService(hSManager, strServiceName, SERVICE_ALL_ACCESS)
        If HService <> 0 Then
            hServiceStatus = QueryServiceStatus(HService, Status)
            If hServiceStatus <> 0 Then
                Select Case Status.dwCurrentState
                Case SERVICE_STOPPED
                    ServiceStatus = "STOPED"
                Case SERVICE_START_PENDING
                    ServiceStatus = "START PENDING"
                    
                Case SERVICE_STOP_PENDING
                    ServiceStatus = "STOP PENDING"
                    
                Case SERVICE_RUNNING
                    ServiceStatus = "RUNNING"
                    
                Case SERVICE_CONTINUE_PENDING
                    ServiceStatus = "CONTINUE PENDING"
                    
                Case SERVICE_PAUSE_PENDING
                    ServiceStatus = "PAUSE PENDING"
                    
                Case SERVICE_PAUSED
                    ServiceStatus = "PAUSED"
                End Select
            End If
            CloseServiceHandle HService
        End If
        CloseServiceHandle hSManager
    End If
End Function

Public Function Start_Service(Optional ByVal Computer As String = vbNullString) As Boolean
    Dim hSManager As Long
    Dim HService As Long
    Dim res As Long
    res = 0
    hSManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        HService = OpenService(hSManager, strServiceName, SERVICE_ALL_ACCESS)
        If HService <> 0 Then
            res = StartService(HService, 0, 0)
            CloseServiceHandle HService
            
        End If
        CloseServiceHandle hSManager
    End If
    If res Then
        Start_Service = True
    Else
        Start_Service = False
    End If
        
End Function
Public Function Stop_Service(Optional ByVal Computer As String) As Boolean
    Dim Status As SERVICE_STATUS
    Dim hSManager As Long
    Dim HService As Long
    Dim res As Long
    res = 0
    hSManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        HService = OpenService(hSManager, strServiceName, SERVICE_ALL_ACCESS)
        If HService <> 0 Then
            res = ControlService(HService, SERVICE_CONTROL_STOP, Status)
            CloseServiceHandle HService
        End If
        CloseServiceHandle hSManager
    End If
    If res Then
        Stop_Service = True
    Else
        Stop_Service = False
    End If
End Function

Public Function Pause_Service(Optional ByVal Computer As String = vbNullString) As Boolean
    Dim Status As SERVICE_STATUS
    Dim hSManager As Long
    Dim HService As Long
    Dim res As Long
    res = 0
    hSManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        HService = OpenService(hSManager, strServiceName, SERVICE_ALL_ACCESS)
        If HService <> 0 Then
            
            res = ControlService(HService, SERVICE_CONTROL_PAUSE, Status)
            CloseServiceHandle HService
        End If
        CloseServiceHandle hSManager
    End If
    If res Then
        Pause_Service = True
    Else
        Pause_Service = False
    End If
End Function

Public Function Resume_Service(Optional ByVal Computer As String = vbNullString) As Boolean

 Dim Status As SERVICE_STATUS
    Dim hSManager As Long
    Dim HService As Long
    Dim res As Long
    res = 0
    hSManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        HService = OpenService(hSManager, strServiceName, SERVICE_ALL_ACCESS)
        If HService <> 0 Then
            
            res = ControlService(HService, SERVICE_CONTROL_CONTINUE, Status)
            CloseServiceHandle HService
        End If
        CloseServiceHandle hSManager
    End If
    If res Then
        Resume_Service = True
    Else
        Resume_Service = False
    End If
End Function

Public Function InstallService(ByVal ServiceFilePath As String, ByVal ServiceDisplayName As String, Optional ByVal Computer As String = vbNullString, Optional ByVal ServiceStartMode As e_ServiceModeStart = e_ServiceType_Automatic) As Boolean
Dim Hscm As Long
Dim HService As Long
Dim Ret As Long
Dim iFlags As Long

Select Case ServiceStartMode
    Case 0: iFlags = SERVICE_BOOT_START
    Case 1: iFlags = SERVICE_SYSTEM_START
    Case 2: iFlags = SERVICE_AUTO_START
    Case 3: iFlags = SERVICE_DEMAND_START
    Case 4: iFlags = SERVICE_DISABLED
End Select


InstallService = False
Hscm = OpenSCManager(Computer, vbNullString, SC_MANAGER_CREATE_SERVICE)
If Hscm <> 0 Then
    HService = CreateService(Hscm, strServiceName, ServiceDisplayName, SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS, iFlags, SERVICE_ERROR_NORMAL, ServiceFilePath, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    If HService <> 0 Then
        InstallService = True
        StrInstallRemoveError = "Service installed sucessfully"
    Else
        'StrInstallRemoveError
        Ret = GetLastError
        Select Case Ret
            Case ERROR_ACCESS_DENIED
                StrInstallRemoveError = "The handle to the specified service control manager database does not have SC_MANAGER_CREATE_SERVICE access."
            Case ERROR_CIRCULAR_DEPENDENCY
                StrInstallRemoveError = "A circular service dependency was specified."
            Case ERROR_DUP_NAME
                StrInstallRemoveError = "The display name already exists in the service control manager database either as a service name or as another display name."
            Case ERROR_INVALID_HANDLE
                StrInstallRemoveError = "The handle to the specified service control manager database is invalid."
            Case ERROR_INVALID_NAME
                StrInstallRemoveError = "The specified service name is invalid."
            Case ERROR_INVALID_PARAMETER
                StrInstallRemoveError = "A parameter that was specified is invalid."
            Case ERROR_INVALID_SERVICE_ACCOUNT
                StrInstallRemoveError = "The user account name specified in the lpServiceStartName parameter does not exist."
            Case ERROR_SERVICE_EXISTS
                StrInstallRemoveError = "The specified service already exists in this database."
            Case Else
                StrInstallRemoveError = "Unknown error"
        End Select
        
    End If
End If
CloseServiceHandle Hscm
CloseServiceHandle HService
End Function
Public Function RemoveService(Optional ByVal Computer As String) As Boolean
Dim Hcsm As Long
Dim HService As Long
Dim Ret As Long
Dim Error As Long
Dim StrRemServiceError As String
Ret = 0
Hscm = OpenSCManager(Computer, vbNullString, SC_MANAGER_ALL_ACCESS)
If Hscm <> 0 Then
    HService = OpenService(Hscm, strServiceName, SERVICE_ALL_ACCESS)
    If HService <> 0 Then
        Ret = DeleteService(HService)
    Else
        Error = GetLastError
         Select Case Error
            Case ERROR_ACCESS_DENIED
                StrRemServiceError = "The handle to the specified service control manager database does not have SC_MANAGER_CREATE_SERVICE access."
            Case ERROR_CIRCULAR_DEPENDENCY
                StrRemServiceError = "A circular service dependency was specified."
            Case ERROR_DUP_NAME
                StrRemServiceError = "The display name already exists in the service control manager database either as a service name or as another display name."
            Case ERROR_INVALID_HANDLE
                StrRemServiceError = "The handle to the specified service control manager database is invalid."
            Case ERROR_INVALID_NAME
                StrRemServiceError = "The specified service name is invalid."
            Case ERROR_INVALID_PARAMETER
                StrRemServiceError = "A parameter that was specified is invalid."
            Case ERROR_INVALID_SERVICE_ACCOUNT
                StrRemServiceError = "The user account name specified in the lpServiceStartName parameter does not exist."
            Case ERROR_SERVICE_EXISTS
                StrRemServiceError = "The specified service already exists in this database."
            Case Else
                StrRemServiceError = "Unknown error"
        End Select
        
    End If
End If
CloseServiceHandle Hscm
CloseServiceHandle HService
If Ret Then
    RemoveService = True
    StrInstallRemoveError = "Service removed sucessfully"
Else
    RemoveService = False
End If
End Function

Public Function ListServices(Optional ByVal Computer As String = vbNullString) As Collection

    Dim Services As New Collection
    Dim strServiceName As String * 250
    Dim lngBytesNeeded As Long
    Dim lngServicesReturned As Long
    Dim hNextUnreadEntry As Long
    Dim lngStructsNeeded As Long
    Dim lngResult As Long
    Dim lpEnumServiceStatus() As ENUM_SERVICE_STATUS
    Dim i As Long
    Dim Hscm As Long
    Hscm = OpenSCManager(Computer, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If Hscm <> 0 Then
        hNextUnreadEntry = 0
        lngResult = EnumServicesStatus(Hscm, SERVICE_WIN32, SERVICE_ACTIVE Or SERVICE_INACTIVE, ByVal &H0, &H0, lngBytesNeeded, lngServicesReturned, hNextUnreadEntry)
        If Not Err.LastDllError = ERROR_MORE_DATA Then
            MsgBox "LastDLLError = " & CStr(Err.LastDllError)
            Exit Function
        End If
        lngStructsNeeded = lngBytesNeeded / Len(lpEnumServiceStatus(0)) + 1
        ReDim lpEnumServiceStatus(lngStructsNeeded - 1)
        lngServiceStatusInfoBuffer = lngStructsNeeded * Len(lpEnumServiceStatus(0))
        
        hNextUnreadEntry = 0
        lngResult = EnumServicesStatus(Hscm, SERVICE_WIN32, SERVICE_ACTIVE Or SERVICE_INACTIVE, lpEnumServiceStatus(0), lngServiceStatusInfoBuffer, lngBytesNeeded, lngServicesReturned, hNextUnreadEntry)
        If lngResult = 0 Then
            MsgBox "EnumServicesStatus failed. LastDllError = " & CStr(Err.LastDllError)
            Exit Function
        End If
        
        For i = 0 To lngServicesReturned - 1
            lngResult = lstrcpy(ByVal strServiceName, ByVal lpEnumServiceStatus(i).lpServiceName)
            Services.Add StripTerminator(strServiceName)
            
            lngResult = lstrcpy(ByVal strServiceName, ByVal lpEnumServiceStatus(i).lpDisplayName)
            Services.Add StripTerminator(strServiceName)
            
            Select Case lpEnumServiceStatus(i).ServiceStatus.dwCurrentState
                Case SERVICE_STOPPED
                    Services.Add "STOPED"
                Case SERVICE_START_PENDING
                    Services.Add "START PENDING"
                    
                Case SERVICE_STOP_PENDING
                    Services.Add "STOP PENDING"
                    
                Case SERVICE_RUNNING
                    Services.Add "RUNNING"
                    
                Case SERVICE_CONTINUE_PENDING
                    Services.Add "CONTINUE PENDING"
                    
                Case SERVICE_PAUSE_PENDING
                    Services.Add "PAUSE PENDING"
                    
                Case SERVICE_PAUSED
                    Services.Add "PAUSED"
            End Select
        Next i

        CloseServiceHandle (Hscm)


    End If
    
    Set ListServices = Services
End Function
Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

