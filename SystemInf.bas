Attribute VB_Name = "SystemInf"
'获取系统基本信息的模块。

Option Explicit
Public Sys_ComputerName As String '计算机名称
Public Sys_ComputerStatus As String '计算机状态
Public Sys_SystemType '系统架构（类型）
Public Sys_Manufacturer As String '厂家
Public Sys_Model As String '型号
Public Sys_TotalMemory '内存（设为Long会溢出）
Public Sys_Domain As String '域
Public Sys_Workgroup As String '工作组
Public Sys_Usename As String '用户名
Public Sys_BootupState As String '启动状态（正常还是安全模式）
Public Sys_OwnerName As String '所有者姓名
Public Sys_CreationClassName As String '系统内核（类型）
Public Sys_Description As String '计算机架构（类型）


Public Sub GetSystemInf()
Dim System, item, i  As Integer
Set System = GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")
For Each item In System
    Sys_ComputerName = item.name
    Sys_ComputerStatus = item.Status
    Sys_SystemType = item.SystemType
    Sys_Manufacturer = item.Manufacturer
    Sys_Model = item.Model
    Sys_TotalMemory = item.totalPhysicalMemory
    Sys_Domain = item.domain
    Sys_Workgroup = item.Workgroup
    Sys_Usename = item.username
    Sys_BootupState = item.BootupState
    Sys_OwnerName = item.PrimaryOwnerName
    Sys_CreationClassName = item.CreationClassName
    Sys_Description = item.Description
    'item.SystemStartupOptions(i)是个数组，写着启动的系统数，懒得写了
    Exit For '只获取第一个
Next
End Sub
