Attribute VB_Name = "SystemInf"
'��ȡϵͳ������Ϣ��ģ�顣

Option Explicit
Public Sys_ComputerName As String '���������
Public Sys_ComputerStatus As String '�����״̬
Public Sys_SystemType 'ϵͳ�ܹ������ͣ�
Public Sys_Manufacturer As String '����
Public Sys_Model As String '�ͺ�
Public Sys_TotalMemory '�ڴ棨��ΪLong�������
Public Sys_Domain As String '��
Public Sys_Workgroup As String '������
Public Sys_Usename As String '�û���
Public Sys_BootupState As String '����״̬���������ǰ�ȫģʽ��
Public Sys_OwnerName As String '����������
Public Sys_CreationClassName As String 'ϵͳ�ںˣ����ͣ�
Public Sys_Description As String '������ܹ������ͣ�


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
    'item.SystemStartupOptions(i)�Ǹ����飬д��������ϵͳ��������д��
    Exit For 'ֻ��ȡ��һ��
Next
End Sub
