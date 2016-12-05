Imports System.Runtime.InteropServices

Module DesktopIconController
    Dim DesktopHandle As IntPtr
    Dim DesktopProcessID As Integer
    Dim HandleProcess As IntPtr
    Public Const SizeOfPoint As Integer = 8

    Public Declare Function GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As IntPtr
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Public Declare Function VirtualAllocEx Lib "kernel32.dll" (hProcess As IntPtr, lpAddress As IntPtr, dwSize As Integer, flAllocationType As Integer, flProtect As Integer) As IntPtr
    Public Declare Function VirtualFreeEx Lib "kernel32.dll" (hProcess As IntPtr, lpAddress As IntPtr, dwSize As Integer, dwFreeType As Integer) As Boolean
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (hwnd As IntPtr, wMsg As Integer, wParam As IntPtr, lParam As IntPtr) As Integer
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (hwnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As IntPtr) As Integer
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (hwnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer) As Integer
    Public Declare Function GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (hWnd As IntPtr, <System.Runtime.InteropServices.OutAttribute()> ByRef ProcessId As Integer) As Integer
    Public Declare Function OpenProcess Lib "kernel32" Alias "OpenProcess" (dwDesiredAccess As Integer, bInheritHandle As Boolean, dwProcessId As Integer) As IntPtr
    Public Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (hObject As IntPtr) As Integer
    Public Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (hProcess As IntPtr, lpBaseAddress As IntPtr, <System.Runtime.InteropServices.OutAttribute()> ByRef lpBuffer As Point, nSize As Integer, <System.Runtime.InteropServices.OutAttribute()> ByRef lpNumberOfBytesRead As Integer) As Boolean
    Public Declare Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (hProcess As IntPtr, lpBaseAddress As IntPtr, lpBuffer As IntPtr, nSize As Integer, <System.Runtime.InteropServices.OutAttribute()> ByRef lpNumberOfBytesRead As Integer) As Boolean
    Public Declare Function WriteProcessMemory Lib "kernel32" Alias "WriteProcessMemory" (hProcess As IntPtr, lpBaseAddress As IntPtr, lpBuffer As IntPtr, nSize As Integer, <System.Runtime.InteropServices.OutAttribute()> ByRef lpNumberOfBytesWritten As Integer) As Integer

    Public Const PROCESS_VM_READ As Integer = 16
    Public Const PROCESS_VM_OPERATION As Integer = 8
    Public Const PROCESS_VM_WRITE As Integer = 32
    Public Const MEM_COMMIT As Integer = 4096
    Public Const PAGE_READWRITE As Integer = 4
    Public Const LVM_GETITEMCOUNT As Integer = 4100
    Public Const LVM_SETITEMPOSITION As Integer = 4111 '64位系统设置坐标
    Public Const LVM_SETITEMPOSITION32 As Integer = 4145 '32位系统设置坐标
    Public Const LVM_GETITEMPOSITION As Integer = 4112
    Public Const LVM_GETITEMTEXT As Integer = 4211
    Public Const MAX_PATH As Integer = 260

    Public Class DesktopIcon
        Public Text As String
        Public Position As Point
    End Class

    Public Structure TagLVItemW
        Public mask As UInteger
        Public iItem As Integer
        Public iSubItem As Integer
        Public state As UInteger
        Public stateMask As UInteger
        Public pszText As Long '32位系统使用 Int32
        Public cchTextMax As Integer
        Public iImage As Integer
        Public lParam As Integer
        Public iIndent As Integer
        Public iGroupId As Integer
        Public cColumns As UInteger
        Public puColumns As IntPtr
        Public piColFmt As IntPtr
        Public iGroup As Integer
    End Structure

    Public Function GetLvItemSize() As Integer
        Return Marshal.SizeOf(GetType(TagLVItemW))
    End Function

    ''' <summary>
    ''' 获取桌面进程信息
    ''' </summary>
    Public Sub OpenExplorer()
        DesktopHandle = GetDesktopIconHandle()
        GetWindowThreadProcessId(DesktopHandle, DesktopProcessID)
        HandleProcess = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, DesktopProcessID)
    End Sub

    ''' <summary>
    ''' 销毁桌面进程句柄
    ''' </summary>
    Public Sub CloseExplorer()
        CloseHandle(HandleProcess)
    End Sub

    ''' <summary>
    ''' 根据桌面图标文本获取图标位置坐标
    ''' </summary>
    ''' <param name="IconText">欲获取的图标文本</param>
    ''' <returns>指定文本的图标位置坐标(有可能是多组坐标)</returns>
    Public Function GetDesktopIconPosition(ByVal IconText As String) As Point()
        Dim IconPosition(0) As Point
        Dim PIconItem As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, GetLvItemSize(), MEM_COMMIT, PAGE_READWRITE)
        Dim PIconText As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, MAX_PATH * 2, MEM_COMMIT, PAGE_READWRITE)
        Dim PIconPosition As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, SizeOfPoint, MEM_COMMIT, PAGE_READWRITE)
        For Index As Integer = 0 To GetDesktopIconCount() - 1
            Dim IconItem As TagLVItemW = New TagLVItemW With {.iSubItem = 0, .cchTextMax = MAX_PATH * 2, .pszText = CLng(PIconText)}
            Dim PIconItemStructure As IntPtr = Marshal.AllocCoTaskMem(GetLvItemSize())
            Marshal.StructureToPtr(IconItem, PIconItemStructure, False)
            Dim NumberOfBytesWritten As Integer
            Dim lpNumberOfBytesRead As Integer
            Dim Result As Integer = WriteProcessMemory(HandleProcess, PIconItem, PIconItemStructure, GetLvItemSize(), NumberOfBytesWritten)
            Marshal.FreeCoTaskMem(PIconItemStructure)
            SendMessage(DesktopHandle, LVM_GETITEMTEXT, Index, PIconItem)
            SendMessage(DesktopHandle, LVM_GETITEMPOSITION, Index, PIconPosition)
            Dim pszName As IntPtr = Marshal.AllocCoTaskMem(MAX_PATH * 2)
            ReadProcessMemory(HandleProcess, PIconText, pszName, MAX_PATH * 2, lpNumberOfBytesRead)
            If Marshal.PtrToStringUni(pszName) = IconText Then
                ReadProcessMemory(HandleProcess, PIconPosition, IconPosition(IconPosition.Length - 1), SizeOfPoint, lpNumberOfBytesRead)
                ReDim Preserve IconPosition(IconPosition.Length)
            End If

            Marshal.FreeCoTaskMem(pszName)
        Next
        VirtualFreeEx(HandleProcess, PIconText, MAX_PATH * 2, 0)
        VirtualFreeEx(HandleProcess, PIconPosition, SizeOfPoint, 0)
        VirtualFreeEx(HandleProcess, PIconItem, GetLvItemSize(), 0)
        ReDim Preserve IconPosition(IconPosition.Length - 1)
        Return IconPosition
    End Function

    ''' <summary>
    ''' 获取当前桌面图标总数
    ''' </summary>
    ''' <returns>桌面图标总数</returns>
    Public Function GetDesktopIconCount() As Integer
        Return SendMessage(DesktopHandle, LVM_GETITEMCOUNT, 0, 0)
    End Function

    ''' <summary>
    ''' 获取指定图标ID的图标文本和位置
    ''' </summary>
    ''' <param name="IconIndex">欲获取的图标ID</param>
    ''' <returns>储存结果的图标结构</returns>
    Public Function GetIconTextAndPosition(ByVal IconIndex As Integer) As DesktopIcon
        Dim Icon As DesktopIcon = New DesktopIcon
        Dim PIconItem As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, GetLvItemSize(), MEM_COMMIT, PAGE_READWRITE)
        Dim PIconText As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, MAX_PATH * 2, MEM_COMMIT, PAGE_READWRITE)
        Dim PIconPosition As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, SizeOfPoint, MEM_COMMIT, PAGE_READWRITE)
        Dim IconItem As TagLVItemW = New TagLVItemW With {.iSubItem = 0, .cchTextMax = MAX_PATH * 2, .pszText = CLng(PIconText)}
        Dim PIconItemStructure As IntPtr = Marshal.AllocCoTaskMem(GetLvItemSize())
        Marshal.StructureToPtr(IconItem, PIconItemStructure, False)
        Dim NumberOfBytesWritten As Integer
        Dim lpNumberOfBytesRead As Integer
        Dim Result As Integer = WriteProcessMemory(HandleProcess, PIconItem, PIconItemStructure, GetLvItemSize(), NumberOfBytesWritten)
        Marshal.FreeCoTaskMem(PIconItemStructure)
        SendMessage(DesktopHandle, LVM_GETITEMTEXT, IconIndex, PIconItem)
        SendMessage(DesktopHandle, LVM_GETITEMPOSITION, IconIndex, PIconPosition)
        ReadProcessMemory(HandleProcess, PIconPosition, Icon.Position, SizeOfPoint, lpNumberOfBytesRead)
        Dim pszName As IntPtr = Marshal.AllocCoTaskMem(MAX_PATH * 2)
        ReadProcessMemory(HandleProcess, PIconText, pszName, MAX_PATH * 2, lpNumberOfBytesRead)
        Icon.Text = Marshal.PtrToStringUni(pszName)
        Marshal.FreeCoTaskMem(pszName)
        VirtualFreeEx(HandleProcess, PIconText, MAX_PATH * 2, 0)
        VirtualFreeEx(HandleProcess, PIconPosition, SizeOfPoint, 0)
        VirtualFreeEx(HandleProcess, PIconItem, GetLvItemSize(), 0)
        'Debug.Print("图标：{0}{1}坐标：{2}，{3}", Icon.Text, vbCrLf, Icon.Position.X, Icon.Position.Y)
        Return Icon
    End Function

    ''' <summary>
    ''' 根据图标ID设置图标位置
    ''' </summary>
    ''' <param name="IconIndex">图标的ID</param>
    ''' <param name="IconPosition">欲设置的位置坐标</param>
    Public Sub SetIconPosition(ByVal IconIndex As Integer, ByVal IconPosition As Point)
        Dim PIconPosition As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, SizeOfPoint, MEM_COMMIT, PAGE_READWRITE)
        Dim PIconText As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, MAX_PATH * 2, MEM_COMMIT, PAGE_READWRITE)
        Dim PIconItem As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, GetLvItemSize(), MEM_COMMIT, PAGE_READWRITE)
        SendMessage(DesktopHandle, LVM_SETITEMPOSITION, IconIndex, IconPosition.X Or CInt(CUShort(IconPosition.Y)) << 16)
        VirtualFreeEx(HandleProcess, PIconPosition, SizeOfPoint, 0)
        VirtualFreeEx(HandleProcess, PIconText, MAX_PATH * 2, 0)
        VirtualFreeEx(HandleProcess, PIconItem, GetLvItemSize(), 0)
    End Sub

    ''' <summary>
    ''' 根据桌面图标文本设置图标位置
    ''' </summary>
    ''' <param name="IconText">欲设置位置的图标文本</param>
    ''' <param name="IconPosition">欲设置的位置坐标</param>
    ''' <param name="IsSingle">是否只处理找到的第一个图标</param>
    Public Sub SetIconPosition(ByVal IconText As String, ByVal IconPosition As Point， Optional ByVal IsSingle As Boolean = False)
        Dim PIconPosition As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, SizeOfPoint, MEM_COMMIT, PAGE_READWRITE)
        Dim PIconText As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, MAX_PATH * 2, MEM_COMMIT, PAGE_READWRITE)
        Dim PIconItem As IntPtr = VirtualAllocEx(HandleProcess, IntPtr.Zero, GetLvItemSize(), MEM_COMMIT, PAGE_READWRITE)
        For Index As Integer = 0 To GetDesktopIconCount() - 1
            If GetIconTextAndPosition(Index).Text = IconText Then
                SendMessage(DesktopHandle, LVM_SETITEMPOSITION, Index, IconPosition.X Or CInt(CUShort(IconPosition.Y)) << 16)
                If IsSingle Then Exit Sub
            End If
        Next
        VirtualFreeEx(HandleProcess, PIconPosition, SizeOfPoint, 0)
        VirtualFreeEx(HandleProcess, PIconText, MAX_PATH * 2, 0)
        VirtualFreeEx(HandleProcess, PIconItem, GetLvItemSize(), 0)
    End Sub

    ''' <summary>
    ''' 获取物理系统桌面图标的句柄
    ''' </summary>
    Public Function GetDesktopIconHandle() As IntPtr
        Dim HandleDesktop As Integer = GetDesktopWindow
        Dim HandleTop As Integer = 0
        Dim LastHandleTop As Integer = 0
        Dim HandleSHELLDLL_DefView As Integer = 0
        Dim HandleSysListView32 As Integer = 0
        '在WorkerW结构里搜索
        Do Until HandleSysListView32 > 0
            HandleTop = FindWindowEx(HandleDesktop, LastHandleTop, "WorkerW", vbNullString)
            HandleSHELLDLL_DefView = FindWindowEx(HandleTop, 0, "SHELLDLL_DefView", vbNullString)
            If HandleSHELLDLL_DefView > 0 Then HandleSysListView32 = FindWindowEx(HandleSHELLDLL_DefView, 0, "SysListView32", "FolderView")
            LastHandleTop = HandleTop
            If LastHandleTop = 0 Then Exit Do
        Loop
        '如果找到了，立即返回
        If HandleSysListView32 > 0 Then Return HandleSysListView32
        '未找到，则在Progman里搜索(用于兼容WinXP系统)
        Do Until HandleSysListView32 > 0
            HandleTop = FindWindowEx(HandleDesktop, LastHandleTop, "Progman", "Program Manager")
            HandleSHELLDLL_DefView = FindWindowEx(HandleTop, 0, "SHELLDLL_DefView", vbNullString)
            If HandleSHELLDLL_DefView > 0 Then HandleSysListView32 = FindWindowEx(HandleSHELLDLL_DefView, 0, "SysListView32", "FolderView")
            LastHandleTop = HandleTop
            If LastHandleTop = 0 Then Exit Do : Return 0
        Loop
        Return HandleSysListView32
    End Function
End Module
