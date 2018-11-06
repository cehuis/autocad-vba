Attribute VB_Name = "bas_CassDatFilter"
Public Sub vba_zzDcx()
frm_CassDatFilter.show
End Sub

Sub CreateMenu()
'创建菜单组
Dim mnuGroup As AcadMenuGroup
Set mnuGroup = ThisDrawing.Application.MenuGroups.Item(0)

'创建新菜单
Dim mnuQinDong As AcadPopupMenu
Set mnuQinDong = mnuGroup.Menus.Add("测量工具箱(&T)")

'创建下拉菜单，执行自编的VBA程序点抽稀过滤vba_zzDCX
Dim mnuDCX As AcadPopupMenuItem
Dim macDCX As String
macDCX = Chr(3) & Chr(3) & Chr(95) & "-vbarun" & Chr(32) & "vba_zzDCX" & Chr(32)
Set mnuDCX = mnuQinDong.AddMenuItem(mnuQinDong.Count + 1, "地形点过滤(&G)", macDCX)

'创建分隔线
Dim mnuSeparator As AcadPopupMenuItem
Set mnuSeparator = mnuQinDong.AddSeparator("")

'创建下拉菜单，执行AutoCAD内部命令
'Dim mnuCopy As AcadPopupMenuItem
'Dim macCopy As String
'macCopy = Chr(3) & Chr(3) & Chr(95) & "copy" & Chr(32)
'Set mnuCopy = mnuQinDong.AddMenuItem(mnuQinDong.Count + 1, "&Copy", macCopy)

'创建子菜单
'Dim mnuFather As AcadPopupMenu
'Set mnuFather = mnuQinDong.AddSubMenu(mnuQinDong.Count + 1, "父菜单")
'Dim mnuChild As AcadPopupMenuItem
'Dim macChild As String
'macChild = Chr(3) & Chr(3) & Chr(95) & "export" & Chr(32)
'Set mnuChild = mnuFather.AddMenuItem(mnuQinDong.Count + 1, "子菜单-导出其它格式", macChild)

'在菜单条上显示菜单
mnuQinDong.InsertInMenuBar ThisDrawing.Application.MenuBar.Count + 1

'删除菜单
'If MsgBox("是否删除 COPY 菜单?", vbYesNo, "AutoCAD提示") = vbYes Then
'mnuCopy.Delete
'End If
End Sub

'Public Sub AcadStartUp()
'Call CreateToolbarExample
'End Sub
'
''添加工具栏
'Public Sub CreateToolbarExample()
'Dim mnuGroup As AcadMenuGroup
'Dim tbTest As AcadToolbar
'Dim tbCopy As AcadToolbarItem
'Dim tbPaste As AcadToolbarItem
'Dim tbSeparator As AcadToolbarItem
'Dim macCopy As String
'Dim macPasteclip As String
'Dim strPath1 As String
'Dim strPath2 As String
'Set mnuGroup = ThisDrawing.Application.MenuGroups.Item(0)
'Set tbTest = mnuGroup.Toolbars.Add("抽稀")
'macCopy = Chr(3) & Chr(3) & Chr(95) & "zzDCX" & Chr(32)
'macPaste = Chr(3) & Chr(3) & Chr(95) & "pasteclip" & Chr(32)
'Set tbCopy = tbTest.AddToolbarButton _
'(tbTest.Count + 1, "复制", "复制", macCopy, False)
'Set tbPaste = tbTest.AddToolbarButton _
'(tbTest.Count + 1, "粘贴 ", "粘贴", macPaste, False)
'Set tbSeparator = tbTest.AddSeparator(tbTest.Count + 1)
'strPath1 = "f:\4.bmp"
'strPath2 = "f:\4.bmp"
'tbCopy.SetBitmaps strPath1, strPath2
''strPath1 = "G:\VBA\paste.bmp"
''strPath2 = "G:\VBA\paste.bmp"
''tbPaste.SetBitmaps strPath1, strPath2
''MsgBox "左"
'tbTest.Dock acToolbarDockLeft
''MsgBox "右"
''tbTest.Float 550, 300, 1
'End Sub

