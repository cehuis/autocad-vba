Attribute VB_Name = "bas_CassDatFilter"
Public Sub vba_zzDcx()
frm_CassDatFilter.show
End Sub

Sub CreateMenu()
'�����˵���
Dim mnuGroup As AcadMenuGroup
Set mnuGroup = ThisDrawing.Application.MenuGroups.Item(0)

'�����²˵�
Dim mnuQinDong As AcadPopupMenu
Set mnuQinDong = mnuGroup.Menus.Add("����������(&T)")

'���������˵���ִ���Ա��VBA������ϡ����vba_zzDCX
Dim mnuDCX As AcadPopupMenuItem
Dim macDCX As String
macDCX = Chr(3) & Chr(3) & Chr(95) & "-vbarun" & Chr(32) & "vba_zzDCX" & Chr(32)
Set mnuDCX = mnuQinDong.AddMenuItem(mnuQinDong.Count + 1, "���ε����(&G)", macDCX)

'�����ָ���
Dim mnuSeparator As AcadPopupMenuItem
Set mnuSeparator = mnuQinDong.AddSeparator("")

'���������˵���ִ��AutoCAD�ڲ�����
'Dim mnuCopy As AcadPopupMenuItem
'Dim macCopy As String
'macCopy = Chr(3) & Chr(3) & Chr(95) & "copy" & Chr(32)
'Set mnuCopy = mnuQinDong.AddMenuItem(mnuQinDong.Count + 1, "&Copy", macCopy)

'�����Ӳ˵�
'Dim mnuFather As AcadPopupMenu
'Set mnuFather = mnuQinDong.AddSubMenu(mnuQinDong.Count + 1, "���˵�")
'Dim mnuChild As AcadPopupMenuItem
'Dim macChild As String
'macChild = Chr(3) & Chr(3) & Chr(95) & "export" & Chr(32)
'Set mnuChild = mnuFather.AddMenuItem(mnuQinDong.Count + 1, "�Ӳ˵�-����������ʽ", macChild)

'�ڲ˵�������ʾ�˵�
mnuQinDong.InsertInMenuBar ThisDrawing.Application.MenuBar.Count + 1

'ɾ���˵�
'If MsgBox("�Ƿ�ɾ�� COPY �˵�?", vbYesNo, "AutoCAD��ʾ") = vbYes Then
'mnuCopy.Delete
'End If
End Sub

'Public Sub AcadStartUp()
'Call CreateToolbarExample
'End Sub
'
''��ӹ�����
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
'Set tbTest = mnuGroup.Toolbars.Add("��ϡ")
'macCopy = Chr(3) & Chr(3) & Chr(95) & "zzDCX" & Chr(32)
'macPaste = Chr(3) & Chr(3) & Chr(95) & "pasteclip" & Chr(32)
'Set tbCopy = tbTest.AddToolbarButton _
'(tbTest.Count + 1, "����", "����", macCopy, False)
'Set tbPaste = tbTest.AddToolbarButton _
'(tbTest.Count + 1, "ճ�� ", "ճ��", macPaste, False)
'Set tbSeparator = tbTest.AddSeparator(tbTest.Count + 1)
'strPath1 = "f:\4.bmp"
'strPath2 = "f:\4.bmp"
'tbCopy.SetBitmaps strPath1, strPath2
''strPath1 = "G:\VBA\paste.bmp"
''strPath2 = "G:\VBA\paste.bmp"
''tbPaste.SetBitmaps strPath1, strPath2
''MsgBox "��"
'tbTest.Dock acToolbarDockLeft
''MsgBox "��"
''tbTest.Float 550, 300, 1
'End Sub

