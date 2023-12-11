Imports SAPbouiCOM.Framework

Public Class ClsCFL
    Public Const Formtype = "OCFL"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim RejHeader As SAPbouiCOM.DBDataSource
    Dim RejLine As SAPbouiCOM.DBDataSource
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objQCForm As SAPbouiCOM.Form
    Dim Row As Integer
    Dim objQCMatrix As SAPbouiCOM.Matrix
    Dim objDTable As SAPbouiCOM.DataTable
    Dim ColAlias As String = ""
    Dim objGrid As SAPbouiCOM.Grid
    Dim QCFUID As String
    Public Sub LoadScreen(ByRef QCFormUID As String, ByRef DocName As String)
        ' objForm = objAddOn.objUIXml.LoadScreenXML("OpenList1.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        'Try
        '    Dim cflForm As SAPbouiCOM.Form
        '    If objAddOn.objApplication.Forms.Count > 0 Then
        '        For frm As Integer = 0 To objAddOn.objApplication.Forms.Count - 1
        '            If objAddOn.objApplication.Forms.Item(frm).UniqueID = "OCFL" Then
        '                cflForm = objAddOn.objApplication.Forms.Item("OCFL")
        '                cflForm.Close()
        '                Exit For
        '            End If
        '        Next
        '    End If
        'Catch ex As Exception
        'End Try
        Try
            objAddOn.objUIXml.AddXML("OpenList1.xml")
            'objAddOn.objUIXml.AddXML("NewOpenList.xml")
            objForm = objAddOn.objApplication.Forms.Item("OCFL")
            objForm.Select()
            bModal = True
            'objMatrix = objForm.Items.Item("3").Specific
            QCFUID = QCFormUID
            objQCForm = objAddOn.objApplication.Forms.Item(QCFormUID) ' objAddOn.objApplication.Forms.GetForm("MIPLQC", QCFormUID)
            If DocName = "G" Then
                objQCForm.Items.Item("13").Specific.String = ""
                objQCForm.Items.Item("13B").Specific.String = ""
            ElseIf DocName = "P" Then
                objQCForm.Items.Item("15").Specific.String = ""
                objQCForm.Items.Item("15B").Specific.String = ""
                objQCForm.Items.Item("51").Specific.String = ""
                objQCForm.Items.Item("51B").Specific.String = ""
            ElseIf DocName = "T" Then
                objQCForm.Items.Item("23").Specific.String = ""
                objQCForm.Items.Item("23B").Specific.String = ""
            ElseIf DocName = "R" Then
                objQCForm.Items.Item("51").Specific.String = ""
                objQCForm.Items.Item("51B").Specific.String = ""
            End If
            objQCMatrix = objQCForm.Items.Item("20").Specific
            objQCForm.Items.Item("27").Specific.String = ""
            If objQCMatrix.VisualRowCount > 0 Then
                objQCMatrix.Clear()
            End If
            TranValue = {""}
            strSQL = objAddOn.objQC.GetListQuery(QCFormUID, DocName, "", "")
            LoadList(DocName, strSQL, False)
        Catch ex As Exception
        End Try

    End Sub

    Public Sub LoadList(ByVal DocName As String, ByVal StrQuery As String, ByVal search As Boolean, Optional ByVal trackentry As Boolean = False)
        Try
            'objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim objRS As SAPbobsCOM.Recordset
            objMatrix = objForm.Items.Item("3").Specific
            'objGrid = objForm.Items.Item("3").Specific
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If search = False And trackentry = False Then
                objAddOn.objApplication.StatusBar.SetText("Loading Entries Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            objForm.Freeze(True)
            If objForm.DataSources.DataTables.Count.Equals(0) Then
                objForm.DataSources.DataTables.Add("DT_List")
            Else
                objForm.DataSources.DataTables.Item("DT_List").Clear()
            End If
            objDTable = objForm.DataSources.DataTables.Item("DT_List")
            objDTable.Clear()
            objDTable.ExecuteQuery(StrQuery)
            If search = True Then
                objRS.DoQuery(StrQuery)
                If objRS.RecordCount = 0 Then
                    Exit Sub
                End If
            End If

            'For i As Integer = 0 To objGrid.Columns.Count - 1
            '    objGrid.Columns.Item(i).TitleObject.Sortable = True
            '    objGrid.Columns.Item(i).Editable = False
            'Next
            'If objGrid.Rows.Count > 0 Then
            '    objGrid.Rows.SelectedRows.Add(0)
            '    objGrid.Columns.Item(0).TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
            '    objAddOn.objApplication.StatusBar.SetText("Successfully Loaded Entries...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'Else
            '    objAddOn.objApplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'End If
            'objGrid.AutoResizeColumns()
            objMatrix.Clear()
            objMatrix.LoadFromDataSourceEx()
            If trackentry = True Then
                Exit Sub
            End If
            If objMatrix.VisualRowCount > 0 Then
                'objMatrix.Columns.Item(0).TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                objMatrix.SelectRow(1, True, False)
                'objMatrix.CommonSetting.EnableArrowKey = True
                If search = False Then
                    objAddOn.objApplication.StatusBar.SetText("Successfully Loaded Open Entries...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            Else
                objAddOn.objApplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            objForm.Settings.Enabled = True
            'objDTable = Nothing
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objForm.Freeze(False)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Dim RowID As Integer
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("3").Specific
            'objGrid = objForm.Items.Item("3").Specific
            If pVal.BeforeAction Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If FormUID = "OCFL" And bModal Then
                            bModal = False
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If pVal.ItemUID = "11" Then

                            RowID = objMatrix.GetNextSelectedRow(pVal.Row, SAPbouiCOM.BoOrderType.ot_SelectionOrder)
                            If Not objMatrix.IsRowSelected(RowID) = True Then
                                Exit Sub
                            End If
                            TranValue = GetSelectedEntry(QCFUID, RowID)
                            objForm.Close()
                        End If
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        objMatrix.AutoResizeColumns()

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "3" And pVal.ActionSuccess = True Then
                            If pVal.Row <> 0 Then
                                objMatrix.SelectRow(pVal.Row, True, False)
                            Else
                                Exit Sub
                            End If
                        ElseIf pVal.ItemUID = "11" And pVal.ActionSuccess = True Then
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.ItemUID = "3" And pVal.ActionSuccess = True Then
                            ColAlias = GetDTAliasName(pVal.ColUID)
                            If pVal.Row <> 0 Then
                                If Not objMatrix.IsRowSelected(pVal.Row) = True Then
                                    Exit Sub
                                End If
                                RowID = objMatrix.GetNextSelectedRow(pVal.Row, SAPbouiCOM.BoOrderType.ot_SelectionOrder)
                                RowID = IIf(RowID = -1, pVal.Row, RowID)
                                TranValue = GetSelectedEntry(QCFUID, RowID)
                                objForm.Close()
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.ItemUID = "3" And pVal.ActionSuccess = True Then
                            If pVal.Row <> 0 Then
                                objMatrix.SelectRow(pVal.Row, True, False)
                                'objGrid.Rows.SelectedRows.Add(pVal.Row)
                            End If
                        ElseIf pVal.ItemUID = "5" And pVal.ActionSuccess = True Then
                            Try
                                Dim QCHeader As SAPbouiCOM.DBDataSource = objQCForm.DataSources.DBDataSources.Item("@MIPLQC")
                                Dim Type As String = QCHeader.GetValue("U_Type", 0)
                                Dim FindString As String
                                FindString = objForm.Items.Item("5").Specific.String
                                If FindString = "" Then
                                    strSQL = objAddOn.objQC.GetListQuery(QCFUID, Type, "", "")
                                    LoadList(Type, strSQL, False, True)
                                    objMatrix.SelectRow(1, True, False)
                                    'objGrid.Rows.SelectedRows.Add(0)
                                    Exit Sub
                                Else
                                    If ColAlias = "" Then
                                        ColAlias = "DocNum"
                                    End If
                                    strSQL = objAddOn.objQC.GetListQuery(QCFUID, Type, FindString.ToUpper, ColAlias)
                                    LoadList(Type, strSQL, True, True)
                                    objMatrix.SelectRow(1, True, False)
                                End If
                                'For i As Integer = 0 To objDTable.Rows.Count - 1
                                '    If objDTable.GetValue(ColAlias, i).ToString Like FindString & "*" Then
                                '        Colcal = objDTable.GetValue(ColAlias, i).ToString
                                '        Val = objGrid.GetDataTableRowIndex(i)
                                '        'Val = objDTable.GetValue("LineId", i).ToString
                                '        objMatrix.SelectRow(Val, True, False)
                                '        ' Val = objGrid.GetDataTableRowIndex(Val)
                                '        objGrid.Rows.SelectedRows.Add(Val)
                                '        Exit For
                                '    End If
                                'Next

                            Catch ex As Exception
                            End Try
                        End If

                        'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        '    If pVal.ItemUID = "5" And pVal.ActionSuccess = True Then
                        '    End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function GetDTAliasName(ByVal UID As String)
        Try
            Dim AliasName As String = ""
            If UID = "0" Then
                AliasName = "DocNum"
            ElseIf UID = "1" Then
                AliasName = "DocNum"
            ElseIf UID = "2" Then
                AliasName = "DocEntry"
            ElseIf UID = "3" Then
                AliasName = "CardCode"
            ElseIf UID = "3A" Then
                AliasName = "ItemCode"
            ElseIf UID = "3AA" Then
                AliasName = "Dscription"
            ElseIf UID = "4" Then
                AliasName = "DocDate"
            ElseIf UID = "5" Then
                AliasName = "DocDueDate"
            ElseIf UID = "6" Then
                AliasName = "BPLId"
            ElseIf UID = "7" Then
                AliasName = "LocCode"
            ElseIf UID = "8" Then
                AliasName = "Remarks"
            ElseIf UID = "8A" Then
                AliasName = "PendQty"
            End If
            Return AliasName
        Catch ex As Exception
            Return "DocNum"
        End Try
    End Function

    Private Function GetSelectedEntry(ByVal QCFUID As String, Row As Integer)
        Try
            objMatrix = objForm.Items.Item("3").Specific
            Dim ObjQCForm As SAPbouiCOM.Form
            Dim DocEntry As String = "", DocNum As String = ""
            ObjQCForm = objAddOn.objApplication.Forms.Item(QCFUID) 'objAddOn.objApplication.Forms.GetForm("MIPLQC", 0)
            Dim QCHeader As SAPbouiCOM.DBDataSource = ObjQCForm.DataSources.DBDataSources.Item("@MIPLQC")
            Dim Type As String = Trim(QCHeader.GetValue("U_Type", 0))
            Try
                DocEntry = objMatrix.Columns.Item("2").Cells.Item(Row).Specific.String
                DocNum = objMatrix.Columns.Item("1").Cells.Item(Row).Specific.String
                'For i As Integer = 1 To objMatrix.VisualRowCount
                '    If objMatrix.IsRowSelected(i) Then
                '        DocEntry = objMatrix.Columns.Item("2").Cells.Item(i).Specific.String
                '        DocNum = objMatrix.Columns.Item("1").Cells.Item(i).Specific.String
                '        Exit For
                '    End If
                'Next
                If Type = "G" Then
                    ObjQCForm.Items.Item("13").Specific.String = DocNum
                    ObjQCForm.Items.Item("13B").Specific.String = DocEntry
                ElseIf Type = "P" Then
                    If ObjQCForm.Items.Item("15").Specific.String = "" Then
                        ObjQCForm.Items.Item("15").Specific.String = DocNum
                        ObjQCForm.Items.Item("15B").Specific.String = DocEntry
                    Else
                        ObjQCForm.Items.Item("51").Specific.String = DocNum
                        ObjQCForm.Items.Item("51B").Specific.String = DocEntry
                    End If
                ElseIf Type = "T" Then
                    ObjQCForm.Items.Item("23").Specific.String = DocNum
                    ObjQCForm.Items.Item("23B").Specific.String = DocEntry
                ElseIf Type = "R" Then
                    ObjQCForm.Items.Item("51").Specific.String = DocNum
                    ObjQCForm.Items.Item("51B").Specific.String = DocEntry
                End If
                Return {DocNum, DocEntry}
            Catch ex As Exception
                Return ""
            End Try
        Catch ex As Exception

        End Try
    End Function
End Class
