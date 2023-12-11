Imports System.IO
Imports SAPbouiCOM.Framework
Public Class clsRejDetails
    Public Const Formtype = "MIREJDET"
    Dim objRejForm As SAPbouiCOM.Form
    Dim strSQL As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim RejHeader As SAPbouiCOM.DBDataSource
    Dim RejLine As SAPbouiCOM.DBDataSource
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objQCForm As SAPbouiCOM.Form
    Dim Row As Integer
    Dim objRejMatrix As SAPbouiCOM.Matrix
    Dim AddedRejNumber As String
    Dim strRejDetails As String = ""
    Dim objQCMatrix As SAPbouiCOM.Matrix

    Public Sub LoadScreen(ByRef QCFormUID As String, ByVal RowID As Integer, Optional RejEntry As String = "")
        objRejForm = objAddOn.objUIXml.LoadScreenXML("RejDetails.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
        objRejMatrix = objRejForm.Items.Item("13").Specific
        If RejEntry <> "" Then
            objRejForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objRejForm.Items.Item("4").Enabled = True
            objRejForm.Items.Item("4").Specific.String = RejEntry
            objRejForm.ActiveItem = "6"
            objRejForm.Items.Item("4").Enabled = False
            objRejForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Else
            objQCForm = objAddOn.objApplication.Forms.Item(QCFormUID)
            Row = RowID
            RejHeader = objRejForm.DataSources.DBDataSources.Item("@MIREJDET")
            RejLine = objRejForm.DataSources.DBDataSources.Item("@MIREJDET1")
            'objForm.Items.Item("21").Specific.validvalues.loadseries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
            'objForm.Items.Item("21").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            If objAddOn.HANA Then
                RejHeader.SetValue("DocEntry", 0, objAddOn.objGenFunc.GetNextDocEntryValue("@MIREJDET"))
            Else
                RejHeader.SetValue("DocEntry", 0, objAddOn.objGenFunc.GetNextDocEntryValue("[@MIREJDET]"))
            End If
            Matrix_Addrow(objRejMatrix, "3", "0")
            objRejForm.Items.Item("6").Specific.Active = True
            objRejForm.Items.Item("6").Specific.String = "A"
            'objRejForm.Items.Item("13").Specific.addrow()
        End If
        objQCMatrix = objQCForm.Items.Item("20").Specific
        RejHeader.SetValue("U_ItemCode", 0, objQCMatrix.Columns.Item("1").Cells.Item(RowID).Specific.string)
        RejHeader.SetValue("U_ItemName", 0, objQCMatrix.Columns.Item("2").Cells.Item(RowID).Specific.string)
        If objRejForm.Items.Item("12").Enabled = False Then
            objRejForm.Items.Item("12").Enabled = True
            RejHeader.SetValue("U_RejQty", 0, objQCMatrix.Columns.Item("7").Cells.Item(RowID).Specific.string)
            objRejForm.Items.Item("12").Enabled = False
        End If
        RejHeader.SetValue("U_RejQty", 0, objQCMatrix.Columns.Item("7").Cells.Item(RowID).Specific.string)
        objRejMatrix.Columns.Item("2").Cells.Item(1).Specific.String = objQCMatrix.Columns.Item("7").Cells.Item(RowID).Specific.string

        objRejMatrix.Columns.Item("2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

    End Sub

    Sub DeleteEmptyRowInFormDataEvent(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal ColumnUID As String, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            If oMatrix.VisualRowCount > 0 Then
                If oMatrix.Columns.Item(ColumnUID).Cells.Item(oMatrix.VisualRowCount).Specific.Value.Equals("") Then
                    oMatrix.DeleteRow(oMatrix.VisualRowCount)
                    oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
                    oMatrix.FlushToDataSource()
                End If
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Delete Empty RowIn Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Sub Matrix_Addrow(ByVal omatrix As SAPbouiCOM.Matrix, Optional ByVal colname As String = "", Optional ByVal rowno_name As String = "", Optional ByVal Error_Needed As Boolean = False)
        Try
            Dim addrow As Boolean = False

            If omatrix.VisualRowCount = 0 Then addrow = True : GoTo addrow
            If colname = "" Then addrow = True : GoTo addrow
            If omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific.string <> "" Then addrow = True : GoTo addrow

addrow:
            If addrow = True Then
                omatrix.AddRow(1)
                omatrix.ClearRowData(omatrix.VisualRowCount)
                If rowno_name <> "" Then omatrix.Columns.Item("0").Cells.Item(omatrix.VisualRowCount).Specific.string = omatrix.VisualRowCount
            Else
                If Error_Needed = True Then objAddOn.objApplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        objRejForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And (objRejForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objRejForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If validate(FormUID) = False Then
                            System.Media.SystemSounds.Asterisk.Play()
                            BubbleEvent = False
                            Exit Sub
                        Else
                            strRejDetails = RejHeader.GetValue("DocEntry", 0)
                        End If

                        'If Not validate(FormUID) Then
                        '    ' objAddOn.objApplication.SetStatusBarMessage("Validation Failed") -- SP is set
                        '    BubbleEvent = False
                        '    Return
                        'Else
                        '    strRejDetails = RejHeader.GetValue("DocEntry", 0)
                        'End If
                    End If

            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And objRejForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                        If objQCMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.String <> "" Then
                            objRejForm.Close()
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And objRejForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                        'objQCMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.String = strRejDetails
                        'objQCMatrix.CommonSetting.SetCellEditable(Row, 17, False)
                        'If objQCMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.String <> "" Then
                        '    objRejForm.Close()
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "13" And pVal.ColUID = "3" Then
                        objMatrix = objRejForm.Items.Item("13").Specific
                        objRejMatrix.Columns.Item("2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                        If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String <> "" Then
                            'objMatrix.AddRow()
                            ' SetNewLine(objMatrix, RejLine, pVal.Row, "3")
                            Matrix_Addrow(objRejMatrix, "3", "0")
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "13" And (pVal.ColUID = "3") Then
                        CFL(FormUID, pVal)
                    End If
            End Select
        End If
    End Sub

    Private Sub CFL(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Try
            Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
            Dim objDataTable As SAPbouiCOM.DataTable
            objCFLEvent = pval
            objDataTable = objCFLEvent.SelectedObjects
            objRejForm = objAddOn.objApplication.Forms.Item(FormUID)
            objRejMatrix = objRejForm.Items.Item("13").Specific
            Select Case objCFLEvent.ChooseFromListUID
                Case "CFL_REJ"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objRejMatrix.Columns.Item("3").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("Name", 0)
                        End If
                    Catch ex As Exception
                        objRejMatrix.Columns.Item("3").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("Name", 0)
                    End Try
            End Select

        Catch ex As Exception
            'MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
        End Try

    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objRejForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
        objRejMatrix = objRejForm.Items.Item("13").Specific
        RejHeader = objRejForm.DataSources.DBDataSources.Item("@MIREJDET")
        RejLine = objRejForm.DataSources.DBDataSources.Item("@MIREJDET1")
        If BusinessObjectInfo.BeforeAction Then
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    ' If BusinessObjectInfo.BeforeAction = True Then
                    If validate(objRejForm.UniqueID) = False Then
                        System.Media.SystemSounds.Asterisk.Play()
                        BubbleEvent = False
                        Exit Sub
                    Else
                        DeleteEmptyRowInFormDataEvent(objRejMatrix, "3", RejLine)
                        objQCMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.String = strRejDetails
                        objQCMatrix.CommonSetting.SetCellEditable(Row, 19, False)
                        'If objQCMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.String <> "" Then
                        '    objRejForm.Close()
                        'Else
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'strRejDetails = RejHeader.GetValue("DocEntry", 0)
                        'objQCMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.String = strRejDetails
                    End If
                    'End If
            End Select
        End If

    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.MenuUID = "ditem" Then
            objRejForm = objAddOn.objApplication.Forms.ActiveForm
            objRejMatrix = objRejForm.Items.Item("13").Specific
            If objAddOn.ZB_row > 0 Then
                objRejMatrix.DeleteRow(objAddOn.ZB_row)
            End If
        End If
    End Sub

    Private Function validate(ByVal FormUID As String) As Boolean
        objRejForm = objAddOn.objApplication.Forms.Item(FormUID)
        'objRejForm = objAddOn.objApplication.Forms.ActiveForm
        objRejMatrix = objRejForm.Items.Item("13").Specific
        RejHeader = objRejForm.DataSources.DBDataSources.Item("@MIREJDET")
        If CDbl(Trim(RejHeader.GetValue("U_RejQty", 0))) <> CDbl(Trim(objRejMatrix.Columns.Item("2").ColumnSetting.SumValue)) Then
            'objAddOn.objApplication.SetStatusBarMessage("Please check the Quantity!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'MessageBox.Show("Please check the Quantity!!!", "Error")
            objAddOn.objApplication.MessageBox("Please check the Quantity!!!",, "OK")
            Return False
        End If
        For i As Integer = 1 To objRejMatrix.RowCount
            If objRejMatrix.Columns.Item("2").Cells.Item(i).Specific.string <> "" Then
                If CDbl(objRejMatrix.Columns.Item("2").Cells.Item(i).Specific.string) > 0 And objRejMatrix.Columns.Item("3").Cells.Item(i).Specific.string.trim = "" Then
                    ' objAddOn.objApplication.SetStatusBarMessage("Please Update the Rejected Reason!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    ' MessageBox.Show("Please Update the Rejected Reason!!!", "Error")
                    objAddOn.objApplication.MessageBox("Please Update the Rejected Reason!!!",, "OK")
                    Return False
                End If
            End If

        Next
        Return True
    End Function
End Class
