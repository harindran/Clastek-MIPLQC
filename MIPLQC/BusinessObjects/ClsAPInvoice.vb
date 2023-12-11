Imports System.IO
Imports SAPbouiCOM.Framework
Public Class ClsAPInvoice
    Public Const Formtype = "141"
    Dim objAPform As SAPbouiCOM.Form
    Dim ObjQCForm As SAPbouiCOM.Form
    Dim ObjGRPOForm As SAPbouiCOM.Form

    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        objAPform = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "BtnVal" Then
                        objAPform.Items.Item("txtGet").Specific.String = GRPOEntries
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    ObjGRPOForm = objAddOn.objApplication.Forms.GetForm("143", 1)
                    Dim DocEntry As String = ObjGRPOForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0)
                    GRPOEntries = DocEntry
            End Select
        Else
            Try
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        If pVal.ActionSuccess Then
                            CreateButton(FormUID)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "BtnVal" And objAPform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            CreateAPDeviation("QCR", "QC report", "20")
                            'If objAPform.Items.Item("txtGet").Specific.String <> "" Then
                            '    CreateAPDeviation("QCR", "QC report", "20")
                            'Else
                            '    objAddOn.objApplication.StatusBar.SetText("Unable to run the QC report due to GRPO entries not tracked...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            'End If

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        'If pVal.ItemUID = "24" Then
                        '    objAPform.Items.Item("txtGet").Specific.String = GRPOEntries
                        '    'If objAPform.Items.Item("24").Specific.String = "" Then
                        '    '    objAPform.Items.Item("txtGet").Specific.String = GRPOEntries
                        '    'End If
                        '    'Dim FinalData As String = Microsoft.VisualBasic.InputBox("Enter the alphabet- g", "Enter the Character", "g", 50, 50)
                        '    'If FinalData = "g" Then
                        '    '    objAPform.Items.Item("txtGet").Specific.String = GRPOEntries
                        '    'End If
                        'End If
                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objAPform = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                End Select
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        'Dim objedit As SAPbouiCOM.EditText
                        'objedit = objAPform.Items.Item("txtGet").Specific
                        'objedit.Item.Enabled = False
                End Select
            End If

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Sub CreateButton(ByVal FormUID As String)
        Try
            Dim objButton As SAPbouiCOM.Button
            Dim objItem As SAPbouiCOM.Item
            objAPform = objAddOn.objApplication.Forms.Item(FormUID)
            objItem = objAPform.Items.Add("BtnVal", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            objItem.Left = objAPform.Items.Item("2").Left + objAPform.Items.Item("2").Width + 10
            'objItem.Left = objForm.Items.Item("10002056").Left + objForm.Items.Item("10002056").Width + 60
            objItem.Width = 120
            objItem.Top = objAPform.Items.Item("2").Top
            objItem.Height = objAPform.Items.Item("2").Height
            objButton = objItem.Specific
            objButton.Caption = "QC Validation"


            Dim objedit As SAPbouiCOM.EditText
            objItem = objAPform.Items.Add("txtGet", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Left = objAPform.Items.Item("BtnVal").Left + objAPform.Items.Item("BtnVal").Width + 10
            objItem.Width = 50
            objItem.Top = objAPform.Items.Item("BtnVal").Top
            objItem.Height = objAPform.Items.Item("BtnVal").Height
            objItem.LinkTo = "BtnVal"
            objedit = objItem.Specific
            objedit.Item.Enabled = False
            objedit.DataBind.SetBound(True, "OPCH", "U_GetNum")
            'objAddOn.objApplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
        End Try

    End Sub

    Private Sub CreateAPDeviation(ByVal FormID As String, ByVal FormTitle As String, ByVal LinkedID As String)
        Dim oCreationParams As SAPbouiCOM.FormCreationParams
        Dim objTempForm As SAPbouiCOM.Form
        Try

            objAddOn.objApplication.Forms.Item(FormID).Visible = True
        Catch ex As Exception
            Dim str_sql As String = ""
            'MsgBox(GRPOEntries)
            If objAddOn.HANA Then
                str_sql = "SELECT T5.""DocEntry"" as ""GRN Entry"",T5.""DocNum"" as ""GRN Num"",T2.""ItemCode"",T2.""Dscription"" as ""Item Description"", T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") AS ""GRN Qty"" ,"
                str_sql += vbCrLf + "(SELECT IFNULL(SUM(T1.""U_AccQty""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON  T0.""DocEntry""=T1.""DocEntry"" "
                str_sql += vbCrLf + "AND T0.""U_GRNEntry"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""Accepted Qty"", "
                str_sql += vbCrLf + "(SELECT IFNULL(SUM(T1.""U_RejQty""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON  T0.""DocEntry""=T1.""DocEntry"" "
                str_sql += vbCrLf + "AND T0.""U_GRNEntry"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""Rejected Qty"" "
                str_sql += vbCrLf + " FROM OPDN T5 join PDN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" "
                str_sql += vbCrLf + "WHERE T2.""DocEntry"" in (" & GRPOEntries & ") AND IFNULL(T3.""U_InspReq"",'') = 'Y' "
                str_sql += vbCrLf + "GROUP BY  T2.""ItemCode"",T2.""Dscription"",T4.""BaseQty"",T4.""AltQty"",T5.""DocNum"",T5.""DocEntry"";"
            Else
                str_sql = "SELECT T5.DocEntry as GRNEntry,T5.DocNum as GRNNum,T2.ItemCode,T2.Dscription as ItemDescription, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) AS GRNQty ,"
                str_sql += vbCrLf + "(SELECT ISNULL(SUM(T1.U_AccQty), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON  T0.DocEntry=T1.DocEntry "
                str_sql += vbCrLf + "AND T0.U_GRNEntry = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS AcceptedQty, "
                str_sql += vbCrLf + "(SELECT ISNULL(SUM(T1.U_RejQty), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON  T0.DocEntry=T1.DocEntry "
                str_sql += vbCrLf + "AND T0.U_GRNEntry = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS RejectedQty "
                str_sql += vbCrLf + " FROM OPDN T5 join PDN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry "
                str_sql += vbCrLf + "WHERE T2.DocEntry in (" & GRPOEntries & ") AND ISNULL(T3.U_InspReq,'') = 'Y' "
                str_sql += vbCrLf + "GROUP BY T2.DocEntry,T2.LineNum,T2.ItemCode,T2.Dscription,T4.BaseQty,T4.AltQty,T5.DocNum,T5.DocEntry"
            End If
          

            Dim objrs As SAPbobsCOM.Recordset
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(str_sql)
            If objrs.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objrs = Nothing : Exit Sub

            oCreationParams = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            oCreationParams.UniqueID = FormID
            objTempForm = objAddOn.objApplication.Forms.AddEx(oCreationParams)
            objTempForm.Title = FormTitle
            objTempForm.Left = 400
            objTempForm.Top = 100
            objTempForm.ClientHeight = 300 '335
            objTempForm.ClientWidth = 700
            objTempForm = objAddOn.objApplication.Forms.Item(FormID)
            Dim oitm As SAPbouiCOM.Item

            Dim oGrid As SAPbouiCOM.Grid
            oitm = objTempForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oitm.Top = 30
            oitm.Left = 2
            oitm.Width = 700
            oitm.Height = 200
            oGrid = objTempForm.Items.Item("Grid").Specific
            objTempForm.DataSources.DataTables.Add("DataTable")
            oitm = objTempForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitm.Top = objTempForm.Items.Item("Grid").Top + objTempForm.Items.Item("Grid").Height + 5
            oitm.Left = 10

            Dim objDT As SAPbouiCOM.DataTable
            objDT = objTempForm.DataSources.DataTables.Item("DataTable")
            objDT.ExecuteQuery(str_sql)

            objTempForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_sql)

            oGrid.DataTable = objTempForm.DataSources.DataTables.Item("DataTable")

            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).TitleObject.Sortable = True
                oGrid.Columns.Item(i).Editable = False
            Next

            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            Dim col As SAPbouiCOM.EditTextColumn
            col = oGrid.Columns.Item(0)
            col.LinkedObjectType = LinkedID
            objTempForm.Visible = True
            oGrid.AutoResizeColumns()
            objTempForm.Update()

        End Try
    End Sub

End Class
