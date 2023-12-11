Imports System.IO
Imports SAPbouiCOM.Framework
Public Class ClsCFLGRPO
    Dim objCFLform As SAPbouiCOM.Form
    Dim objAPform As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objGetVal, GetDate As SAPbouiCOM.EditText
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        objCFLform = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objCFLform.Items.Item("7").Specific
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
            End Select
        Else
            Dim GetDocEntry As String = ""
            Dim DocDate As Date
            Try
                If objAddOn.objApplication.Forms.ActiveForm.TypeEx = "MIPLQC" Then
                    Exit Sub
                End If
                objAPform = objAddOn.objApplication.Forms.GetForm("141", 0)
                objGetVal = objAPform.Items.Item("txtGet").Specific
                objGetVal.Value = ""
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.ItemUID = "1") And pVal.ActionSuccess Then
                            'If GRPOEntries = "" Then
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                GetDate = objMatrix.Columns.Item("DocDate").Cells.Item(i).Specific
                                If objMatrix.IsRowSelected(i) Then
                                    DocDate = Date.ParseExact(GetDate.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                                    GetDocEntry = objAddOn.objGenFunc.getSingleValue("Select ""DocEntry"" from OPDN where ""DocNum""= " & objMatrix.Columns.Item("DocNum").Cells.Item(i).Specific.String & " and ""DocDate""='" & DocDate.ToString("yyyyMMdd") & "' ")
                                    If GRPOEntries = "" Then
                                        GRPOEntries = GetDocEntry
                                    Else
                                        GRPOEntries += "," + GetDocEntry
                                    End If
                                End If
                            Next
                            'End If
                        ElseIf (pVal.ItemUID = "7" And pVal.Modifiers <> SAPbouiCOM.BoModifiersEnum.mt_CTRL) And pVal.ActionSuccess Then
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                GetDate = objMatrix.Columns.Item("DocDate").Cells.Item(i).Specific
                                If objMatrix.IsRowSelected(i) Then
                                    DocDate = Date.ParseExact(GetDate.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                                    GetDocEntry = objAddOn.objGenFunc.getSingleValue("Select ""DocEntry"" from OPDN where ""DocNum""= " & objMatrix.Columns.Item("DocNum").Cells.Item(i).Specific.String & " and ""DocDate""='" & DocDate.ToString("yyyyMMdd") & "' ")
                                    If GRPOEntries = "" Then
                                        GRPOEntries = GetDocEntry
                                    Else
                                        GRPOEntries += "," + GetDocEntry
                                    End If
                                End If
                            Next
                        End If
                        objGetVal.Value = GRPOEntries
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If pVal.CharPressed = "13" And pVal.ActionSuccess Then
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                GetDate = objMatrix.Columns.Item("DocDate").Cells.Item(i).Specific
                                If objMatrix.IsRowSelected(i) Then
                                    DocDate = Date.ParseExact(GetDate.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                                    GetDocEntry = objAddOn.objGenFunc.getSingleValue("Select ""DocEntry"" from OPDN where ""DocNum""= " & objMatrix.Columns.Item("DocNum").Cells.Item(i).Specific.String & " and ""DocDate""='" & DocDate.ToString("yyyyMMdd") & "' ")
                                    If GRPOEntries = "" Then
                                        GRPOEntries = GetDocEntry
                                    Else
                                        GRPOEntries += "," + GetDocEntry
                                    End If
                                End If
                            Next
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        objGetVal.Value = ""
                        GRPOEntries = ""
                End Select
            Catch ex As Exception
                ' MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub
End Class
