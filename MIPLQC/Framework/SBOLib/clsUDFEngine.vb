Namespace Mukesh.SBOLib


    Public Class UDFEngine
        Private objCompany As SAPbobsCOM.Company

        Public Sub New(ByVal Company As SAPbobsCOM.Company)
            objCompany = Company
        End Sub


#Region "Table Functions"
        Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
            Dim objUserTableMD As SAPbobsCOM.UserTablesMD
            Dim ret As Integer
            Dim str As String = ""
            objUserTableMD = Nothing
            GC.Collect()
            objUserTableMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            Try
                If (Not objUserTableMD.GetByKey(TableName)) Then
                    objUserTableMD.TableName = TableName
                    objUserTableMD.TableDescription = TableDescription
                    objUserTableMD.TableType = TableType
                    If objUserTableMD.Add() = 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
                        objUserTableMD = Nothing
                        Return True
                    Else
                        objAddOn.objCompany.GetLastError(ret, str)
                        MsgBox(str)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
                        objUserTableMD = Nothing
                        Return False
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            Finally
                objUserTableMD = Nothing
                GC.Collect()
            End Try
        End Function
#End Region

        '#Region "Loading Default form"
        Public Sub LoadDefaultForm(ByVal sFormUID As String)
            Dim i As Integer
            ' Link to the Default Forms menu
            Dim sboMenu As SAPbouiCOM.MenuItem = objAddOn.objApplication.Menus.Item("47616")
            Try
                ' Iterate through the submenus to find the correct UDO
                If sboMenu.SubMenus.Count > 0 Then
                    For i = 0 To sboMenu.SubMenus.Count - 1
                        If sboMenu.SubMenus.Item(i).String.Contains(sFormUID) Then
                            sboMenu.SubMenus.Item(i).Activate()
                        End If
                    Next
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        'End Region

        Public Sub AddCol(ByVal TableName As String, ByVal ColName As String, ByVal ColDesc As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, Optional ByVal EditSize As Integer = 10, Optional ByVal SubType As SAPbobsCOM.BoFldSubTypes = 0)
            Dim objUserFields As SAPbobsCOM.UserFieldsMD
            Dim intError As Integer

            objUserFields = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            objUserFields.TableName = TableName
            objUserFields.Name = ColName
            objUserFields.Type = FieldType
            objUserFields.SubType = SubType
            objUserFields.Description = ColDesc
            objUserFields.EditSize = EditSize
            intError = objUserFields.Add()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFields)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            If intError <> 0 Then
                Throw New Exception(objCompany.GetLastErrorDescription)
            End If
        End Sub

        Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "")
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue)
        End Sub

        Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        End Sub

        Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue)
        End Sub

        Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
            Dim intLoop As Integer
            Dim ret As Integer
            Dim str As String
            Dim strValue, strDesc As Array
            Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD

            strValue = ValidValues.Split(Convert.ToChar(","))
            'MsgBox(strValue(0))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Valid value Code and Descriptions mismatching")
            End If
            ''new one
            objUserFieldMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            objUserFieldMD = Nothing
            GC.Collect()

            objUserFieldMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                If (Not isColumnExist(TableName, ColumnName)) Then
                    objUserFieldMD.TableName = TableName
                    objUserFieldMD.Name = ColumnName
                    objUserFieldMD.Description = ColDescription
                    objUserFieldMD.Type = FieldType
                    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                        objUserFieldMD.Size = Size
                    Else
                        objUserFieldMD.EditSize = Size
                    End If
                    objUserFieldMD.SubType = SubType
                    objUserFieldMD.DefaultValue = SetValidValue
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                    If objUserFieldMD.Add() <> 0 Then
                        objAddOn.objCompany.GetLastError(ret, str)
                        'MsgBox(Str)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                        objUserFieldMD = Nothing
                        ' Throw New Exception(objCompany.GetLastErrorDescription)
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                        objUserFieldMD = Nothing
                    End If
                End If
            Catch ex As Exception

                Throw ex
            Finally
                objUserFieldMD = Nothing
                GC.Collect()
            End Try
        End Sub

        Public Sub AddNumericField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        End Sub

        Public Sub AddNumericField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal DefultValue As String)
            Try
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, DefultValue)
            Catch ex As Exception
                MsgBox(ex.Message & TableName & ColumnName)
            End Try
        End Sub

        Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "")
        End Sub

        Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "")
        End Sub

        Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
            Dim objRecordSet As SAPbobsCOM.Recordset
            objRecordSet = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                If objAddOn.HANA Then
                    objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & TableName & "' AND ""AliasID"" = '" & ColumnName & "'")
                Else
                    objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE TableID = '" & TableName & "' AND AliasID = '" & ColumnName & "'")
                End If

                If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
                GC.Collect()
            End Try
        End Function

        Public Sub createUDO1(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
            objAddOn.objApplication.SetStatusBarMessage("UDO Created Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

            Dim creationPackage As SAPbouiCOM.FormCreationParams
            Dim objform As SAPbouiCOM.Form
            'Dim i As Integer
            Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
            Dim lRetCode As Long

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            'oUserObjectMD = Nothing
            'GC.Collect()

            oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUserObjectMD.GetByKey(udocode) Then
                oUserObjectMD.Code = udocode
                oUserObjectMD.Name = udoname
                oUserObjectMD.ObjectType = type
                oUserObjectMD.TableName = tblname
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.LogTableName = "A" + tblname
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                If DfltForm = True Then
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.LogTableName = tblname
                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname

                        Case "UDO Name"
                            oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                            oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                            oUserObjectMD.FormColumns.Add()
                            oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                            oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                            oUserObjectMD.FormColumns.Add()
                            oUserObjectMD.FormColumns.FormColumnAlias = "U_Rate"
                            oUserObjectMD.FormColumns.FormColumnDescription = "RatePerDay"
                            oUserObjectMD.FormColumns.Add()

                    End Select
                End If
                If FindForm = True Then
                    If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                        Select Case udoname

                            Case "UDO Name"
                                oUserObjectMD.FindColumns.ColumnAlias = "Code"
                                oUserObjectMD.FindColumns.ColumnDescription = "Code"
                                oUserObjectMD.FindColumns.Add()
                                oUserObjectMD.FindColumns.ColumnAlias = "Name"
                                oUserObjectMD.FindColumns.ColumnDescription = "Name"
                                oUserObjectMD.FindColumns.Add()
                        End Select

                    ElseIf type = SAPbobsCOM.BoUDOObjType.boud_Document Then
                        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                        Select Case udoname
                        End Select
                    End If
                End If
                'If childTable.Length > 0 Then
                '    For i = 0 To childTable.Length - 2
                '        If Trim(childTable(i)) <> "" Then
                '            oUserObjectMD.ChildTables.TableName = childTable(i)
                '            oUserObjectMD.ChildTables.Add()
                '        End If
                '    Next
                'End If
                lRetCode = oUserObjectMD.Add()
                If lRetCode <> 0 Then

                    MsgBox("error" + CStr(lRetCode))
                    MsgBox(objAddOn.objCompany.GetLastErrorDescription)
                Else

                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.Collect()


                If DfltForm = True Then
                    creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                    ' Need to set the parameter with the object unique ID
                    creationPackage.ObjectType = "1"
                    creationPackage.UniqueID = udoname
                    creationPackage.FormType = udoname
                    creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                    objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
                End If

            End If
            objAddOn.objApplication.SetStatusBarMessage("UDO Created Successfully..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Sub


        Public Sub createUDO(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            'Dim creationPackage As SAPbouiCOM.FormCreationParams
            'Dim objform As SAPbouiCOM.Form
            Dim i As Integer
            'Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
            Dim lRetCode As Long
            oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(udocode) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.Collect()
                Return
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.Collect()
            oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If 1 Then
                oUserObjectMD.Code = udocode
                oUserObjectMD.Name = udoname
                oUserObjectMD.ObjectType = type
                oUserObjectMD.TableName = tblname
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.LogTableName = "A" + tblname
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                If DfltForm = True Then
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                    'oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                    'oUserObjectMD.LogTableName = tblname
                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                    ' oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

                    'oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                    'oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                    'oUserObjectMD.FormColumns.Add()
                    'oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                    'oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                    'oUserObjectMD.FormColumns.Add()
                End If
                If FindForm = True Then
                    If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.FindColumns.ColumnAlias = "Code"
                        oUserObjectMD.FindColumns.ColumnDescription = "Code"
                        oUserObjectMD.FindColumns.Add()

                    Else
                        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                        'Select Case udocode

                        '    Case "MI_JOBCRD"
                        oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                        oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                        oUserObjectMD.FindColumns.Add()

                        ' End Select
                    End If
                End If
                If childTable.Length > 0 Then
                    For i = 0 To childTable.Length - 2
                        If Trim(childTable(i)) <> "" Then
                            oUserObjectMD.ChildTables.TableName = childTable(i)
                            oUserObjectMD.ChildTables.Add()
                        End If
                    Next
                End If
                lRetCode = oUserObjectMD.Add()
                If lRetCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription)
                Else
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.Collect()


                'If DfltForm = True Then
                '    creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                '    ' Need to set the parameter with the object unique ID
                '    creationPackage.ObjectType = udocode
                '    creationPackage.UniqueID = udocode
                '    creationPackage.FormType = udocode
                '    creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                '    objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
                'End If
            End If

        End Sub
        Public Sub createDefaultForm(ByVal UDOCode As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType)
            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            'Dim creationPackage As SAPbouiCOM.FormCreationParams
            'Dim objform As SAPbouiCOM.Form
            'Dim i As Integer
            'Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
            Dim lRetCode As Long
            oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(UDOCode) Then

                oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

                If UDOType = SAPbobsCOM.BoUDOObjType.boud_Document Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                    oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                    oUserObjectMD.FindColumns.Add()
                Else

                    oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                    oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                    oUserObjectMD.FormColumns.Add()
                    oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                    oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                    oUserObjectMD.FormColumns.Add()

                End If
                lRetCode = oUserObjectMD.Update()
                If lRetCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription)
                Else
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.Collect()
        End Sub

        Public Sub AddDefaultFormUDO(ByVal strUDO As String, ByVal strUDODesc As String, ByVal nObjectType As SAPbobsCOM.BoUDOObjType, ByVal strTable As String, ByVal childTable() As String, ByVal sFind() As String, _
                           ByVal DefaultForm As Boolean, ByVal DefaultColumnValue() As String, Optional ByVal canlog As Boolean = False, Optional ByVal Manageseries As Boolean = False)

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim tablecount As Integer = 0
            Try
                oUserObjectMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                If oUserObjectMD.GetByKey(strUDO) = 0 Then

                    oUserObjectMD.Code = strUDO
                    oUserObjectMD.Name = strUDODesc
                    oUserObjectMD.ObjectType = nObjectType
                    oUserObjectMD.TableName = strTable

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                    If DefaultForm Then
                        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES
                        If Not DefaultColumnValue Is Nothing Then
                            For i = 0 To DefaultColumnValue.Length - 1
                                If Not DefaultColumnValue(i) Is Nothing Then
                                    If DefaultColumnValue(i).Trim() <> String.Empty Then
                                        oUserObjectMD.FormColumns.FormColumnAlias = DefaultColumnValue(i)
                                        oUserObjectMD.FormColumns.FormColumnDescription = DefaultColumnValue(i)
                                        oUserObjectMD.FormColumns.Add()
                                    End If
                                End If
                            Next
                        End If
                    End If

                    If Manageseries Then oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES Else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    If canlog Then
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.LogTableName = "A" + strTable.ToString
                    Else
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                        oUserObjectMD.LogTableName = ""
                    End If

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.ExtensionName = ""

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    tablecount = 1
                    If sFind.Length > 0 Then
                        For i = 0 To sFind.Length - 1
                            If Trim(sFind(i)) = "" Then Continue For
                            oUserObjectMD.FindColumns.ColumnAlias = sFind(i)
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount)
                            tablecount = tablecount + 1
                        Next
                    End If

                    tablecount = 0
                    If Not childTable Is Nothing Then
                        If childTable.Length > 0 Then
                            For i = 0 To childTable.Length - 1
                                If Trim(childTable(i)) = "" Then Continue For
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount)
                                oUserObjectMD.ChildTables.TableName = childTable(i)
                                oUserObjectMD.ChildTables.Add()
                                tablecount = tablecount + 1
                            Next
                        End If
                    End If

                    If oUserObjectMD.Add() <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If

            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub
    End Class

End Namespace