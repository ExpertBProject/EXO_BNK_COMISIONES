Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_GCOM
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        cargamenu()
        If actualizar Then
            CrearView()
            cargaCampos()
        End If

    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        Dim sSQL As String = ""

        If objGlobal.SBOApp.Menus.Exists("EXO-MnCOMI") = True Then
            sSQL = "SELECT T0.""U_EXO_PATH"" FROM ""@EXO_OGEN""  T0 "
            Path = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            Path &= "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnCOMI.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnCOMI").Image = Path & "\MnCOMI.png"
                Else
                    'Sino existe lo copiamos y asignamos
                    EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), "MnCOMI.png", Path & "\MnCOMI.png")

                    objGlobal.SBOApp.Menus.Item("EXO-MnCOMI").Image = Path & "\MnCOMI.png"
                End If
            End If
        End If
    End Sub
    Private Sub CrearView()
        Dim sSQL As String = ""
        Dim bResultado As String = ""
        If objGlobal.refDi.comunes.esAdministrador Then
            sSQL = EXO_GLOBALES.LeerEmbebido("EXO_BNK_COMISIONES.VIEW_FACTURAS.sql")
            bResultado = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            objGlobal.SBOApp.StatusBar.SetText("Creado View: VIEW_FACTURAS", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_GCOM.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_GCOM.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OINV.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_OINV.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnGCOM"
                        'Cargamos UDO 
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_GCOM")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_GCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_GCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_Before(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select

                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_GCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_GCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Freeze(True)
                'Empleado
                oForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                'Año
                oForm.Items.Item("13_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("13_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("13_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                'Código
                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                CargaCombos(oForm)

                '' A petición de Carlos Zea, ocultamos columna 3 y 4
                'CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Visible = False
                'CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Visible = False
            End If

            EventHandler_FORM_VISIBLE = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Function
    Private Function CargaCombos(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaCombos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)

            'Vendedores
            sSQL = " Select ""SlpCode"",""SlpName"" FROM ""OSLP"" Order by ""SlpName"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("Item_1").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("Item_1").Specific, SAPbouiCOM.ComboBox).Item.DisplayDesc = True
            End If

            'Grupo de vendedores
            sSQL = " SELECT ""Code"" , ""Name"" FROM ""@EXO_GCOMI"" Order by ""Name"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").DisplayDesc = True
            'CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ExpandType = BoExpandType.et_DescriptionOnly

            CargaCombos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sImporte As String = "" : Dim dImporte As Double = 0

        EventHandler_VALIDATE_After = False
        Try
            oForm.Freeze(True)
            If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                If pVal.ItemUID = "Item_1" Or pVal.ItemUID = "13_U_E" Then
                    CrearCodigo(oForm)
                End If
            ElseIf pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_3" Then
                sImporte = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value
                dImporte = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, sImporte)
                If dImporte = 0 Then
                    CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = 0
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(pVal.Row, 4, False)
                Else

                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(pVal.Row, 4, True)
                End If
            End If

            EventHandler_VALIDATE_After = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Sub CrearCodigo(ByRef oForm As SAPbouiCOM.Form)
        Dim sAnno As String = "" : Dim sVendedor As String = "" : Dim sComision As String = ""
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            sAnno = CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
            If CType(oForm.Items.Item("Item_1").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sVendedor = CType(oForm.Items.Item("Item_1").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
            End If

            sComision = sAnno & "_" & sVendedor

            If sComision.Length > 1 Then
                'Buscamos para ver si existe
                sSQL = "SELECT ""Code"" FROM ""@EXO_GCOM"" WHERE ""Code""='" & sComision & "' "
                oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = sComision
                Else
                    oForm.Mode = BoFormMode.fm_OK_MODE
                    objGlobal.SBOApp.ActivateMenuItem("1281") ' Buscar
                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = sComision
                    CType(oForm.Items.Item("1").Specific, SAPbouiCOM.Button).Item.Click()
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Function EventHandler_COMBO_SELECT(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sGrupo As String = ""
        EventHandler_COMBO_SELECT = False
        Dim sTable As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemChanged = True And pVal.ItemUID = "Item_1" Then
                CrearCodigo(oForm)
            ElseIf pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then
                If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sGrupo = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    'Comprobamos que en toda la matrix no haya mas de 1 codigo de grupo seleccionado
                    If MatrixToNet(oForm, sGrupo) = False Then
                        Exit Function
                    End If
                End If
            End If


            EventHandler_COMBO_SELECT = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_VALIDATE_Before(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sGrupo As String = ""
        EventHandler_VALIDATE_Before = False
        Dim sTable As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then
                If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sGrupo = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    'Comprobamos que en toda la matrix no haya mas de 1 codigo de grupo seleccionado
                    If MatrixToNet(oForm, sGrupo) = False Then
                        Exit Function
                    End If
                End If
            End If


            EventHandler_VALIDATE_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
#Region "Métodos auxiliares"
    Private Function MatrixToNet(ByRef oForm As SAPbouiCOM.Form, ByVal sGrupo As String) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing
        Dim sGrupoleido As String = "" : Dim iGrupoTotal As Integer = 0
        Dim oArrCampos As Boolean = False
        Dim sMatrixUID As String = ""

        MatrixToNet = False

        Try
            sXML = CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")
            iGrupoTotal = 0

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")

                'Inicializamos para de dejar a False

                oArrCampos = False

                'Inicializamos los datos del registro
                sGrupoleido = ""

                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "C_0_1" Then 'CodigoGrupo
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")

                        sGrupoleido = oXmlNodeField.InnerText

                        oArrCampos = True
                        If sGrupo = sGrupoleido Then
                            iGrupoTotal += 1
                        End If
                    End If

                    If oArrCampos = True And iGrupoTotal >= 2 Then
                        Exit For
                    End If
                Next

                'Hemos recorrido el registro, y comprobamos el almacén
                If iGrupoTotal >= 2 Then
                    objGlobal.SBOApp.StatusBar.SetText("No es posible seleccionar el grupo " & sGrupo & " varias veces", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
            Next

            MatrixToNet = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
End Class
