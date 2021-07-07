Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OITB
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        cargamenu()
        If actualizar Then
            cargaCampos()
        End If

    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

        If objGlobal.SBOApp.Menus.Exists("EXO-MnCOMI") = True Then
            Path = objGlobal.pathMenus   'objGlobal.compañia.conexionSAP.path & "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnCOMI.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnCOMI").Image = Path & "\MnCOMI.png"
                Else
                    'Sino existe lo copiamos y asignamos
                    EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), "MnCOMI.png", Path & "\MnCOMI.png")

                    Try
                        objGlobal.SBOApp.Menus.Item("EXO-MnCOMI").Image = Path & "\MnCOMI.png"
                    Catch
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OITB.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_OITB.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
    Public Overrides Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "1282"
                        'Comprobamos si el formulario es el de Grupos de artículos 
                        oForm = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm.TypeEx = "63" Then
                            If oForm.Visible = True Then
                                If oForm.Mode = BoFormMode.fm_ADD_MODE And CType(oForm.Items.Item("cbGCOMI").Specific, SAPbouiCOM.ComboBox).Selected Is Nothing Then
                                    CType(oForm.Items.Item("cbGCOMI").Specific, SAPbouiCOM.ComboBox).Select("SINCOMISION", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If
                        End If
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
                        Case "63"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "63"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "63"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_FORM_LOAD_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "63"
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

    Private Function EventHandler_FORM_LOAD_After(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        EventHandler_FORM_LOAD_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)
#Region "Campo Grupo de comisión"
            Dim oItem As SAPbouiCOM.Item

            oItem = oForm.Items.Add("cbGCOMI", BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = oForm.Items.Item("124").Top + oForm.Items.Item("124").Height + 5
            oItem.Left = oForm.Items.Item("124").Left
            oItem.Height = oForm.Items.Item("124").Height
            oItem.Width = oForm.Items.Item("124").Width
            oItem.FromPane = 1
            oItem.ToPane = 1
            oItem.Enabled = True
            oItem.DisplayDesc = True
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OITB", "U_EXO_GCOMI")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = oForm.Items.Add("lblGCOMI", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("cbGCOMI").Top
            oItem.Left = oForm.Items.Item("123").Left
            oItem.Height = oForm.Items.Item("123").Height
            oItem.Width = oForm.Items.Item("123").Width
            oItem.LinkTo = "cbGCOMI"
            oItem.FromPane = 1
            oItem.ToPane = 1
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Grupo de comisión"
#End Region
            CargaCombos(oForm)
            If oForm.Mode = BoFormMode.fm_ADD_MODE And CType(oForm.Items.Item("cbGCOMI").Specific, SAPbouiCOM.ComboBox).Selected Is Nothing Then
                CType(oForm.Items.Item("cbGCOMI").Specific, SAPbouiCOM.ComboBox).Select("SINCOMISION", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            EventHandler_FORM_LOAD_After = True


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

            'Grupo comisión
            sSQL = " Select ""Code"",""Name"" FROM ""@EXO_GCOMI"" Order by ""Name"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbGCOMI").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            End If


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

End Class
