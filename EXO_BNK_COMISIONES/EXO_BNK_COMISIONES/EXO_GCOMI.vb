Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_GCOMI
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
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_GCOMI.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_GCOMI.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'Introducir los datos de los grupos 
            Carga_Datos_GCOM()
        End If
    End Sub
    Private Sub Carga_Datos_GCOM()
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
            Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
            Dim oCompService As SAPbobsCOM.CompanyService = objGlobal.compañia.GetCompanyService()

            sSQL = "SELECT * FROM ""@EXO_GCOMI"" WHERE Code='SINCOMISION' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount = 0 Then
                objGlobal.SBOApp.StatusBar.SetText("Actualizando datos de Grupos de comisiones ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oGeneralService = oCompService.GetGeneralService("EXO_GCOMI")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralData.SetProperty("Code", "SINCOMISION")
                'oGeneralData.SetProperty("DocEntry", CInt(sCodigo))
                oGeneralData.SetProperty("Name", "Sin Comisión")
                oGeneralService.Add(oGeneralData)
                objGlobal.SBOApp.StatusBar.SetText("Actualizado datos de Grupos de comisiones.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
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
                    Case "EXO-MnGCOMI"
                        'Cargamos UDO 
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_GCOMI")
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

End Class
