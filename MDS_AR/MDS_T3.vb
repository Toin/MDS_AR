
Option Explicit On
Option Strict Off

Imports MDS_AR.MIS_Utils

Public Class MDS_T3

    Public oCompany As SAPbobsCOM.Company
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter
    Public blnFind As Boolean = False
    Dim oMIS_Utils As New MIS_Utils

    'Error handling variables
    Public sErrMsg As String
    Public lErrCode As Integer
    Public lRetCode As Integer

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Or EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged _
        Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then
            SBO_Application.MessageBox("MDS AR Collection Addon now terminate...")
            End
        End If
    End Sub


    Public Sub New()
        MyBase.New()
        'Class_Initialize_Renamed()
        If Not MDS_AR.SBOConnection.SBOApplication Is Nothing And Not MDS_AR.SBOConnection.SBOCompany Is Nothing Then
            SBO_Application = MDS_AR.SBOConnection.SBOApplication
            oCompany = MDS_AR.SBOConnection.SBOCompany

        Else
            Dim SBOConnection = New SBOConnection
            SBO_Application = MDS_AR.SBOConnection.SBOApplication
            oCompany = MDS_AR.SBOConnection.SBOCompany
        End If

        SBO_Application.SetStatusBarMessage("connected DI!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        Try
            LoadFromXML_Menu("MDSARMenus.xml")

        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try

        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Private Function LoadFromXML(ByVal FileName As String) As String

        Dim oXmlDoc As Xml.XmlDocument
        Dim sPath As String

        oXmlDoc = New Xml.XmlDocument

        sPath = System.Windows.Forms.Application.StartupPath

        oXmlDoc.Load(sPath & "\" & FileName)

        Return (oXmlDoc.InnerXml)

    End Function

    Private Sub LoadFromXML_Menu(ByVal FileName As String)
        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        Dim sPath As String
        sPath = System.Windows.Forms.Application.StartupPath
        oXmlDoc.Load(sPath & "\" & FileName)

        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = SBO_Application.GetLastBatchResults()

    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        If BusinessObjectInfo.BeforeAction = True Then
            If BusinessObjectInfo.FormTypeEx = "Pdc" Then
                Dim oForm As SAPbouiCOM.Form = Nothing

                If BusinessObjectInfo.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And BusinessObjectInfo.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And BusinessObjectInfo.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    oForm = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                End If
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        SBO_Application.MessageBox("Data Can,t Update", 1, "Ok")
                        BubbleEvent = False
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        SBO_Application.MessageBox("Data Can,t Save", 1, "Ok")
                        BubbleEvent = False
                End Select
            End If
        End If
    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        PdcInputAplicationItem(FormUID, pVal, BubbleEvent)
        PdcTolakAplicationItem(FormUID, pVal, BubbleEvent)

        If pVal.FormTypeEx = "-170" Then
            If pVal.ItemUID = "U_PDCNo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.ActionSuccess = True Then
                Dim oFormUdf As SAPbouiCOM.Form = Nothing
                Dim oFormMain As SAPbouiCOM.Form = Nothing
                Dim objColumnsPayment As SAPbouiCOM.Columns = Nothing
                Dim objMatrixPayment As SAPbouiCOM.Matrix = Nothing
                Dim strsql As String
                Dim oRecSet As SAPbobsCOM.Recordset = Nothing
                oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oFormUdf = SBO_Application.Forms.Item(FormUID)
                oFormMain = SBO_Application.Forms.GetForm(170, 1)
                objMatrixPayment = oFormMain.Items.Item("20").Specific
                objColumnsPayment = objMatrixPayment.Columns

                If oFormMain.Items.Item("5").Specific.value = "" Then
                    SBO_Application.SetStatusBarMessage("BP Code Must fill", SAPbouiCOM.BoMessageTime.bmt_Short)
                    BubbleEvent = False
                End If

                If oFormUdf.Items.Item("U_PDCBankID").Specific.value = "" Then
                    SBO_Application.SetStatusBarMessage("Bank Id Must fill", SAPbouiCOM.BoMessageTime.bmt_Short)
                    BubbleEvent = False
                End If

                If oFormUdf.Items.Item("U_PDCNo").Specific.value = "" Then
                    SBO_Application.SetStatusBarMessage("PDC No Must fill", SAPbouiCOM.BoMessageTime.bmt_Short)
                End If

                If BubbleEvent = True Then
                    For I = 1 To objMatrixPayment.RowCount

                        strsql = "SELECT T1.U_OINVDocNum,T1.U_CollectAmount,T0.U_CardCode " & _
                            "FROM [@MIS_PDC] AS T0 RIGHT OUTER JOIN [@MIS_PDCL] AS T1 ON T1.DocEntry = T0.DocEntry " & _
                            "WHERE T0.U_CardCode='" & oFormMain.Items.Item("5").Specific.value & "' AND " & _
                            "T0.U_PDCBankID='" & oFormUdf.Items.Item("U_PDCBankID").Specific.value & "' AND " & _
                            "T0.U_PDCNo='" & oFormUdf.Items.Item("U_PDCNo").Specific.value & "' AND " & _
                            "T1.U_OINVDocNum='" & objColumnsPayment.Item("1").Cells.Item(I).Specific.value & "'"
                        oRecSet.DoQuery(strsql)

                        If oRecSet.RecordCount > 0 Then
                            objColumnsPayment.Item("10000127").Cells.Item(I).Specific.Checked = True
                            objColumnsPayment.Item("24").Cells.Item(I).Specific.value = oRecSet.Fields.Item("U_CollectAmount").Value
                        Else
                            objColumnsPayment.Item("10000127").Cells.Item(I).Specific.Checked = False
                            objColumnsPayment.Item("24").Cells.Item(I).Specific.value = objColumnsPayment.Item("7").Cells.Item(I).Specific.value
                        End If

                    Next
                End If
               
                oFormUdf = Nothing
                oFormMain = Nothing
                objColumnsPayment = Nothing
                objMatrixPayment = Nothing
            End If
        End If

        Select Case FormUID
            Case "DeleteT3"
                'karno delete T3 Tahap 3 (2011.05.31 11:45:00)
                If pVal.ItemUID = "T3Number" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBO_Application.Forms.Item(FormUID)
                    If oForm.Items.Item("T3Number").Specific.value <> "" Then
                        LostFocusCollect(oForm)
                    End If
                End If
                'karno delete T3 Tahap 3 (2011.05.31 13:34:00)
                If pVal.BeforeAction = True Then
                    If pVal.ItemUID = "BtnShow" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oDeleteT3StatusGrid As SAPbouiCOM.Grid

                        Dim dt As SAPbouiCOM.DataTable

                        dt = oForm.DataSources.DataTables.Item("DelT3StatusLst")

                        oDeleteT3StatusGrid = oForm.Items.Item("myGridGen").Specific

                        DeleteT3Show(oForm)
                    End If
                ElseIf pVal.ItemUID = "BtnDel" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBO_Application.Forms.Item(FormUID)
                    Dim oDeleteT3StatusGrid As SAPbouiCOM.Grid

                    Dim dt As SAPbouiCOM.DataTable

                    dt = oForm.DataSources.DataTables.Item("DelT3StatusLst")

                    oDeleteT3StatusGrid = oForm.Items.Item("myGridGen").Specific

                    DeleteT3Status(oForm)
                    DeleteT3Show(oForm)
                End If

                'karno Input T3 Tahap 3 (2011.05.30 09:41:00)
            Case "InputReceiptT3"
                    If pVal.BeforeAction = True Then
                        If pVal.ColUID = "Receipt T3" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = SBO_Application.Forms.Item(FormUID)
                            Dim oInputT3StatusGrid As SAPbouiCOM.Grid = Nothing
                            Dim dt As SAPbouiCOM.DataTable
                            dt = oForm.DataSources.DataTables.Item("InT3StatusLst")
                            oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific
                            SelectUnselect(oForm)
                        End If
                    End If

                    If pVal.ItemUID = "ColCode" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        LostFocusCardCode(oForm)

                    ElseIf pVal.ItemUID = "BtnShow" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.BeforeAction = True Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = SBO_Application.Forms.Item(FormUID)
                            Dim oInputT3StatusGrid As SAPbouiCOM.Grid = Nothing

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("InT3StatusLst")

                            oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific

                            InputT3Show(oForm)
                        End If

                    ElseIf pVal.ItemUID = "BtnAdd" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.BeforeAction = True Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = SBO_Application.Forms.Item(FormUID)
                            Dim oInputT3StatusGrid As SAPbouiCOM.Grid = Nothing

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("InT3StatusLst")

                            oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific

                            InputT3Status(oForm)
                            InputT3Show(oForm)
                        End If
                    End If

                    'karno generate T3 Tahap 3 (2011.05.26 16:35:00)
            Case "GenerateT3"
                If pVal.BeforeAction = True Then
                    If pVal.ColUID = "AR Generate" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid = Nothing
                        Dim dt As SAPbouiCOM.DataTable
                        dt = oForm.DataSources.DataTables.Item("GenT3StatusLst")
                        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific
                        SelectUnselectGenerate(oForm)
                    End If
                End If

                If pVal.BeforeAction = True Then
                    If pVal.ItemUID = "BtnShow" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid

                        Dim dt As SAPbouiCOM.DataTable

                        dt = oForm.DataSources.DataTables.Item("GenT3StatusLst")

                        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific

                        GenerateT3Show(oForm)

                    ElseIf pVal.ItemUID = "BtnOK" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid

                        Dim dt As SAPbouiCOM.DataTable

                        dt = oForm.DataSources.DataTables.Item("GenT3StatusLst")

                        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific

                        GenerateT3Status(oForm)
                        GenerateT3Show(oForm)
                    End If
                End If
        End Select

        'karno generate T3 tahap 3 (2011.05.26 16:35:00)
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        If pVal.BeforeAction = False Then
            Select Case pVal.MenuUID
                'karno generate T3 Tahap 1(2011.05.26 16:35:00)
                Case "AR01_01"
                    GenerateT3()
                    'karno Input T3 Tahap 1 (2011.05.30 09:20:00)
                Case "AR01_02"
                    InputT3()
                    'karno Delete T3 Tahap 1 (2011.05.31 11:15:00)
                Case "AR01_03"
                    DeleteT3()
                Case "AR01_04"
                    PdcFirstLoad()
                Case "AR01_05"
                    PdcTolakFirstLoad()
            End Select
        End If
        PdcInputAplicationMenu(pVal, BubbleEvent)

        If pVal.BeforeAction = True Then
            Dim oForm As SAPbouiCOM.Form

            oForm = SBO_Application.Forms.ActiveForm
            'MsgBox(oForm.Type)
            'MsgBox(oForm.TypeEx)
            'MsgBox(oForm.UniqueID)
            Select Case pVal.MenuUID
                Case "1290" ' 1st Record
                Case "1289" ' Prev Record
                Case "1288" ' Next Record
                Case "1291" ' Last Record
                Case "1292" ' Add a row
                Case "1293" ' Delete a row
                Case "1282"
            End Select

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            GC.Collect()

        End If
    End Sub
    ' Karno Lost Focus (2011.05.31 11:46:00)

    Private Sub LostFocusCollect(ByVal oForm As SAPbouiCOM.Form)
        Dim StrSql As String
        Dim oRecSet As SAPbobsCOM.Recordset = Nothing
        oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        StrSql = "SELECT T0.U_CardCode, T0.U_CardName, T0.U_DocDate, T0.U_CollectID, T1.U_CollectorName FROM [@MIS_T3] T0" & _
                " LEFT JOIN [@COLLECTOR] T1 ON T0.U_CollectID = T1.U_CollectorID WHERE T0.DocNum = '" & oForm.Items.Item("T3Number").Specific.Value & "' "
        oRecSet.DoQuery(StrSql)

        If oRecSet.RecordCount = 0 Then
            SBO_Application.StatusBar.SetText("Not Matching Record Found !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Items.Item("CardCode").Specific.Value = ""
            oForm.Items.Item("Customer").Specific.Value = ""
            oForm.Items.Item("T3Date").Specific.Value = ""
            oForm.Items.Item("CollectId").Specific.Value = ""
            oForm.Items.Item("Collector").Specific.Value = ""
            oForm.Items.Item("T3Number").Click()
            oForm.Items.Item("CardCode").Enabled = True
            oForm.Items.Item("Customer").Enabled = True
            oForm.Items.Item("T3Date").Enabled = True
            oForm.Items.Item("CollectId").Enabled = True
            oForm.Items.Item("Collector").Enabled = True
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            Exit Sub
        Else
            oForm.Items.Item("CardCode").Specific.Value = oRecSet.Fields.Item("U_CardCode").Value
            oForm.Items.Item("Customer").Specific.Value = oRecSet.Fields.Item("U_CardName").Value
            oForm.Items.Item("T3Date").Specific.string = oMIS_Utils.fctFormatDate(oRecSet.Fields.Item("U_DocDate").Value, oCompany, 5)
            oForm.Items.Item("CollectId").Specific.Value = oRecSet.Fields.Item("U_CollectID").Value
            oForm.Items.Item("Collector").Specific.Value = oRecSet.Fields.Item("U_CollectorName").Value
            oForm.Items.Item("T3Number").Click()
            oForm.Items.Item("CardCode").Enabled = True
            oForm.Items.Item("Customer").Enabled = False
            oForm.Items.Item("T3Date").Enabled = False
            oForm.Items.Item("CollectId").Enabled = False
            oForm.Items.Item("Collector").Enabled = False
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)

    End Sub

    Private Sub LostFocusCardCode(ByVal oForm As SAPbouiCOM.Form)
        Dim StrSql As String
        Dim oRecSet As SAPbobsCOM.Recordset = Nothing
        oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        StrSql = "SELECT U_CollectorName CardName FROM [@COLLECTOR] WHERE U_CollectorID = '" & oForm.Items.Item("ColCode").Specific.Value & "' "
        oRecSet.DoQuery(StrSql)

        If oRecSet.RecordCount = 0 Then
            SBO_Application.StatusBar.SetText("Collector Name Not Matching Record Found ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Items.Item("ColName").Specific.Value = ""
            oForm.Items.Item("ColCode").Click()
            oForm.Items.Item("ColName").Enabled = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            Exit Sub
        Else
            oForm.Items.Item("ColName").Specific.Value = oRecSet.Fields.Item("CardName").Value
            oForm.Items.Item("T3Date").Click()
            oForm.Items.Item("ColName").Enabled = False
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
    End Sub

    Private Sub SelectUnselectGenerate(ByVal oForm As SAPbouiCOM.Form)
        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid
        Dim idx As Long
        Dim dt As SAPbouiCOM.DataTable

        dt = oForm.DataSources.DataTables.Item("GenT3StatusLst")
        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        If oGenerateT3StatusGrid.Columns.Item("AR Generate").TitleObject.Caption = "Select All" Then
            'select/check all
            oForm.Freeze(True)
            For idx = 0 To oGenerateT3StatusGrid.Rows.Count - 1
                dt.SetValue("AR Generate", idx, "Y")
            Next
            oGenerateT3StatusGrid.Columns.Item("AR Generate").TitleObject.Caption = "Reset All"
            oForm.Freeze(False)
        Else
            'unselect/uncheck all
            oForm.Freeze(True)
            For idx = 0 To oGenerateT3StatusGrid.Rows.Count - 1
                dt.SetValue("AR Generate", idx, "N")
            Next
            oGenerateT3StatusGrid.Columns.Item("AR Generate").TitleObject.Caption = "Select All"
            oForm.Freeze(False)
        End If
    End Sub

    Private Sub SelectUnselect(ByVal oForm As SAPbouiCOM.Form)
        Dim oInputT3StatusGrid As SAPbouiCOM.Grid
        Dim idx As Long
        Dim dt As SAPbouiCOM.DataTable

        dt = oForm.DataSources.DataTables.Item("InT3StatusLst")
        oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        If oInputT3StatusGrid.Columns.Item("Receipt T3").TitleObject.Caption = "Select All" Then
            'select/check all
            oForm.Freeze(True)
            For idx = 0 To oInputT3StatusGrid.Rows.Count - 1
                dt.SetValue("Receipt T3", idx, "Y")
            Next
            oInputT3StatusGrid.Columns.Item("Receipt T3").TitleObject.Caption = "Reset All"
            oForm.Freeze(False)
        Else
            'unselect/uncheck all
            oForm.Freeze(True)
            For idx = 0 To oInputT3StatusGrid.Rows.Count - 1
                dt.SetValue("Receipt T3", idx, "N")
            Next
            oInputT3StatusGrid.Columns.Item("Receipt T3").TitleObject.Caption = "Select All"
            oForm.Freeze(False)
        End If
    End Sub

    'karno Input T3  Tahap 2(2011.05.30 10:35:00)
    Private Sub DeleteT3()
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing

        Dim oDeleteT3StatusGrid As SAPbouiCOM.Grid = Nothing

        Try
            oForm = SBO_Application.Forms.Item("DeleteT3")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "DeleteT3"
            fcp.UniqueID = "DeleteT3"
            fcp.XmlData = LoadFromXML("DeleteT3.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)
            'oForm.ClientHeight = 448
            oForm.DataSources.DataTables.Add("DelT3StatusLst")
            oForm.DataSources.UserDataSources.Add("T3Number", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("Customer", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("T3Date", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("CollectId", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("Collector", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            oEditText = oForm.Items.Item("T3Number").Specific
            oEditText.DataBind.SetBound(True, "", "T3Number")
            oEditText = oForm.Items.Item("CardCode").Specific
            oEditText.DataBind.SetBound(True, "", "CardCode")
            oEditText = oForm.Items.Item("Customer").Specific
            oEditText.DataBind.SetBound(True, "", "Customer")
            oEditText = oForm.Items.Item("T3Date").Specific
            oEditText.DataBind.SetBound(True, "", "T3Date")
            oEditText = oForm.Items.Item("CollectId").Specific
            oEditText.DataBind.SetBound(True, "", "CollectId")
            oEditText = oForm.Items.Item("Collector").Specific
            oEditText.DataBind.SetBound(True, "", "Collector")

            oForm.Items.Item("T3Date").Specific.value = DateTime.Today.ToString("yyyyMMdd")

            ' add a GRID item to the form
            oItem = oForm.Items.Add("myGridGen", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 120
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            oDeleteT3StatusGrid = oItem.Specific

            oDeleteT3StatusGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            GC.Collect()
            oForm.Freeze(False)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDeleteT3StatusGrid)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
            oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oDeleteT3StatusGrid = Nothing



        End Try


    End Sub

    'karno Input T3  Tahap 2(2011.05.30 10:35:00)
    Private Sub InputT3()
        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText

        Dim oInputT3StatusGrid As SAPbouiCOM.Grid

        Try
            oForm = SBO_Application.Forms.Item("InputReceiptT3")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "InputReceiptT3"
            fcp.UniqueID = "InputReceiptT3"
            fcp.XmlData = LoadFromXML("InputReceiptT3.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)
            'oForm.ClientHeight = 443
            oForm.DataSources.DataTables.Add("InT3StatusLst")
            oForm.DataSources.UserDataSources.Add("T3DtRcpt", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("ColCode", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("ColName", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("T3Date", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("Region", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            
            oEditText = oForm.Items.Item("T3DtRcpt").Specific
            oEditText.DataBind.SetBound(True, "", "T3DtRcpt")
            oEditText = oForm.Items.Item("ColCode").Specific
            oEditText.DataBind.SetBound(True, "", "ColCode")
            oEditText = oForm.Items.Item("ColName").Specific
            oEditText.DataBind.SetBound(True, "", "ColName")
            oEditText = oForm.Items.Item("T3Date").Specific
            oEditText.DataBind.SetBound(True, "", "T3Date")
            oEditText = oForm.Items.Item("Region").Specific
            oEditText.DataBind.SetBound(True, "", "Region")

            oForm.Items.Item("T3DtRcpt").Specific.value = DateTime.Today.ToString("yyyyMMdd")


            ' add a GRID item to the form
            oItem = oForm.Items.Add("myGridGen", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 120
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            oInputT3StatusGrid = oItem.Specific

            oInputT3StatusGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            GC.Collect()

            oForm.Freeze(False)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oInputT3StatusGrid)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
            oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oInputT3StatusGrid = Nothing



        End Try
    End Sub

    'karno generate T3  Tahap 2(2011.05.26 16:35:00)
    Private Sub GenerateT3()
        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oCombobox As SAPbouiCOM.ComboBox = Nothing
        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid
        'Dim iDocNum As Long

        Try
            oForm = SBO_Application.Forms.Item("GenerateT3")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "GenerateT3"
            fcp.UniqueID = "GenerateT3"
            fcp.XmlData = LoadFromXML("GenerateT3.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)
            'oForm.ClientHeight = 476
            oForm.DataSources.DataTables.Add("GenT3StatusLst")

            'oForm.DataSources.UserDataSources.Add("CmbNo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            'oForm.DataSources.UserDataSources.Add("T3No", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            oForm.DataSources.UserDataSources.Add("DateFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("DateTo", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("T3Date", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("Customer", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Item("DateFrom").Value = oForm.Items.Item("DateFrom").Specific.string
            oForm.DataSources.UserDataSources.Item("DateTo").Value = oForm.Items.Item("DateTo").Specific.string
            oForm.DataSources.UserDataSources.Item("T3Date").Value = oForm.Items.Item("T3Date").Specific.string
            oForm.DataSources.UserDataSources.Item("Customer").Value = oForm.Items.Item("Customer").Specific.string

            'oCombobox = oForm.Items.Item("CmbNo").Specific
            'oCombobox.ValidValues.LoadSeries("T3", SAPbouiCOM.BoSeriesMode.sf_Add)
            'oCombobox.DataBind.SetBound(True, "", "CmbNo")

            'oCombobox.SelectExclusive("2011", SAPbouiCOM.BoSearchKey.psk_ByDescription)

            'oEditText = oForm.Items.Item("T3No").Specific
            'oEditText.DataBind.SetBound(True, "", "T3No")

            'iDocNum = oForm.BusinessObject.GetNextSerialNumber("Series", "T3")
            'oEditText = oForm.Items.Item("T3No").Specific
            'oEditText.String = iDocNum
            'oForm = SBO_Application.Forms.ActiveForm
         

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)

            oEditText = oForm.Items.Item("DateFrom").Specific
            oEditText.DataBind.SetBound(True, "", "DateFrom")
            oEditText = oForm.Items.Item("DateTo").Specific
            oEditText.DataBind.SetBound(True, "", "DateTo")
            oEditText = oForm.Items.Item("T3Date").Specific
            oEditText.DataBind.SetBound(True, "", "T3Date")
            oEditText = oForm.Items.Item("Customer").Specific
            oEditText.DataBind.SetBound(True, "", "Customer")


            oForm.Items.Item("DateTo").Specific.value = DateTime.Today.ToString("yyyyMMdd")
            oForm.Items.Item("T3Date").Specific.value = DateTime.Today.ToString("yyyyMMdd")
            ' add a GRID item to the form
            oItem = oForm.Items.Add("myGridGen", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 120
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            oGenerateT3StatusGrid = oItem.Specific

            oGenerateT3StatusGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Items.Item("DateFrom").Click()

            GC.Collect()
            oForm.Freeze(False)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGenerateT3StatusGrid)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
            oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oGenerateT3StatusGrid = Nothing



        End Try


    End Sub

    'karno Delete T3  Tahap 2(2011.05.26 13:35:00)
    Private Sub DeleteT3Show(ByVal oForm As SAPbouiCOM.Form)

        Dim ARStatusQuery As String
        Dim oDeleteT3StatusGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

        Dim oMis_Utils As MIS_Utils

        oMis_Utils = New MIS_Utils

        oDeleteT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        If oForm.Items.Item("T3Number").Specific.value = "" Then
            SBO_Application.MessageBox("T3 Number Must fill", 1, "OK")
            oForm.Items.Item("T3Number").Click()
            Exit Sub
        Else

            ARStatusQuery = "SELECT 'N' AS [SelectAll], T1.LineId, T1.U_OINVDocNum AS [Invoice No], T1.U_TAXTaxNum AS [FP No], T1.U_OINVDocDate AS [Invoice Date], T1.U_OINVDocDueDate [Due Date], T1.U_OINVNetAmount [Invoice Net Amount], T1.U_OINVDocTotal [Invoice Amount]" & _
                            " FROM [@MIS_T3] T0 JOIN [@MIS_T3L] T1 ON T0.DocEntry = T1.DocEntry WHERE T1.U_T3LineStatus <> 'D' AND T0.DocNum = " & oForm.Items.Item("T3Number").Specific.value & ""


        End If
        oForm.DataSources.DataTables.Item("DelT3StatusLst").ExecuteQuery(ARStatusQuery)
        oDeleteT3StatusGrid.DataTable = oForm.DataSources.DataTables.Item("DelT3StatusLst")

        oDeleteT3StatusGrid.Columns.Item("SelectAll").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oDeleteT3StatusGrid.Columns.Item("SelectAll").TitleObject.Sortable = True

        oColumn = oDeleteT3StatusGrid.Columns.Item("LineId")
        oColumn.Editable = False

        oColumn = oDeleteT3StatusGrid.Columns.Item("Invoice No")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        oDeleteT3StatusGrid.Columns.Item("Invoice No").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oDeleteT3StatusGrid.Columns.Item("FP No")
        oDeleteT3StatusGrid.Columns.Item("FP No").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oDeleteT3StatusGrid.Columns.Item("Invoice Date")
        oDeleteT3StatusGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oDeleteT3StatusGrid.Columns.Item("Due Date")
        oDeleteT3StatusGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oDeleteT3StatusGrid.Columns.Item("Invoice Net Amount")
        oDeleteT3StatusGrid.Columns.Item("Invoice Net Amount").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oDeleteT3StatusGrid.Columns.Item("Invoice Amount")
        oDeleteT3StatusGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        oColumn.Editable = False

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDeleteT3StatusGrid)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

    End Sub

    'karno Input T3  Tahap 2(2011.05.30 10:26:00)
    Private Sub InputT3Show(ByVal oForm As SAPbouiCOM.Form)

        Dim InputT3StatusQuery As String
        Dim oInputT3StatusGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

        Dim oMis_Utils As MIS_Utils

        oMis_Utils = New MIS_Utils

        oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        If oForm.Items.Item("T3Date").Specific.value = "" Then
            SBO_Application.MessageBox("T3 Date From Must fill", 1, "OK")
            oForm.Items.Item("T3Date").Click()
            Exit Sub
            'ElseIf oForm.Items.Item("Region").Specific.value = "" Then
            '    SBO_Application.MessageBox("Region To Must fill", 1, "OK")
            '    oForm.Items.Item("Region").Click()
            '    Exit Sub
        ElseIf oForm.Items.Item("T3DtRcpt").Specific.value = "" Then
            SBO_Application.MessageBox("T3 Receipt Date Must Fill", 1, "OK")
            oForm.Items.Item("T3DtRcpt").Click()
            Exit Sub
        ElseIf oForm.Items.Item("ColCode").Specific.value = "" Then
            SBO_Application.MessageBox("Collector Code Must Fill", 1, "OK")
            oForm.Items.Item("ColCode").Click()
            Exit Sub
        ElseIf oForm.Items.Item("ColName").Specific.value = "" Then
            SBO_Application.MessageBox("Collector Name Must Fill", 1, "OK")
            oForm.Items.Item("ColName").Click()
            Exit Sub
        Else

            InputT3StatusQuery = "SELECT Distinct 'Y' AS [Receipt T3], T0.U_DocDate AS T3Date, T0.U_WilayahCollect AS Wilayah, T0.DocNum AS T3No, T0.U_CardCode AS CustomerCode, T0.U_CardName AS CustomerName FROM [@MIS_T3] T0 " & _
                            "INNER JOIN [@MIS_T3L] T1 " & _
                            "ON T0.DocEntry = T1.DocEntry " & _
                            "WHERE T0.U_DocDate = '" & oForm.Items.Item("T3Date").Specific.value & "' " & _
                            "AND T0.U_WilayahCollect LIKE '%" & oForm.Items.Item("Region").Specific.value & "%' AND T1.U_T3LineStatus <> 'R' AND T1.U_T3LineStatus = 'A'  " & _
                            "ORDER BY T0.U_DocDate, T0.U_WilayahCollect"

        End If
        oForm.DataSources.DataTables.Item("InT3StatusLst").ExecuteQuery(InputT3StatusQuery)
        oInputT3StatusGrid.DataTable = oForm.DataSources.DataTables.Item("InT3StatusLst")

        oInputT3StatusGrid.Columns.Item("Receipt T3").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oInputT3StatusGrid.Columns.Item("Receipt T3").TitleObject.Sortable = True

        oColumn = oInputT3StatusGrid.Columns.Item("T3Date")
        oInputT3StatusGrid.Columns.Item("T3Date").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oInputT3StatusGrid.Columns.Item("Wilayah")
        oInputT3StatusGrid.Columns.Item("Wilayah").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oInputT3StatusGrid.Columns.Item("T3No")
        oInputT3StatusGrid.Columns.Item("T3No").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oInputT3StatusGrid.Columns.Item("CustomerCode")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
        oInputT3StatusGrid.Columns.Item("CustomerCode").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oInputT3StatusGrid.Columns.Item("CustomerName")
        oInputT3StatusGrid.Columns.Item("CustomerName").TitleObject.Sortable = True
        oColumn.Editable = False

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInputT3StatusGrid)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

    End Sub

    'karno generate T3  Tahap 2(2011.05.26 16:35:00)
    Private Sub GenerateT3Show(ByVal oForm As SAPbouiCOM.Form)

        Dim ARStatusQuery As String
        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

        Dim oMis_Utils As MIS_Utils


        oMis_Utils = New MIS_Utils

        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        If oForm.Items.Item("DateFrom").Specific.value = "" Then
            SBO_Application.MessageBox("Date From Must fill", 1, "OK")
            oForm.Items.Item("DateFrom").Click()
            Exit Sub
        ElseIf oForm.Items.Item("DateTo").Specific.value = "" Then
            SBO_Application.MessageBox("Date To Must fill", 1, "OK")
            oForm.Items.Item("DateTo").Click()
            Exit Sub
        ElseIf oForm.Items.Item("Customer").Specific.value = "" Then
            SBO_Application.MessageBox("Customer Must Fill", 1, "OK")
            oForm.Items.Item("Customer").Click()
            Exit Sub
        Else

            ARStatusQuery = "SELECT 'N' [AR Generate], T0.Docnum [Invoice No], T0.docentry [Invoice], T0.DocCur [Ccy],  " & _
                            "CASE WHEN (SELECT MainCurncy FROM DBO.OADM) = T0.DocCur THEN  T0.DocTotal ELSE T0.DocTotalFC END [Invoice Amount], " & _
                            "CASE WHEN ISNULL(T5.U_DocTotal, 0) <> 0 THEN (T5.U_DocTotal - T5.U_TotUM + T5.U_PPNDPP) " & _
                            "ELSE " & _
                            "	CASE WHEN LEFT(T0.CardCode, 2) = 'CP' THEN 0 " & _
                            "   ELSE T0.DocTotal End " & _
                            "END [FP Amount], T0.U_ProjectDesc [Project], T0.DocDate [Invoice Date], " & _
                            "T0.DocDueDate [Due Date], T1.U_WilayahCollect [Wilayah], T0.CardCode [Customer Code], T0.CardName [Customer Name], T0.Address [Address] From OINV T0 " & _
                            "INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode INNER JOIN OCTG T4 ON T0.GroupNum = T4.GroupNum " & _
                            "JOIN [@MIS_TAX] T5 ON T5.U_OINVDcNm = T0.DocNum " & _
                            "WHERE T0.DocStatus <> 'C' AND (T0.DocDate >= '" & oForm.Items.Item("DateFrom").Specific.value & "' AND T0.DocDate <= '" & oForm.Items.Item("DateTo").Specific.Value & "') " & _
                            "AND T0.CardCode Like '%" & oForm.Items.Item("Customer").Specific.Value & "%' " & _
                            "AND T0.DocEntry NOT IN (SELECT T3.U_OINVDocEntry From [@MIS_T3] T2 " & _
                            "INNER JOIN [@MIS_T3L] T3 ON T2.DocEntry = T3.DocEntry WHERE T3.U_T3LineStatus <> 'D' ) AND T4.ExtraDays > 0 ORDER BY T1.U_WilayahCollect, T0.CardCode, T0.DocNum "


        End If
        oForm.DataSources.DataTables.Item("GenT3StatusLst").ExecuteQuery(ARStatusQuery)
        oGenerateT3StatusGrid.DataTable = oForm.DataSources.DataTables.Item("GenT3StatusLst")

        oGenerateT3StatusGrid.Columns.Item("AR Generate").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oGenerateT3StatusGrid.Columns.Item("AR Generate").TitleObject.Sortable = True


        oColumn = oGenerateT3StatusGrid.Columns.Item("Invoice No")
        oGenerateT3StatusGrid.Columns.Item("Invoice No").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Ccy")
        oGenerateT3StatusGrid.Columns.Item("Ccy").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Invoice Amount")
        oGenerateT3StatusGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("FP Amount")
        oGenerateT3StatusGrid.Columns.Item("FP Amount").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Wilayah")
        oGenerateT3StatusGrid.Columns.Item("Wilayah").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Customer Code")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
        oGenerateT3StatusGrid.Columns.Item("Customer Code").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Customer Name")
        oGenerateT3StatusGrid.Columns.Item("Customer Name").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Project")
        oGenerateT3StatusGrid.Columns.Item("Project").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Invoice")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        oGenerateT3StatusGrid.Columns.Item("Invoice").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Invoice Date")
        oGenerateT3StatusGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Due Date")
        oGenerateT3StatusGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        oColumn.Editable = False

        oColumn = oGenerateT3StatusGrid.Columns.Item("Address")
        oGenerateT3StatusGrid.Columns.Item("Address").TitleObject.Sortable = True
        oColumn.Editable = False

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGenerateT3StatusGrid)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

    End Sub

    'Karno Delete T3 (2011.05.31 14:48)
    Private Sub DeleteT3Status(ByVal oForm As SAPbouiCOM.Form)

        Dim oDeleteT3StatusGrid As SAPbouiCOM.Grid

        Dim idx As Long

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralDataChild As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        oCmpSrv = oCompany.GetCompanyService

        oGeneralService = oCmpSrv.GetGeneralService("T3") ' UDO unique id
        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

        Dim oChild As SAPbobsCOM.GeneralData
        Dim vCompany As SAPbobsCOM.Company = Nothing
        Dim errConnect As String = ""

        Dim strQry As String = ""
        Dim StrQty As String = ""
        Dim oT3DocSeries As String = ""
        Dim oPdODocSeriesJasa As String = ""

        oDeleteT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        'GRID - Order by column checkbox
        oDeleteT3StatusGrid.Columns.Item("SelectAll").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


        Dim Temp As Integer
        For idx = oDeleteT3StatusGrid.Rows.Count - 1 To 0 Step -1
            If oDeleteT3StatusGrid.DataTable.GetValue(("SelectAll"), oDeleteT3StatusGrid.GetDataTableRowIndex(idx)) = "Y" Then
                Temp = Temp + 1
            End If
        Next

        If Temp = 0 Then
            SBO_Application.MessageBox("You Must Check One Invoice", 1, "OK")
        Else
            For idx = oDeleteT3StatusGrid.Rows.Count - 1 To 0 Step -1
                If oDeleteT3StatusGrid.DataTable.GetValue(("SelectAll"), oDeleteT3StatusGrid.GetDataTableRowIndex(idx)) = "Y" Then
                    SBO_Application.SetStatusBarMessage("Generating Delete T3.... Start !!! " & Temp - 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    If Not oCompany.InTransaction Then
                        oCompany.StartTransaction()
                    End If
                    Dim oRS As SAPbobsCOM.Recordset
                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    strQry = "SELECT T1.Docentry AS DocEntry, T1.LineId  FROM [@MIS_T3] T0 JOIN [@MIS_T3L] T1 ON T0.docEntry = T1.docentry " & _
                                " WHERE T1.U_OINVDocNum = '" & oDeleteT3StatusGrid.DataTable.GetValue(("Invoice No"), oDeleteT3StatusGrid.GetDataTableRowIndex(idx).ToString) & "' " & _
                                " AND T0.DocNum = '" & oForm.Items.Item("T3Number").Specific.Value & "' "

                    oRS.DoQuery(strQry)
                    Dim DocEntry As Integer
                    Dim LineId As Integer
                    DocEntry = oRS.Fields.Item("DocEntry").Value
                    LineId = oRS.Fields.Item("LineId").Value

                    oGeneralParams.SetProperty("DocEntry", DocEntry)
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                    oGeneralDataChild = oGeneralData.Child("MIS_T3L")

                    If oRS.RecordCount = Temp Then
                        oGeneralData.SetProperty("U_T3Status", "D")
                    End If

                    'For i = 0 To oRS.RecordCount - 1
                    oChild = oGeneralDataChild.Item(LineId - 1)
                    oChild.SetProperty("U_T3LineStatus", "D")
                    'Next

                    oGeneralService.Update(oGeneralData)
                    'End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                    oRS = Nothing

                    If lRetCode <> 0 Then
                        oCompany.GetLastError(lErrCode, sErrMsg)
                        SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)

                        If oCompany.InTransaction Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Else

                        If oCompany.InTransaction Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                    End If

        End If
            Next

            SBO_Application.SetStatusBarMessage("Generating Delete T3.... Finished !!! " & Temp - 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End If
        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

    Private Sub InputT3Status(ByVal oForm As SAPbouiCOM.Form)

        Dim oInputT3StatusGrid As SAPbouiCOM.Grid

        Dim idx As Long

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralDataChild As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        oCmpSrv = oCompany.GetCompanyService

        oGeneralService = oCmpSrv.GetGeneralService("T3") ' UDO unique id
        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

        Dim oChild As SAPbobsCOM.GeneralData
        Dim vCompany As SAPbobsCOM.Company = Nothing
        Dim errConnect As String = ""

        Dim strQry As String = ""
        Dim StrQty As String = ""
        Dim oT3DocSeries As String = ""
        Dim oPdODocSeriesJasa As String = ""

        oInputT3StatusGrid = oForm.Items.Item("myGridGen").Specific

        'GRID - Order by column checkbox
        oInputT3StatusGrid.Columns.Item("Receipt T3").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


            Dim Temp As Integer
            For idx = oInputT3StatusGrid.Rows.Count - 1 To 0 Step -1
                If oInputT3StatusGrid.DataTable.GetValue(("Receipt T3"), oInputT3StatusGrid.GetDataTableRowIndex(idx)) = "Y" Then
                    Temp = Temp + 1
                End If
            Next

            If Temp = 0 Then
                SBO_Application.MessageBox("You Must Check One Invoice", 1, "OK")
            Else
            For idx = oInputT3StatusGrid.Rows.Count - 1 To 0 Step -1
                If oInputT3StatusGrid.DataTable.GetValue(("Receipt T3"), oInputT3StatusGrid.GetDataTableRowIndex(idx)) = "Y" Then
                    SBO_Application.SetStatusBarMessage("Generating receipt T3.... Start !!! " & Temp - 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    If Not oCompany.InTransaction Then
                        oCompany.StartTransaction()
                    End If
                    Dim oRS As SAPbobsCOM.Recordset
                    Dim oRS1 As SAPbobsCOM.Recordset
                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRS1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    strQry = "SELECT T0.Docentry AS DocEntry FROM [@MIS_T3] T0" & _
                                " WHERE T0.DocNum = '" & oInputT3StatusGrid.DataTable.GetValue(("T3No"), oInputT3StatusGrid.GetDataTableRowIndex(idx).ToString) & "'"

                    oRS.DoQuery(strQry)
                    Dim DocEntry As Integer
                    DocEntry = oRS.Fields.Item("DocEntry").Value

                    oGeneralParams.SetProperty("DocEntry", DocEntry)
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                    oGeneralDataChild = oGeneralData.Child("MIS_T3L")

                    oGeneralData.SetProperty("U_T3TgldiTerima", Format(CDate(oForm.Items.Item("T3DtRcpt").Specific.string), "yyyy/MM/dd"))
                    oGeneralData.SetProperty("U_CollectID", oForm.Items.Item("ColCode").Specific.string)
                    oGeneralData.SetProperty("U_T3Status", "R")


                    StrQty = "SELECT T0.Docentry AS DocEntry FROM [@MIS_T3] T0 JOIN [@MIS_T3L] T1 ON T0.docentry = T1.docentry" & _
                             " WHERE T0.DocNum = '" & oInputT3StatusGrid.DataTable.GetValue(("T3No"), oInputT3StatusGrid.GetDataTableRowIndex(idx).ToString) & "'"

                    oRS1.DoQuery(StrQty)

                    For i = 0 To oRS1.RecordCount - 1
                        oChild = oGeneralDataChild.Item(i)
                        oChild.SetProperty("U_T3LineStatus", "R")
                    Next

                    oGeneralService.Update(oGeneralData)
                    'End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1)
                    oRS = Nothing
                    oRS1 = Nothing

                    If lRetCode <> 0 Then
                        oCompany.GetLastError(lErrCode, sErrMsg)
                        SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)

                        If oCompany.InTransaction Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Else

                        If oCompany.InTransaction Then
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                    End If

                End If
            Next

            SBO_Application.SetStatusBarMessage("Generating Receipt T3.... Finished !!! " & Temp - 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

    Private Sub GenerateT3Status(ByVal oForm As SAPbouiCOM.Form)

        Dim oGenerateT3StatusGrid As SAPbouiCOM.Grid

        Dim idx As Long

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralDataChild As SAPbobsCOM.GeneralDataCollection
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        oCmpSrv = oCompany.GetCompanyService

        Dim vCompany As SAPbobsCOM.Company = Nothing
        Dim errConnect As String = ""
        Dim oT3DocSeriesRec As SAPbobsCOM.Recordset

        Dim strQry As String = ""
        Dim oT3DocSeries As String = ""
        Dim oNextnumber As String = ""
        Dim Customer As String = ""

        oGenerateT3StatusGrid = oForm.Items.Item("myGridGen").Specific
        Customer = oForm.Items.Item("Customer").Specific.Value
        'GRID - Order by column checkbox
        oGenerateT3StatusGrid.Columns.Item("AR Generate").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oT3DocSeriesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


        If Left(Customer, 2) = "CP" Then
            strQry = "SELECT TOP 1 Series, NextNumber FROM NNM1 WHERE ObjectCode = 'T3' AND LEFT(SeriesName, 2) = '20' AND Indicator = YEAR(GETDATE()) ORDER BY Series desc "
        ElseIf Left(Customer, 2) = "CN" Then
            strQry = "SELECT TOP 1 Series, NextNumber FROM NNM1 WHERE ObjectCode = 'T3' AND LEFT(SeriesName, 2) = '10' AND Indicator = YEAR(GETDATE()) ORDER BY Series desc "
        Else
            strQry = "SELECT TOP 1 Series, NextNumber FROM NNM1 WHERE ObjectCode = 'T3' AND RIGHT(SeriesName, 2) = '11' AND Indicator = YEAR(GETDATE()) ORDER BY Series desc "
        End If

        oT3DocSeriesRec.DoQuery(strQry)
        '??? 
        If oT3DocSeriesRec.RecordCount <> 0 Then
            oT3DocSeries = oT3DocSeriesRec.Fields.Item("Series").Value
            oNextnumber = oT3DocSeriesRec.Fields.Item("NextNumber").Value
        Else
            MsgBox("T3 Document Series Kaca Order Tidak ada, Mohon Setup T3 Document Series!")
            Exit Sub
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oT3DocSeriesRec)
        oT3DocSeriesRec = Nothing
        GC.Collect()

        If oT3DocSeries <> "" Then
            oGeneralService = oCmpSrv.GetGeneralService("T3") ' UDO unique id

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            oGeneralData.SetProperty("Series", oT3DocSeries)

            Dim Temp As Integer
            For idx = oGenerateT3StatusGrid.Rows.Count - 1 To 0 Step -1
                If oGenerateT3StatusGrid.DataTable.GetValue(("AR Generate"), oGenerateT3StatusGrid.GetDataTableRowIndex(idx)) = "Y" Then
                    Temp = Temp + 1
                End If
            Next

            If Temp = 0 Then
                SBO_Application.MessageBox("You Must Check One Invoice", 1, "OK")
            Else

                For idx = oGenerateT3StatusGrid.Rows.Count - 1 To 0 Step -1
                    If oGenerateT3StatusGrid.DataTable.GetValue(("AR Generate"), oGenerateT3StatusGrid.GetDataTableRowIndex(idx)) = "Y" Then
                        SBO_Application.SetStatusBarMessage("Generating T3.... Start !!! " & Temp - 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                        If Not oCompany.InTransaction Then
                            oCompany.StartTransaction()
                        End If
                        Dim oRS As SAPbobsCOM.Recordset
                        oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'strQry = "SELECT T0.CardCode AS CustomerCode, T0.CardName AS CustomerName, T1.U_WilayahCollect AS Wilayah, T0.DocEntry AS Invoice, T0.DocCur AS Currency, T0.U_ProjectDesc AS Project, T0.DocRate AS Rate, " & _
                        '        " T0.DocNum AS InvoiceNo, T0.DocDate AS InvoiceDate, T0.DocDueDate AS DueDate, T0.DocTotal - T0.VatSum AS NetAmount, " & _
                        '        " T0.VatSum AS VatSum, T0.DocTotal AS DocTotal, ISNULL(T2.U_TaxDcDt,'') AS TaxDate, Convert(Varchar(19),T2.U_TaxNum) AS NomorPajak " & _
                        '        " FROM OINV T0 INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode LEFT JOIN [@MIS_TAX] T2 ON T0.DocEntry = T2.U_OinvDcEn " & _
                        '        " WHERE T0.DocEntry = '" & oGenerateT3StatusGrid.DataTable.GetValue(("Invoice"), oGenerateT3StatusGrid.GetDataTableRowIndex(idx).ToString) & "'"

                        If oGenerateT3StatusGrid.DataTable.GetValue(("FP Amount"), oGenerateT3StatusGrid.GetDataTableRowIndex(idx).ToString) = 0 Then
                            SBO_Application.MessageBox("Invoice#: " & oGenerateT3StatusGrid.DataTable.GetValue(("Invoice No"), oGenerateT3StatusGrid.GetDataTableRowIndex(idx).ToString) & " Amount = 0, tidak boleh digenerate!")

                        Else
                            strQry = "SELECT T0.CardCode AS CustomerCode, T0.CardName AS CustomerName, T1.U_WilayahCollect AS Wilayah, T0.DocEntry AS Invoice, T0.DocCur AS Currency, T0.U_ProjectDesc AS Project, T0.DocRate AS Rate, " & _
                                    " T0.DocNum AS InvoiceNo, T0.DocDate AS InvoiceDate, T0.DocDueDate AS DueDate, " & _
                                    " CASE" & _
                                    " WHEN ISNULL(T2.U_DocTotal, 0) <> 0 THEN " & _
                                    "   T2.U_DocTotal - T2.U_TotUM + T2.U_PPNDPP " & _
                                    " Else " & _
                                    "   CASE " & _
                                    "   WHEN LEFT(T0.CardCode, 2) = 'CP' THEN " & _
                                    "       ROUND(	" & _
                                    "           CASE " & _
                                    "           WHEN (SELECT MainCurncy FROM DBO.OADM) = T0.DocCur THEN " & _
                                    "               CASE " & _
                                    "               WHEN LEFT(T0.CardCode, 2) = 'CP' THEN " & _
                                    "                   (T0.DocTotal - T0.VatSum) " & _
                                    "               ELSE (T0.DocTotal - T0.VatSum) " & _
                                    "               End " & _
                                    "           Else " & _
                                    "               (T0.DocTotalFC - T0.VatSumFC) " & _
                                    "           END " & _
                                    "           + " & _
                                    "           CASE " & _
                                    "           WHEN (SELECT MainCurncy FROM DBO.OADM) = T0.DocCur THEN  " & _
                                    "               CASE " & _
                                    "               WHEN T1.U_isPungut = 'PUNGUT' THEN " & _
                                    "                   (T0.VatSum) " & _
                                    "               Else " & _
                                    "                   0 " & _
                                    "               End " & _
                                    "           Else " & _
                                    "               CASE " & _
                                    "               WHEN T1.U_isPungut = 'PUNGUT' THEN " & _
                                    "                   (T0.VatSumFC) " & _
                                    "               Else " & _
                                    "                   0 " & _
                                    "               End " & _
                                    "           END " & _
                                    "           , " & _
                                    "           CASE " & _
                                    "           WHEN (SELECT MainCurncy FROM DBO.OADM) = T0.DocCur THEN 0 " & _
                                    "           ELSE 3 " & _
                                    "           END) " & _
                                    "   ELSE " & _
                                    "   	CASE " & _
                                    "       WHEN (SELECT MainCurncy FROM DBO.OADM) = T0.DocCur THEN " & _
                                    "           T0.DocTotal " & _
                                    "       Else " & _
                                    "           T0.DocTotalFC " & _
                                    "       END " & _
                                    "   End " & _
                                    "END t3_from_fp_or_inv, " & _
                                    " T0.DocTotal - T0.VatSum AS NetAmount, " & _
                                    " T0.VatSum AS VatSum, T0.DocTotal AS DocTotal, ISNULL(T2.U_TaxDcDt,'') AS TaxDate, Convert(Varchar(19),T2.U_TaxNum) AS NomorPajak " & _
                                    " FROM OINV T0 INNER JOIN OCRD T1 ON T0.CardCode = T1.CardCode LEFT JOIN [@MIS_TAX] T2 ON T0.DocEntry = T2.U_OinvDcEn " & _
                                    " WHERE T0.DocEntry = '" & oGenerateT3StatusGrid.DataTable.GetValue(("Invoice"), oGenerateT3StatusGrid.GetDataTableRowIndex(idx).ToString) & "'"

                            oRS.DoQuery(strQry)

                            Dim NomorPajak As String
                            Dim CustomerCode As String
                            NomorPajak = oRS.Fields.Item("NomorPajak").Value
                            CustomerCode = oRS.Fields.Item("CustomerCode").Value

                            oGeneralData.SetProperty("U_DocDate", Format(CDate(oForm.Items.Item("T3Date").Specific.string), "yyyy/MM/dd"))

                            'oGeneralData.GetProperty(CustomerCode)
                            oGeneralData.SetProperty("U_KWIDocEntry", oNextnumber)
                            oGeneralData.SetProperty("U_CardCode", CustomerCode)
                            oGeneralData.SetProperty("U_CardName", oRS.Fields.Item("CustomerName").Value)
                            oGeneralData.SetProperty("U_WilayahCollect", oRS.Fields.Item("Wilayah").Value)
                            oGeneralData.SetProperty("U_T3Status", "A")

                            oGeneralDataChild = oGeneralData.Child("MIS_T3L")
                            oChild = oGeneralDataChild.Add
                            oChild.SetProperty("U_OINVDocEntry", oRS.Fields.Item("Invoice").Value)
                            oChild.SetProperty("U_OINVDocNum", oRS.Fields.Item("InvoiceNo").Value)
                            oChild.SetProperty("U_OINVDocCur", oRS.Fields.Item("Currency").Value)
                            oChild.SetProperty("U_OINVDocRate", oRS.Fields.Item("Rate").Value)
                            oChild.SetProperty("U_OINVProject", oRS.Fields.Item("Project").Value)
                            oChild.SetProperty("U_OINVDocDate", oRS.Fields.Item("InvoiceDate").Value)
                            oChild.SetProperty("U_OINVDocDueDate", oRS.Fields.Item("DueDate").Value)
                            oChild.SetProperty("U_TAXDueDate", oRS.Fields.Item("TaxDate").Value)
                            oChild.SetProperty("U_TAXTaxNum", NomorPajak)
                            oChild.SetProperty("U_OINVNetAmount", oRS.Fields.Item("NetAmount").Value)
                            oChild.SetProperty("U_OINVVATSum", oRS.Fields.Item("VatSum").Value)
                            'oChild.SetProperty("U_OINVDocTotal", oRS.Fields.Item("DocTotal").Value)
                            oChild.SetProperty("U_OINVDocTotal", oRS.Fields.Item("t3_from_fp_or_inv").Value)
                            oChild.SetProperty("U_T3LineStatus", "A")

                            'End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                            oRS = Nothing

                            If lRetCode <> 0 Then
                                oCompany.GetLastError(lErrCode, sErrMsg)
                                SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)

                                If oCompany.InTransaction Then
                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                            Else

                                If oCompany.InTransaction Then
                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If
                            End If
                        End If
                    End If

                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)
                SBO_Application.SetStatusBarMessage("Generating T3.... Finished !!! " & Temp - 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End If
        End If  ' Checking PdO Series

        Exit Sub


errHandler:
                MsgBox("Exception: " & Err.Description)
                Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

    Private Sub SetApplication()

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        SboGuiApi = New SAPbouiCOM.SboGuiApi

        sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

        Try
            SboGuiApi.Connect(sConnectionString)
            SBO_Application = SboGuiApi.GetApplication()
        Catch ex As Exception
            MsgBox("Make Sure That SAP Business One Application is running!!! ", MsgBoxStyle.Information)
            End
        End Try

    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String

        oCompany = New SAPbobsCOM.Company

        sCookie = oCompany.GetContextCookie

        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function


#Region "Pdc Input Class"

    Private Sub PdcInputAplicationMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        If pVal.BeforeAction = False Then
            Dim oForm As SAPbouiCOM.Form

            oForm = SBO_Application.Forms.ActiveForm
            If oForm.UniqueID = "Pdc_01" Then
                Select Case pVal.MenuUID
                    Case "1290" ' 1st Record
                        BindingPdcInput(oForm, BubbleEvent)
                    Case "1289" ' Prev Record
                        BindingPdcInput(oForm, BubbleEvent)
                    Case "1288" ' Next Record
                        BindingPdcInput(oForm, BubbleEvent)
                    Case "1291" ' Last Record
                        BindingPdcInput(oForm, BubbleEvent)

                    Case "1281" 'Find



                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            oForm.Items.Item("BtnSave").Visible = False
                            oForm.Items.Item("BtnShow").Visible = False
                            oForm.Items.Item("1").Visible = True
                            oForm.Items.Item("DocNum").Click()
                        ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            oForm.Items.Item("BtnSave").Visible = False
                            oForm.Items.Item("BtnShow").Visible = False
                            oForm.Items.Item("1").Visible = False
                        ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oForm.Items.Item("BtnSave").Visible = True
                            oForm.Items.Item("BtnShow").Visible = True
                            oForm.Items.Item("1").Visible = False
                        End If

                    Case "1282" 'Add
                        Dim oEditText As SAPbouiCOM.EditText
                        oEditText = oForm.Items.Item("Status").Specific
                        oEditText.DataBind.SetBound(False, "@MIS_PDC", "U_PDCStatus")

                        oForm.Items.Item("TotalCol").Specific.value = 0
                        oForm.Items.Item("BtnSave").Visible = True
                        oForm.Items.Item("BtnShow").Visible = True

                        oForm.Items.Item("Status").Specific.value = "Outstanding"
                        oForm.Items.Item("1").Visible = False
                        oForm.Items.Item("PdcDate").Specific.value = DateTime.Today.ToString("yyyyMMdd")
                        oForm.Items.Item("PdcBank").Click()


                End Select
            End If
        End If
    End Sub

    Private Sub BindingPdcInput(ByVal oForm As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)
        On Error GoTo Keluar
        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim PdcQuery As String
        Dim oEditText As SAPbouiCOM.EditText
        BubbleEvent = False
        oForm.Freeze(True)

        oForm.Items.Item("TotalCol").Specific.value = 0
        oEditText = oForm.Items.Item("Status").Specific
        oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCStatus")

        PdcGrid = oForm.Items.Item("Grid").Specific


        PdcQuery = "SELECT 'Y' as [Check] ,[@MIS_T3].DocNum as [T3 No.],[@MIS_PDCL].U_OINVDocEntry as [Invoice Doc Entry] ,[@MIS_PDCL].U_OINVDocNum as [Invoice No.]," & _
                    "[@MIS_T3L].U_OINVDocDate as [Invoice Date], [@MIS_T3L].U_OINVDocDueDate as [Due Date]," & _
                    "[@MIS_T3L].U_OINVDocTotal as [Invoice Amount], [@MIS_PDCL].U_CollectAmount as [Collection Amount] " & _
                    "FROM [@MIS_PDCL] LEFT OUTER JOIN " & _
                    "[@MIS_T3L] ON [@MIS_PDCL].U_T3DocEntry = [@MIS_T3L].DocEntry AND " & _
                    "[@MIS_PDCL].U_T3LineId = [@MIS_T3L].LineId LEFT OUTER JOIN " & _
                    "[@MIS_T3] ON [@MIS_PDCL].U_T3DocEntry = [@MIS_T3].DocEntry LEFT OUTER JOIN " & _
                    "[@MIS_PDC] ON [@MIS_PDCL].DocEntry = [@MIS_PDC].DocEntry " & _
                    "WHERE [@MIS_PDC].DocNum= '" & oForm.Items.Item("DocNum").Specific.value & "' order by [@MIS_PDCL].U_OINVDocEntry "


        ' Grid #: 1


        oForm.DataSources.DataTables.Item("PdcGLst")
        oForm.DataSources.DataTables.Item("PdcGLst").ExecuteQuery(PdcQuery)
        PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcGLst")

        Dim strstatus = oForm.Items.Item("Status").Specific.value

        oEditText = oForm.Items.Item("Status").Specific
        oEditText.DataBind.SetBound(False, "@MIS_PDC", "U_PDCStatus")

        If strstatus = "O" Then
            oForm.Items.Item("Status").Specific.value = "Outstanding"
        ElseIf strstatus = "C" Then
            oForm.Items.Item("Status").Specific.value = "Cair"
        ElseIf strstatus = "V" Then
            oForm.Items.Item("Status").Specific.value = "Void"
        End If
        oForm.Items.Item("T3No").Click()
        If blnFind = False Then
            oForm.Items.Item("PdcBank").Enabled = False
            oForm.Items.Item("PdcNo").Enabled = False
            oForm.Items.Item("PdcDate").Enabled = False
            oForm.Items.Item("PdcAmount").Enabled = False
            oForm.Items.Item("CustCode").Enabled = False
            ' oForm.Items.Item("T3No").Enabled = False
            oForm.Items.Item("Status").Enabled = False
            'oForm.Items.Item("BtnSave").Enabled = False
            'oForm.Items.Item("BtnShow").Enabled = False
            oForm.Items.Item("cmbSr").Enabled = False
            oForm.Items.Item("DocNum").Enabled = False
        End If


        PdcGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        PdcGrid.Columns.Item("Check").TitleObject.Sortable = False
        PdcGrid.Columns.Item("Check").Editable = False
        PdcGrid.Columns.Item("Check").Width = 80

        PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("T3 No.").Editable = False
        PdcGrid.Columns.Item("T3 No.").Width = 90

        PdcGrid.Columns.Item("Invoice No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice No.").Editable = False
        PdcGrid.Columns.Item("Invoice No.").Width = 100

        oColumn = PdcGrid.Columns.Item("Invoice Doc Entry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        PdcGrid.Columns.Item("Invoice Doc Entry").TitleObject.Sortable = True
        oColumn.Editable = False
        PdcGrid.Columns.Item("Invoice Doc Entry").Width = 100


        PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Date").Editable = False
        PdcGrid.Columns.Item("Invoice Date").Width = 100

        PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Due Date").Editable = False
        PdcGrid.Columns.Item("Due Date").Width = 100

        PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Amount").Editable = False
        PdcGrid.Columns.Item("Invoice Amount").Width = 130

        PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = False
        PdcGrid.Columns.Item("Collection Amount").Editable = False
        PdcGrid.Columns.Item("Collection Amount").Width = 130
        SumCollection(oForm)
Keluar:
        BubbleEvent = True
        oForm.Freeze(False)
        oEditText = Nothing
        oForm = Nothing
        oColumn = Nothing
        PdcGrid = Nothing
    End Sub

    Private Sub PdcFirstLoad()
        Dim oForm As SAPbouiCOM.Form
        Dim PdcQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim lNextSeriesNumOptimization As Long
        Dim PdcGrid As SAPbouiCOM.Grid
        Dim oCombobox As SAPbouiCOM.ComboBox = Nothing

        Try
            oForm = SBO_Application.Forms.Item("Pdc_01")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "Pdc"
            fcp.UniqueID = "Pdc_01"
            fcp.ObjectType = "PDC"
            fcp.XmlData = LoadFromXML("FormPdcInput.srf")

            oForm = SBO_Application.Forms.AddEx(fcp)
            oForm.Freeze(True)
            oForm.ClientHeight = 527
            oForm.DataSources.DBDataSources.Add("@MIS_PDC")

            oCombobox = oForm.Items.Item("cmbSr").Specific
            oCombobox.ValidValues.LoadSeries("PDC", SAPbouiCOM.BoSeriesMode.sf_Add)
            oCombobox.DataBind.SetBound(True, "@MIS_PDC", "SERIES")

            oEditText = oForm.Items.Item("DocNum").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "DocNum")

            lNextSeriesNumOptimization = oForm.BusinessObject.GetNextSerialNumber("SERIES")
            oEditText = oForm.Items.Item("DocNum").Specific
            oEditText.String = lNextSeriesNumOptimization



            oEditText = oForm.Items.Item("PdcBank").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCBankID")

            oEditText = oForm.Items.Item("BankNm").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCBankName")

            oEditText = oForm.Items.Item("PdcNo").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCNo")

            oEditText = oForm.Items.Item("PdcDate").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCDate")

            oEditText = oForm.Items.Item("PdcAmount").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCAmount")

            oEditText = oForm.Items.Item("CustCode").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_CardCode")

            oEditText = oForm.Items.Item("CustName").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_CardName")

            oEditText = oForm.Items.Item("Status").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_PDCStatus")

            oEditText = oForm.Items.Item("Status").Specific
            oEditText.DataBind.SetBound(False, "@MIS_PDC", "U_PDCStatus")

            oEditText = oForm.Items.Item("OutPay").Specific
            oEditText.DataBind.SetBound(True, "@MIS_PDC", "U_ORCTDocEntry")

            oForm.DataSources.UserDataSources.Add("TotalCol", SAPbouiCOM.BoDataType.dt_QUANTITY)
            oForm.DataSources.UserDataSources.Add("T3No", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)


            oEditText = oForm.Items.Item("T3No").Specific

            oForm.DataSources.UserDataSources.Item("TotalCol").Value = oMIS_Utils.fctFormatNumSBO(0, oCompany)
            oEditText = oForm.Items.Item("TotalCol").Specific
            oEditText.DataBind.SetBound(True, "", "TotalCol")

            oForm.Items.Item("Status").Specific.value = "Outstanding"
            oForm.Items.Item("PdcDate").Specific.value = DateTime.Today.ToString("yyyyMMdd")



            PdcGrid = oForm.Items.Item("Grid").Specific

            PdcGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto


            PdcQuery = "SELECT 'N' as [Check],[@MIS_T3].DocNum as [T3 No.],[@MIS_T3L].U_OINVDocEntry as [Invoice Doc Entry], [@MIS_T3L].U_OINVDocNum as [Invoice No.]," & _
                  "[@MIS_T3L].U_OINVDocDate as [Invoice Date]," & _
                  "[@MIS_T3L].U_OINVDocDueDate as [Due Date],[@MIS_T3L].U_OINVDocTotal as [Invoice Amount]," & _
                  "[@MIS_T3L].U_OINVDocTotal-[@MIS_T3L].U_OINVDocTotal as [Collection Amount], " & _
                  " 0 as [Oustanding Pdc Amount],0 as [Paid Amount] " & _
                  "FROM [@MIS_T3L] LEFT OUTER JOIN " & _
                   "[@MIS_PDCL] ON [@MIS_T3L].DocEntry = [@MIS_PDCL].U_T3DocEntry AND " & _
                   "[@MIS_T3L].LineId = [@MIS_PDCL].U_T3LineId LEFT OUTER JOIN " & _
                   "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry where 1=0 "


            ' Grid #: 1

            oForm.DataSources.DataTables.Add("PdcGLst")
            oForm.DataSources.DataTables.Item("PdcGLst").ExecuteQuery(PdcQuery)
            PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcGLst")


            PdcGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            PdcGrid.Columns.Item("Check").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Check").Width = 50
            PdcGrid.Columns.Item("Check").TitleObject.Caption = "Select All"

            PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
            PdcGrid.Columns.Item("T3 No.").Editable = False
            PdcGrid.Columns.Item("T3 No.").Width = 80

            PdcGrid.Columns.Item("Invoice No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Invoice No.").Editable = False
            PdcGrid.Columns.Item("Invoice No.").Width = 90

            oColumn = PdcGrid.Columns.Item("Invoice Doc Entry")
            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
            PdcGrid.Columns.Item("Invoice Doc Entry").TitleObject.Sortable = True
            oColumn.Editable = False
            PdcGrid.Columns.Item("Invoice Doc Entry").Width = 90


            PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Invoice Date").Editable = False
            PdcGrid.Columns.Item("Invoice Date").Width = 80

            PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Due Date").Editable = False
            PdcGrid.Columns.Item("Due Date").Width = 80

            PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Invoice Amount").Editable = False
            PdcGrid.Columns.Item("Invoice Amount").Width = 90

            PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Collection Amount").Width = 90


            PdcGrid.Columns.Item("Oustanding Pdc Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Oustanding Pdc Amount").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Oustanding Pdc Amount").Width = 90
            PdcGrid.Columns.Item("Oustanding Pdc Amount").Editable = False

            PdcGrid.Columns.Item("Paid Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Paid Amount").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Paid Amount").Width = 90
            PdcGrid.Columns.Item("Paid Amount").Editable = False

            oForm.DataBrowser.BrowseBy = "DocNum"


            oForm.Items.Item("1").Visible = False
            oForm.EnableMenu("1283", False)
            oForm.Items.Item("PdcBank").Click()
            oForm.Freeze(False)
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            PdcGrid = Nothing



            GC.Collect()


        End Try
        'oForm.Freeze(False)
        oForm.Visible = True
    End Sub

    Private Sub PdcEmpty(ByVal oForm As SAPbouiCOM.Form)

        oForm.Freeze(True)

        oForm.Items.Item("PdcBank").Specific.value = ""
        oForm.Items.Item("PdcNo").Specific.value = ""
        oForm.Items.Item("PdcAmount").Specific.value = 0
        oForm.Items.Item("TotalCol").Specific.value = 0
        oForm.Items.Item("CustCode").Specific.value = ""
        oForm.Items.Item("CustName").Specific.value = ""
        oForm.Items.Item("T3No").Specific.value = ""
        oForm.Items.Item("Status").Specific.value = ""

        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim PdcQuery As String

        PdcGrid = oForm.Items.Item("Grid").Specific


        PdcGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto



        PdcQuery = "SELECT 1 as [No.],[@MIS_T3].DocNum as [T3 No.],[@MIS_T3L].U_OINVDocNum as [Invoice No.]," & _
                   "[@MIS_T3L].U_OINVDocDate as [Invoice Date]," & _
                   "[@MIS_T3L].U_OINVDocDueDate as [Due Date],[@MIS_T3L].U_OINVDocTotal as [Invoice Amount],0 as [Collection Amount] " & _
                    "FROM [@MIS_T3L] LEFT OUTER JOIN " & _
                    "[@MIS_PDCL] ON [@MIS_T3L].DocEntry = [@MIS_PDCL].U_T3DocEntry AND " & _
                    "[@MIS_T3L].LineId = [@MIS_PDCL].U_T3LineId LEFT OUTER JOIN " & _
                    "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry where 1=0 "




        ' Grid #: 1

        oForm.DataSources.DataTables.Item("PdcGLst")
        oForm.DataSources.DataTables.Item("PdcGLst").ExecuteQuery(PdcQuery)
        PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcGLst")


        PdcGrid.Columns.Item("No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("No.").Width = 50

        PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("T3 No.").Editable = False
        PdcGrid.Columns.Item("T3 No.").Width = 120

        oColumn = PdcGrid.Columns.Item("Invoice No.")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
        oColumn.Editable = False
        PdcGrid.Columns.Item("Invoice No.").Width = 120


        PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Date").Editable = False
        PdcGrid.Columns.Item("Invoice Date").Width = 120

        PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Due Date").Editable = False
        PdcGrid.Columns.Item("Due Date").Width = 120

        PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Amount").Editable = False
        PdcGrid.Columns.Item("Invoice Amount").Width = 120

        PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Collection Amount").Width = 120


        oForm.Freeze(False)
        PdcGrid = Nothing
        oColumn = Nothing

    End Sub


    Private Sub PdcShow(ByVal oForm As SAPbouiCOM.Form)

        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim PdcQuery As String

        oForm.Freeze(True)
        If oForm.Items.Item("PdcBank").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Bank To Must fill", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcDate").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Date To Must fill", 1, "OK")
            oForm.Items.Item("PdcDate").Click()
            GoTo Keluar
        End If


        If oForm.Items.Item("PdcNo").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc No To Must fill", 1, "OK")
            oForm.Items.Item("PdcNo").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcAmount").Specific.value = "" Or CDec(oForm.Items.Item("PdcAmount").Specific.value) = 0 Then
            SBO_Application.MessageBox("Pdc Amount To Must fill", 1, "OK")
            oForm.Items.Item("PdcAmount").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("CustCode").Specific.value = "" Then
            SBO_Application.MessageBox("Customer Code To Must fill", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If

        oForm.Items.Item("TotalCol").Specific.value = 0
        oForm.Items.Item("Status").Specific.value = "Outstanding"

        PdcGrid = oForm.Items.Item("Grid").Specific

        'PdcQuery = "SELECT 'N' as [Check],[@MIS_T3].DocNum as [T3 No.],[@MIS_T3L].U_OINVDocNum as [Invoice No.]," & _
        '           "[@MIS_T3L].U_OINVDocDate as [Invoice Date]," & _
        '           "[@MIS_T3L].U_OINVDocDueDate as [Due Date],[@MIS_T3L].U_OINVDocTotal as [Invoice Amount],[@MIS_T3L].U_OINVDocTotal-[@MIS_T3L].U_OINVDocTotal as [Collection Amount], " & _
        '           "[@MIS_T3L].U_OINVDocEntry as [Invoice Doc Entry],[@MIS_T3L].LineId as [T3 Line No],[@MIS_T3].DocEntry as [T3 DocEntry]  " & _
        '           "FROM [@MIS_T3L] LEFT OUTER JOIN " & _
        '            "[@MIS_PDCL] ON [@MIS_T3L].DocEntry = [@MIS_PDCL].U_T3DocEntry AND " & _
        '            "[@MIS_T3L].LineId = [@MIS_PDCL].U_T3LineId LEFT OUTER JOIN " & _
        '            "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry where [@MIS_T3L].U_T3LineStatus='R' And " & _
        '            "[@MIS_PDCL].DocEntry IS NULL AND [@MIS_T3].U_CardCode='" & oForm.Items.Item("CustCode").Specific.value & "'"

        'PdcQuery = "SELECT 'N' AS [Check], [@MIS_T3].DocNum AS [T3 No.],[@MIS_T3L].U_OINVDocNum AS [Invoice No.], " & _
        '            "[@MIS_T3L].U_OINVDocDate AS [Invoice Date],[@MIS_T3L].U_OINVDocDueDate AS [Due Date]," & _
        '            "[@MIS_T3L].U_OINVDocTotal AS [Invoice Amount], " & _
        '            "[@MIS_T3L].U_OINVDocTotal - [@MIS_T3L].U_OINVDocTotal AS [Collection Amount], " & _
        '            "[@MIS_T3L].U_OINVDocEntry AS [Invoice Doc Entry],[@MIS_T3L].LineId AS [T3 Line No], " & _
        '            "[@MIS_T3].DocEntry AS [T3 DocEntry], " & _
        '            "CASE VW_PDCL.U_InvPaidStatus WHEN 'V' THEN 0 ELSE isnull(VW_PDCL.U_CollectAmount,0) END AS [Oustanding Pdc Amount], " & _
        '            "OINV.PaidSum AS [Paid Amount], " & _
        '            "[@MIS_T3L].U_OINVDocTotal-(CASE VW_PDCL.U_InvPaidStatus WHEN 'V' THEN 0 ELSE isnull(VW_PDCL.U_CollectAmount,0) END)-OINV.PaidSum as [Max Collection] " & _
        '            "FROM [@MIS_PDC] RIGHT OUTER JOIN " & _
        '            "VW_PDCL ON [@MIS_PDC].DocEntry = VW_PDCL.DocEntry RIGHT OUTER JOIN " & _
        '            "[@MIS_T3L] LEFT OUTER JOIN " & _
        '            "OINV ON [@MIS_T3L].U_OINVDocEntry = OINV.DocEntry ON VW_PDCL.U_T3DocEntry = [@MIS_T3L].DocEntry AND " & _
        '            "VW_PDCL.U_T3LineId = [@MIS_T3L].LineId LEFT OUTER JOIN " & _
        '            "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry  " & _
        '            "WHERE ([@MIS_T3L].U_T3LineStatus = 'R')  " & _
        '            "AND isnull([@MIS_PDC].U_PDCStatus,'') <>'C' " & _
        '            "AND isnull([@MIS_T3L].U_T3PDCStatus,'') <>'F' " & _
        '            "AND [@MIS_T3].U_CardCode='" & oForm.Items.Item("CustCode").Specific.value & "'" & _
        '            "AND ([@MIS_T3L].U_OINVDocTotal-(CASE VW_PDCL.U_InvPaidStatus WHEN 'V' THEN 0 ELSE isnull(VW_PDCL.U_CollectAmount,0) END)-OINV.PaidSum)>0"

        'PdcQuery = "With TPDCL ( " & _
        '            "DocEntry, LineId, U_T3DocEntry, U_T3LineId, U_OINVDocEntry, U_OINVDocNum, " & _
        '            "U_CollectAmount) as " & _
        '            "(SELECT [@MIS_PDCL].DocEntry, [@MIS_PDCL].LineId, [@MIS_PDCL].U_T3DocEntry, [@MIS_PDCL].U_T3LineId," & _
        '            "[@MIS_PDCL].U_OINVDocEntry, [@MIS_PDCL].U_OINVDocNum,SUM([@MIS_PDCL].U_CollectAmount) " & _
        '            "FROM [@MIS_PDCL] LEFT OUTER JOIN [@MIS_PDC] ON [@MIS_PDCL].DocEntry = [@MIS_PDC].DocEntry " & _
        '            "WHERE [@MIS_PDC].U_PDCStatus = 'O' " & _
        '            "group by [@MIS_PDCL].DocEntry, [@MIS_PDCL].LineId, [@MIS_PDCL].U_T3DocEntry, [@MIS_PDCL].U_T3LineId," & _
        '            "[@MIS_PDCL].U_OINVDocEntry, [@MIS_PDCL].U_OINVDocNum)  " & _
        '            "SELECT 'N' AS [Check], [@MIS_T3].DocNum AS [T3 No.],[@MIS_T3L].U_OINVDocNum AS [Invoice No.],  " & _
        '            "[@MIS_T3L].U_OINVDocDate AS [Invoice Date],[@MIS_T3L].U_OINVDocDueDate AS [Due Date], " & _
        '            "[@MIS_T3L].U_OINVDocTotal AS [Invoice Amount],  " & _
        '            "[@MIS_T3L].U_OINVDocTotal - [@MIS_T3L].U_OINVDocTotal AS [Collection Amount],  " & _
        '            "[@MIS_T3L].U_OINVDocEntry AS [Invoice Doc Entry],[@MIS_T3L].LineId AS [T3 Line No],  " & _
        '            "[@MIS_T3].DocEntry AS [T3 DocEntry],  " & _
        '            "isnull(TPDCL.U_CollectAmount,0) AS [Oustanding Pdc Amount],  " & _
        '            "OINV.PaidSum AS [Paid Amount],  " & _
        '            "[@MIS_T3L].U_OINVDocTotal-isnull(TPDCL.U_CollectAmount,0)-OINV.PaidSum as [Max Collection]  " & _
        '            "FROM [@MIS_PDC] RIGHT OUTER JOIN  " & _
        '            "TPDCL ON [@MIS_PDC].DocEntry = TPDCL.DocEntry RIGHT OUTER JOIN  " & _
        '            "[@MIS_T3L] LEFT OUTER JOIN  " & _
        '            "OINV ON [@MIS_T3L].U_OINVDocEntry = OINV.DocEntry ON TPDCL.U_T3DocEntry = [@MIS_T3L].DocEntry AND  " & _
        '            "TPDCL.U_T3LineId = [@MIS_T3L].LineId LEFT OUTER JOIN  " & _
        '            "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry   " & _
        '            "WHERE ([@MIS_T3L].U_T3LineStatus = 'R')   " & _
        '            "AND [@MIS_T3].U_CardCode='" & oForm.Items.Item("CustCode").Specific.value & "'" & _
        '            "AND ([@MIS_T3L].U_OINVDocTotal- isnull(TPDCL.U_CollectAmount,0)-OINV.PaidSum)>0 "

        PdcQuery = "With TPDCL ( " & _
                    "U_T3DocEntry, U_T3LineId, U_OINVDocEntry, U_OINVDocNum, " & _
                    "U_CollectAmount) as " & _
                    "(SELECT [@MIS_PDCL].U_T3DocEntry, [@MIS_PDCL].U_T3LineId," & _
                    "[@MIS_PDCL].U_OINVDocEntry, [@MIS_PDCL].U_OINVDocNum,SUM([@MIS_PDCL].U_CollectAmount) " & _
                    "FROM [@MIS_PDCL] LEFT OUTER JOIN [@MIS_PDC] ON [@MIS_PDCL].DocEntry = [@MIS_PDC].DocEntry " & _
                    "WHERE [@MIS_PDC].U_PDCStatus = 'O' " & _
                    "group by [@MIS_PDCL].U_T3DocEntry, [@MIS_PDCL].U_T3LineId," & _
                    "[@MIS_PDCL].U_OINVDocEntry, [@MIS_PDCL].U_OINVDocNum)  " & _
                    "SELECT 'N' AS [Check], [@MIS_T3].DocNum AS [T3 No.],[@MIS_T3L].U_OINVDocEntry AS [Invoice Doc Entry],[@MIS_T3L].U_OINVDocNum AS [Invoice No.],  " & _
                    "[@MIS_T3L].U_OINVDocDate AS [Invoice Date],[@MIS_T3L].U_OINVDocDueDate AS [Due Date], " & _
                    "[@MIS_T3L].U_OINVDocTotal AS [Invoice Amount],  " & _
                    "[@MIS_T3L].U_OINVDocTotal - [@MIS_T3L].U_OINVDocTotal AS [Collection Amount],  " & _
                    "[@MIS_T3L].LineId AS [T3 Line No],  " & _
                    "[@MIS_T3].DocEntry AS [T3 DocEntry],  " & _
                    "isnull(TPDCL.U_CollectAmount,0) AS [Oustanding Pdc Amount],  " & _
                    "OINV.PaidSum AS [Paid Amount],  " & _
                    "[@MIS_T3L].U_OINVDocTotal-isnull(TPDCL.U_CollectAmount,0)-OINV.PaidSum as [Max Collection]  " & _
                    "FROM [@MIS_T3L] LEFT OUTER JOIN  " & _
                    "OINV ON [@MIS_T3L].U_OINVDocNum =OINV.DocNum LEFT OUTER JOIN  " & _
                    "TPDCL ON [@MIS_T3L].DocEntry =TPDCL.U_T3DocEntry AND   " & _
                    "[@MIS_T3L].LineId =TPDCL.U_T3LineId LEFT OUTER JOIN  " & _
                    "[@MIS_T3] ON [@MIS_T3L].DocEntry =[@MIS_T3].DocEntry  " & _
                    "WHERE ([@MIS_T3L].U_T3LineStatus = 'R')   " & _
                    "AND [@MIS_T3].U_CardCode='" & oForm.Items.Item("CustCode").Specific.value & "'" & _
                    "AND ([@MIS_T3L].U_OINVDocTotal- isnull(TPDCL.U_CollectAmount,0)-OINV.PaidSum)>0 "

        If oForm.Items.Item("T3No").Specific.value <> "" Then
            PdcQuery = PdcQuery + " AND [@MIS_T3].DocNum='" & oForm.Items.Item("T3No").Specific.value & "'"
        End If

        PdcQuery = PdcQuery + " order by [@MIS_T3L].U_OINVDocEntry"


        ' Grid #: 1


        oForm.DataSources.DataTables.Item("PdcGLst")
        oForm.DataSources.DataTables.Item("PdcGLst").ExecuteQuery(PdcQuery)
        PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcGLst")



        PdcGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        'PdcGrid.Columns.Item("Check").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Check").Width = 50
        PdcGrid.Columns.Item("Check").TitleObject.Caption = "Select All"

        PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("T3 No.").Editable = False
        PdcGrid.Columns.Item("T3 No.").Width = 70

        PdcGrid.Columns.Item("Invoice No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice No.").Editable = False
        PdcGrid.Columns.Item("Invoice No.").Width = 90

        oColumn = PdcGrid.Columns.Item("Invoice Doc Entry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        PdcGrid.Columns.Item("Invoice Doc Entry").TitleObject.Sortable = True
        oColumn.Editable = False
        PdcGrid.Columns.Item("Invoice Doc Entry").Width = 90


        PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Date").Editable = False
        PdcGrid.Columns.Item("Invoice Date").Width = 80

        PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Due Date").Editable = False
        PdcGrid.Columns.Item("Due Date").Width = 80

        PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Amount").Editable = False
        PdcGrid.Columns.Item("Invoice Amount").Width = 90

        PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Collection Amount").Width = 90

        PdcGrid.Columns.Item("Oustanding Pdc Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Oustanding Pdc Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Oustanding Pdc Amount").Width = 120
        PdcGrid.Columns.Item("Oustanding Pdc Amount").Editable = False

        PdcGrid.Columns.Item("Paid Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Paid Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Paid Amount").Width = 90
        PdcGrid.Columns.Item("Paid Amount").Editable = False

        PdcGrid.Columns.Item("T3 DocEntry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 DocEntry").Visible = False
        PdcGrid.Columns.Item("T3 DocEntry").Width = 50

        PdcGrid.Columns.Item("T3 Line No").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 Line No").Visible = False
        PdcGrid.Columns.Item("T3 Line No").Width = 50

        PdcGrid.Columns.Item("Max Collection").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Max Collection").Visible = False
        PdcGrid.Columns.Item("Max Collection").Width = 50



Keluar:
        oForm.Freeze(False)
        oForm = Nothing
        oColumn = Nothing
        PdcGrid = Nothing


    End Sub

    Private Sub PdcSave(ByVal oForm As SAPbouiCOM.Form)
        On Error GoTo Keluar
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralDataChild As SAPbobsCOM.GeneralDataCollection
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim strQuery As String
        Dim strseries As String
        Dim i As Integer
        Dim CheckInt As Integer = 0
        Dim prmAmount As Decimal = 0
        Dim invstatus As String = String.Empty
        Dim intMsgBox As Integer

        oCmpSrv = oCompany.GetCompanyService

        PdcGrid = oForm.Items.Item("Grid").Specific


        If oForm.Items.Item("PdcBank").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Bank To Must fill", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("BankNm").Specific.value = "" Then
            SBO_Application.MessageBox("Bank Name To Must fill", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcDate").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Date To Must fill", 1, "OK")
            oForm.Items.Item("PdcDate").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcDate").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Date To Must fill", 1, "OK")
            oForm.Items.Item("PdcDate").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcNo").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc No To Must fill", 1, "OK")
            oForm.Items.Item("PdcNo").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcAmount").Specific.value = "" Or CDec(oForm.Items.Item("PdcAmount").Specific.value) = 0 Then
            SBO_Application.MessageBox("Pdc Amount To Must fill", 1, "OK")
            oForm.Items.Item("PdcAmount").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("CustCode").Specific.value = "" Then
            SBO_Application.MessageBox("Customer Code To Must fill", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("CustName").Specific.value = "" Then
            SBO_Application.MessageBox("Customer Name To Must fill, Please Chooise Customer Code", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If


        For i = 0 To PdcGrid.Rows.Count - 1
            If PdcGrid.DataTable.GetValue(("Check"), PdcGrid.GetDataTableRowIndex(i)) = "Y" Then
                CheckInt = CheckInt + 1
            End If
        Next

        If CheckInt = 0 Then
            SBO_Application.MessageBox("Please select the data that will be Save ", 1, "OK")
            GoTo Keluar
        End If

        For i = 0 To PdcGrid.Rows.Count - 1
            If PdcGrid.DataTable.GetValue(("Check"), PdcGrid.GetDataTableRowIndex(i)) = "Y" Then
                prmAmount = PdcGrid.DataTable.GetValue(("Max Collection"), PdcGrid.GetDataTableRowIndex(i))
                If prmAmount < CDec(PdcGrid.DataTable.GetValue(("Collection Amount"), PdcGrid.GetDataTableRowIndex(i))) Then
                    SBO_Application.MessageBox("Invoice No " + PdcGrid.DataTable.GetValue(("Invoice No."), PdcGrid.GetDataTableRowIndex(i)).ToString + _
                                                            " Collection Amount max is " + prmAmount.ToString, 1, "OK")
                    GoTo Keluar
                End If
            End If
        Next


        SumCollection(oForm)
        If CDec(oForm.Items.Item("TotalCol").Specific.value) > CDec(oForm.Items.Item("PdcAmount").Specific.value) Then
            SBO_Application.MessageBox("Total Collection is greater than the Pdc Amount ", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        ElseIf CDec(oForm.Items.Item("TotalCol").Specific.value) < CDec(oForm.Items.Item("PdcAmount").Specific.value) Then
            intMsgBox = SBO_Application.MessageBox("Total Collection is Less than the Pdc Amount,are you sure save this Pdc No ?", 2, "Yes", "No")
            If intMsgBox = 2 Then GoTo Keluar
        End If

        oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strQuery = "SELECT TOP 1 U_CardName FROM [@MIS_T3] WHERE U_CardCode = '" & oForm.Items.Item("CustCode").Specific.value & "'"

        oRecordset.DoQuery(strQuery)
        If oRecordset.RecordCount < 1 Then
            SBO_Application.MessageBox("Customer Code Not Found ", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If

        strQuery = "SELECT U_BankID FROM [@BANKGIRO] WHERE U_BankID = '" & oForm.Items.Item("PdcBank").Specific.value & "'"

        oRecordset.DoQuery(strQuery)
        If oRecordset.RecordCount < 1 Then
            SBO_Application.MessageBox("Bank Code Not Found ", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If


        strQuery = "SELECT TOP 1 * FROM [@MIS_PDC] WHERE U_PDCBankID='" & oForm.Items.Item("PdcBank").Specific.value & "'" & _
                   "AND U_PDCNo='" & oForm.Items.Item("PdcNo").Specific.value & "' and U_PDCStatus<>'V'"
        oRecordset.DoQuery(strQuery)

        If oRecordset.RecordCount <> 0 Then
            SBO_Application.MessageBox("PDC No " + oForm.Items.Item("PdcNo").Specific.value.ToString.Trim + " From Bank Code " + oForm.Items.Item("PdcBank").Specific.value.ToString.Trim + " Already Existing", 1, "OK")
            GoTo Keluar
        End If


        'strQuery = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = 'PDC' AND Indicator = YEAR(GETDATE()) "
        Dim cmbSeries As SAPbouiCOM.ComboBox = Nothing
        cmbSeries = oForm.Items.Item("cmbSr").Specific
        strseries = cmbSeries.Value



        If strseries = "" Then
            'strQuery = "select top 1 DfltSeries from onnm where ObjectCode='PDC' and DfltSeries<>0 "

            strQuery = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = 'PDC' AND RIGHT(SeriesName, 2) = '11' AND Indicator = YEAR(GETDATE()) "
            oRecordset.DoQuery(strQuery)

            If oRecordset.RecordCount <> 0 Then
                strseries = oRecordset.Fields.Item("Series").Value
            Else
                SBO_Application.MessageBox("PDC Document Series Default Not Found, Please Setup Default PDC Document Series! OR Please select Series Doc.", 1, "OK")
                GoTo Keluar
            End If
        End If



        If Not oCompany.InTransaction Then
            oCompany.StartTransaction()
        End If

        oGeneralService = oCmpSrv.GetGeneralService("PDC")

        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

        oGeneralData.SetProperty("Series", strseries)
        oGeneralData.SetProperty("U_PDCBankID", oForm.Items.Item("PdcBank").Specific.value)
        oGeneralData.SetProperty("U_PDCBankName", oForm.Items.Item("BankNm").Specific.value)
        oGeneralData.SetProperty("U_PDCNo", oForm.Items.Item("PdcNo").Specific.value)
        Dim strdate As String
        strdate = Left(oForm.Items.Item("PdcDate").Specific.value.ToString, 4) + "/" + _
                  Mid(oForm.Items.Item("PdcDate").Specific.value.ToString, 5, 2) + "/" + _
                  Right(oForm.Items.Item("PdcDate").Specific.value.ToString, 2)

        oGeneralData.SetProperty("U_PDCDate", strdate)
        oGeneralData.SetProperty("U_PDCAmount", oForm.Items.Item("PdcAmount").Specific.value)
        oGeneralData.SetProperty("U_CardCode", oForm.Items.Item("CustCode").Specific.value)
        oGeneralData.SetProperty("U_CardName", oForm.Items.Item("CustName").Specific.value)
        oGeneralData.SetProperty("U_PDCStatus", Left(oForm.Items.Item("Status").Specific.value, 1))

        oGeneralDataChild = oGeneralData.Child("MIS_PDCL")


        For i = 0 To PdcGrid.Rows.Count - 1
            If PdcGrid.DataTable.GetValue(("Check"), PdcGrid.GetDataTableRowIndex(i)) = "Y" Then
                oChild = oGeneralDataChild.Add
                oChild.SetProperty("U_T3DocEntry", PdcGrid.DataTable.GetValue(("T3 DocEntry"), PdcGrid.GetDataTableRowIndex(i).ToString))
                oChild.SetProperty("U_T3LineId", PdcGrid.DataTable.GetValue(("T3 Line No"), PdcGrid.GetDataTableRowIndex(i).ToString))
                oChild.SetProperty("U_OINVDocEntry", PdcGrid.DataTable.GetValue(("Invoice Doc Entry"), PdcGrid.GetDataTableRowIndex(i).ToString))
                oChild.SetProperty("U_OINVDocNum", PdcGrid.DataTable.GetValue(("Invoice No."), PdcGrid.GetDataTableRowIndex(i).ToString))
                oChild.SetProperty("U_CollectAmount", PdcGrid.DataTable.GetValue(("Collection Amount"), PdcGrid.GetDataTableRowIndex(i).ToString))
                If CDec(PdcGrid.DataTable.GetValue(("Collection Amount"), PdcGrid.GetDataTableRowIndex(i).ToString)) = CDec(PdcGrid.DataTable.GetValue(("Invoice Amount"), PdcGrid.GetDataTableRowIndex(i).ToString)) Then
                    invstatus = "F"
                ElseIf CDec(PdcGrid.DataTable.GetValue(("Collection Amount"), PdcGrid.GetDataTableRowIndex(i).ToString)) < CDec(PdcGrid.DataTable.GetValue(("Invoice Amount"), PdcGrid.GetDataTableRowIndex(i).ToString)) Then
                    invstatus = "H"
                End If

                oChild.SetProperty("U_InvPaidStatus", invstatus)

                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                    GoTo Keluar
                    'Else
                    'strQuery = "Update [@MIS_T3L] set U_T3PDCStatus='" & invstatus & "' " & _
                    '           "where DocEntry='" & PdcGrid.DataTable.GetValue(("T3 DocEntry"), PdcGrid.GetDataTableRowIndex(i).ToString) & "'" & _
                    '           "and LineId='" & PdcGrid.DataTable.GetValue(("T3 Line No"), PdcGrid.GetDataTableRowIndex(i).ToString) & "'"
                    'oRecordset.DoQuery(strQuery)
                End If

            End If
        Next

        oGeneralParams = oGeneralService.Add(oGeneralData)

        If oCompany.InTransaction Then
            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If

        PdcEmpty(oForm)
        oForm.Items.Item("PdcBank").Click()

Keluar:
        If Err.Description <> "" Then
            SBO_Application.MessageBox("Exception: " & Err.Description, 1, "OK")
        End If

        If lRetCode <> 0 Or Err.Description <> "" Then
            If oCompany.InTransaction Then
                Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        End If

        oGeneralService = Nothing
        oGeneralData = Nothing
        oGeneralDataChild = Nothing
        oChild = Nothing
        oGeneralParams = Nothing
        oCmpSrv = Nothing
        PdcGrid = Nothing
        oColumn = Nothing
        oRecordset = Nothing

    End Sub

    Private Sub SumCollection(ByVal oForm As SAPbouiCOM.Form)
        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim DecCollection As Decimal = 0

        PdcGrid = oForm.Items.Item("Grid").Specific

        Dim i As Integer
        For i = 0 To PdcGrid.Rows.Count - 1
            If PdcGrid.DataTable.GetValue(("Check"), i) = "Y" Then
                DecCollection = DecCollection + CDec(PdcGrid.DataTable.GetValue(("Collection Amount"), i))
            End If
        Next

        oForm.Items.Item("TotalCol").Specific.value = DecCollection
        PdcGrid = Nothing
    End Sub

    Private Sub PdcInputAplicationItem(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        If pVal.ActionSuccess = True Then
            If pVal.FormTypeEx = "Pdc" Then
                Dim oForm As SAPbouiCOM.Form = Nothing

                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    oForm = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" Then
                            If oForm.Items.Item("1").Specific.caption = "Find" Then
                                blnFind = True
                                oForm.Items.Item("1").Visible = False
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            oForm.Items.Item("1").Visible = False
                        End If
                        'Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        '    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        '        oForm.Items.Item("1").Visible = False

                        '    End If
                End Select




            End If
        End If

        If pVal.Before_Action = True Then
            If pVal.FormTypeEx = "Pdc" Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    oForm = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                SBO_Application.MessageBox("Data Can't Edit", 1, "OK")
                                oForm.Items.Item("1").Visible = False
                                BubbleEvent = False
                                ' Exit Sub
                            ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                SBO_Application.MessageBox("Data Can't Save", 1, "OK")
                                oForm.Items.Item("1").Visible = False
                                BubbleEvent = False
                                'Exit Sub
                            End If
                        End If
                End Select

            End If


        End If


        If pVal.BeforeAction = False Then

            If pVal.FormTypeEx = "Pdc" Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    oForm = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                        If (pVal.ItemUID = "PdcBank") Then
                            If blnFind = True Then
                                BindingPdcInput(oForm, BubbleEvent)
                                blnFind = False
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        If (pVal.ItemUID = "T3No") Then
                            If oForm.Items.Item("T3No").Specific.value <> "" Then
                                If IsNumeric(oForm.Items.Item("T3No").Specific.value) = False Then
                                    SBO_Application.MessageBox("T3 No Must be Numeric", 1, "OK")
                                    oForm.Items.Item("T3No").Click()
                                End If
                            End If
                        End If
                        

                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If (pVal.ColUID = "Collection Amount") Then

                            Dim PdcGrid As SAPbouiCOM.Grid = Nothing
                            Dim prmAmount As Decimal

                            oForm = SBO_Application.Forms.Item(FormUID)
                            oForm.Freeze(True)

                            PdcGrid = oForm.Items.Item("Grid").Specific

                            If PdcGrid.DataTable.GetValue(("Check"), pVal.Row) = "Y" Then
                                prmAmount = PdcGrid.DataTable.GetValue(("Max Collection"), pVal.Row)

                                If prmAmount < CDec(PdcGrid.DataTable.GetValue(("Collection Amount"), pVal.Row)) Then
                                    SBO_Application.MessageBox("Invoice No " + PdcGrid.DataTable.GetValue(("Invoice No."), pVal.Row).ToString + _
                                                              " Collection Amount max is " + prmAmount.ToString, 1, "OK")
                                    PdcGrid.DataTable.SetValue("Collection Amount", pVal.Row, prmAmount.ToString)
                                End If
                            ElseIf PdcGrid.DataTable.GetValue(("Check"), pVal.Row) = "N" Then
                                PdcGrid.DataTable.SetValue("Collection Amount", pVal.Row, "0")
                            End If

                            
                            SumCollection(oForm)
                            oForm.Freeze(False)
                            PdcGrid = Nothing


                        End If

                        If pVal.ItemUID = "CustCode" Then
                            Dim oRecordset As SAPbobsCOM.Recordset
                            Dim strQuery As String
                            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            strQuery = "SELECT TOP 1 U_CardName FROM [@MIS_T3] WHERE U_CardCode = '" & oForm.Items.Item("CustCode").Specific.value & "'"

                            oRecordset.DoQuery(strQuery)
                            If oRecordset.RecordCount > 0 Then
                                oForm.Items.Item("CustName").Specific.value = oRecordset.Fields.Item("U_CardName").Value
                            Else
                                oForm.Items.Item("CustName").Specific.value = ""
                            End If

                            oRecordset = Nothing
                        End If

                        If pVal.ItemUID = "PdcBank" Then
                            Dim oRecordset As SAPbobsCOM.Recordset
                            Dim strQuery As String
                            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            strQuery = "SELECT TOP 1 U_BankName FROM [@BANKGIRO] WHERE U_BankID = '" & oForm.Items.Item("PdcBank").Specific.value & "'"

                            oRecordset.DoQuery(strQuery)
                            If oRecordset.RecordCount > 0 Then
                                oForm.Items.Item("BankNm").Specific.value = oRecordset.Fields.Item("U_BankName").Value
                            Else
                                oForm.Items.Item("BankNm").Specific.value = ""
                            End If

                            oRecordset = Nothing
                        End If


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If pVal.ItemUID = "BtnShow" Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            PdcShow(oForm)
                        ElseIf pVal.ItemUID = "BtnSave" Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            PdcSave(oForm)
                        ElseIf pVal.ItemUID = "BtnCancel" Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            oForm.Close()
                        End If
                        If pVal.ItemUID = "CustCode" Then
                            Dim oRecordset As SAPbobsCOM.Recordset
                            Dim strQuery As String
                            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            strQuery = "SELECT TOP 1 U_CardName FROM [@MIS_T3] WHERE U_CardCode = '" & oForm.Items.Item("CustCode").Specific.value & "'"

                            oRecordset.DoQuery(strQuery)
                            If oRecordset.RecordCount > 0 Then
                                oForm.Items.Item("CustName").Specific.value = oRecordset.Fields.Item("U_CardName").Value
                            Else
                                oForm.Items.Item("CustName").Specific.value = ""
                            End If

                            oRecordset = Nothing
                        End If



                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.ColUID = "Check") And pVal.Row > -1 Then
                            Dim PdcGrid As SAPbouiCOM.Grid = Nothing
                            Dim prmAmount As Decimal

                            oForm = SBO_Application.Forms.Item(FormUID)
                            oForm.Freeze(True)

                            PdcGrid = oForm.Items.Item("Grid").Specific
                            If PdcGrid.Columns.Item("Check").Editable = True Then
                                If PdcGrid.DataTable.GetValue(("Check"), pVal.Row) = "Y" Then
                                    prmAmount = PdcGrid.DataTable.GetValue(("Max Collection"), pVal.Row)
                                    PdcGrid.DataTable.SetValue("Collection Amount", pVal.Row, prmAmount.ToString)

                                ElseIf PdcGrid.DataTable.GetValue(("Check"), pVal.Row) = "N" Then
                                    PdcGrid.DataTable.SetValue("Collection Amount", pVal.Row, "0")
                                End If

                                SumCollection(oForm)
                            End If


                            oForm.Freeze(False)
                            PdcGrid = Nothing
                        End If
                        If (pVal.ColUID = "Check") And pVal.Row = -1 Then
                            Dim PdcGrid As SAPbouiCOM.Grid = Nothing
                            Dim i As Integer
                            Dim prmAmount As Decimal
                            oForm.Freeze(True)
                            PdcGrid = oForm.Items.Item("Grid").Specific

                            If PdcGrid.Columns.Item("Check").Editable = True Then
                                If PdcGrid.Columns.Item("Check").TitleObject.Caption = "Select All" Then
                                    For i = 0 To PdcGrid.Rows.Count - 1
                                        PdcGrid.DataTable.SetValue("Check", i, "Y")
                                        prmAmount = PdcGrid.DataTable.GetValue(("Max Collection"), i)
                                        PdcGrid.DataTable.SetValue("Collection Amount", i, prmAmount.ToString)
                                    Next
                                    PdcGrid.Columns.Item("Check").TitleObject.Caption = "Reset All"

                                Else
                                    For i = 0 To PdcGrid.Rows.Count - 1
                                        PdcGrid.DataTable.SetValue("Check", i, "N")
                                        PdcGrid.DataTable.SetValue("Collection Amount", i, "0")
                                    Next
                                    PdcGrid.Columns.Item("Check").TitleObject.Caption = "Select All"
                                End If

                                SumCollection(oForm)
                            End If
                            
                            oForm.Freeze(False)
                            PdcGrid = Nothing
                        End If


                End Select
            End If
        End If


    End Sub

#End Region

#Region "Pdc Yang Di Tolak Class"

    Private Sub PdcTolakFirstLoad()
        Dim oForm As SAPbouiCOM.Form
        Dim PdcQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

        Dim PdcGrid As SAPbouiCOM.Grid

        Try
            oForm = SBO_Application.Forms.Item("PdcUpd_01")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "PdcUpd"
            fcp.UniqueID = "PdcUpd_01"


            fcp.XmlData = LoadFromXML("FormPdcTolak.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)
            oForm.ClientHeight = 507
            oForm.DataSources.UserDataSources.Add("PdcBank", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("PdcDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("PdcNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("PdcAmount", SAPbouiCOM.BoDataType.dt_QUANTITY)
            oForm.DataSources.UserDataSources.Add("TotalCol", SAPbouiCOM.BoDataType.dt_QUANTITY)
            oForm.DataSources.UserDataSources.Add("CustCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("CustName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Item("PdcDate").Value = DateTime.Today.ToString("yyyyMMdd")


            oEditText = oForm.Items.Item("PdcDate").Specific
            oEditText.DataBind.SetBound(True, "", "PdcDate")


            oEditText = oForm.Items.Item("PdcDate").Specific
            oEditText.DataBind.SetBound(True, "", "PdcDate")

            oEditText = oForm.Items.Item("PdcNo").Specific
            oEditText = oForm.Items.Item("PdcAmount").Specific
            oEditText = oForm.Items.Item("TotalCol").Specific

            oEditText = oForm.Items.Item("CustCode").Specific
            oEditText.DataBind.SetBound(True, "", "CustCode")

            oEditText = oForm.Items.Item("CustName").Specific


            oForm.DataSources.UserDataSources.Item("PdcAmount").Value = oMIS_Utils.fctFormatNumSBO(0, oCompany)
            oEditText = oForm.Items.Item("PdcAmount").Specific
            oEditText.DataBind.SetBound(True, "", "PdcAmount")

            oForm.DataSources.UserDataSources.Item("TotalCol").Value = oMIS_Utils.fctFormatNumSBO(0, oCompany)
            oEditText = oForm.Items.Item("TotalCol").Specific
            oEditText.DataBind.SetBound(True, "", "TotalCol")


            PdcGrid = oForm.Items.Item("Grid").Specific

            PdcGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto



            PdcQuery = "SELECT 1 as [Pdc Line],[@MIS_T3].DocNum as [T3 No.],[@MIS_T3L].U_OINVDocNum as [Invoice No.]," & _
                       "[@MIS_T3L].U_OINVDocDate as [Invoice Date]," & _
                       "[@MIS_T3L].U_OINVDocDueDate as [Due Date],[@MIS_T3L].U_OINVDocTotal as [Invoice Amount],0 as [Collection Amount] " & _
                        "FROM [@MIS_T3L] LEFT OUTER JOIN " & _
                        "[@MIS_PDCL] ON [@MIS_T3L].DocEntry = [@MIS_PDCL].U_T3DocEntry AND " & _
                        "[@MIS_T3L].LineId = [@MIS_PDCL].U_T3LineId LEFT OUTER JOIN " & _
                        "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry where 1=0 "




            ' Grid #: 1

            oForm.DataSources.DataTables.Add("PdcUpd")
            oForm.DataSources.DataTables.Item("PdcUpd").ExecuteQuery(PdcQuery)
            PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcUpd")


            PdcGrid.Columns.Item("Pdc Line").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Pdc Line").Editable = False
            PdcGrid.Columns.Item("Pdc Line").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Pdc Line").Width = 50

            PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
            PdcGrid.Columns.Item("T3 No.").Editable = False
            PdcGrid.Columns.Item("T3 No.").Width = 120

            oColumn = PdcGrid.Columns.Item("Invoice No.")
            oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
            PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
            oColumn.Editable = False
            PdcGrid.Columns.Item("Invoice No.").Width = 120


            PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Invoice Date").Editable = False
            PdcGrid.Columns.Item("Invoice Date").Width = 120

            PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Due Date").Editable = False
            PdcGrid.Columns.Item("Due Date").Width = 120

            PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Invoice Amount").Editable = False
            PdcGrid.Columns.Item("Invoice Amount").Width = 120

            PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = True
            PdcGrid.Columns.Item("Collection Amount").Width = 120
            PdcGrid.Columns.Item("Collection Amount").Editable = False

            oForm.Items.Item("CustCode").Click()

            oForm.Freeze(False)
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            PdcGrid = Nothing

            GC.Collect()


        End Try
        oForm.Freeze(False)
        oForm.Visible = True
    End Sub

    Private Sub PdcTolakEmpty(ByVal oForm As SAPbouiCOM.Form)

        oForm.Freeze(True)

        oForm.Items.Item("PdcBank").Specific.value = ""
        oForm.Items.Item("PdcNo").Specific.value = ""
        oForm.Items.Item("PdcAmount").Specific.value = 0
        oForm.Items.Item("TotalCol").Specific.value = 0
        oForm.Items.Item("CustCode").Specific.value = ""
        oForm.Items.Item("CustName").Specific.value = ""


        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim PdcQuery As String

        PdcGrid = oForm.Items.Item("Grid").Specific


        PdcGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto



        PdcQuery = "SELECT 1 as [No.],[@MIS_T3].DocNum as [T3 No.],[@MIS_T3L].U_OINVDocNum as [Invoice No.]," & _
                       "[@MIS_T3L].U_OINVDocDate as [Invoice Date]," & _
                       "[@MIS_T3L].U_OINVDocDueDate as [Due Date],[@MIS_T3L].U_OINVDocTotal as [Invoice Amount],0 as [Collection Amount] " & _
                        "FROM [@MIS_T3L] LEFT OUTER JOIN " & _
                        "[@MIS_PDCL] ON [@MIS_T3L].DocEntry = [@MIS_PDCL].U_T3DocEntry AND " & _
                        "[@MIS_T3L].LineId = [@MIS_PDCL].U_T3LineId LEFT OUTER JOIN " & _
                        "[@MIS_T3] ON [@MIS_T3L].DocEntry = [@MIS_T3].DocEntry where 1=0 "




        ' Grid #: 1

        oForm.DataSources.DataTables.Item("PdcUpd")
        oForm.DataSources.DataTables.Item("PdcUpd").ExecuteQuery(PdcQuery)
        PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcUpd")


        PdcGrid.Columns.Item("No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("No.").Width = 50
        PdcGrid.Columns.Item("No.").Editable = False

        PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("T3 No.").Editable = False
        PdcGrid.Columns.Item("T3 No.").Width = 120

        oColumn = PdcGrid.Columns.Item("Invoice No.")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
        oColumn.Editable = False
        PdcGrid.Columns.Item("Invoice No.").Width = 120


        PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Date").Editable = False
        PdcGrid.Columns.Item("Invoice Date").Width = 120

        PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Due Date").Editable = False
        PdcGrid.Columns.Item("Due Date").Width = 120

        PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Amount").Editable = False
        PdcGrid.Columns.Item("Invoice Amount").Width = 120

        PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Collection Amount").Width = 120
        PdcGrid.Columns.Item("Collection Amount").Editable = False

        oForm.Freeze(False)
        PdcGrid = Nothing
        oColumn = Nothing

    End Sub

    Private Sub PdcTolakShow(ByVal oForm As SAPbouiCOM.Form)

        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim PdcQuery As String

        oForm.Freeze(True)

        If oForm.Items.Item("CustCode").Specific.value = "" Then
            SBO_Application.MessageBox("Customer Code To Must fill", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcBank").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Bank To Must fill", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If

        If oForm.Items.Item("PdcNo").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc No To Must fill", 1, "OK")
            oForm.Items.Item("PdcNo").Click()
            GoTo Keluar
        End If


        oForm.Items.Item("TotalCol").Specific.value = 0


        PdcGrid = oForm.Items.Item("Grid").Specific




        PdcQuery = "SELECT [@MIS_PDCL].LineId as [Pdc Line], " & _
                   "[@MIS_T3].DocNum as [T3 No.], " & _
                   "[@MIS_PDCL].U_OINVDocEntry as [Invoice Doc Entry], " & _
                   "[@MIS_PDCL].U_OINVDocNum as [Invoice No.], " & _
                   "[@MIS_T3L].U_OINVDocDate as [Invoice Date]," & _
                   "[@MIS_T3L].U_OINVDocDueDate as [Due Date], " & _
                   "[@MIS_T3L].U_OINVDocTotal as [Invoice Amount]," & _
                   "[@MIS_PDCL].U_CollectAmount as [Collection Amount], " & _
                   "[@MIS_PDCL].DocEntry as [Pdc Doc Entry], " & _
                   "[@MIS_PDC].U_PDCDate as [Pdc Date], " & _
                   "[@MIS_PDC].U_PDCAmount as [Pdc Amount], " & _
                   "[@MIS_PDCL].U_T3DocEntry as [T3 Doc Entry], " & _
                   "[@MIS_PDCL].U_T3LineId as [T3 Line No] " & _
                   "FROM [@MIS_T3] RIGHT OUTER JOIN " & _
                   "[@MIS_T3L] ON [@MIS_T3].DocEntry =[@MIS_T3L].DocEntry RIGHT OUTER JOIN " & _
                   "[@MIS_PDCL] INNER JOIN " & _
                   "[@MIS_PDC] ON [@MIS_PDCL].DocEntry =[@MIS_PDC].DocEntry ON [@MIS_T3L].DocEntry =[@MIS_PDCL].U_T3DocEntry AND " & _
                   "[@MIS_T3L].LineId =[@MIS_PDCL].U_T3LineId where U_PDCStatus='O' And " & _
                   "[@MIS_PDC].U_CardCode='" & oForm.Items.Item("CustCode").Specific.value & "' " & _
                   "and [@MIS_PDC].U_PDCBankID ='" & oForm.Items.Item("PdcBank").Specific.value & "' " & _
                   "and  [@MIS_PDC].U_PDCNo='" & oForm.Items.Item("PdcNo").Specific.value & "' order by [@MIS_PDCL].LineId "




        ' Grid #: 1


        oForm.DataSources.DataTables.Item("PdcUpd")
        oForm.DataSources.DataTables.Item("PdcUpd").ExecuteQuery(PdcQuery)
        PdcGrid.DataTable = oForm.DataSources.DataTables.Item("PdcUpd")



        PdcGrid.Columns.Item("Pdc Line").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Pdc Line").Editable = False
        PdcGrid.Columns.Item("Pdc Line").Width = 60
        PdcGrid.Columns.Item("Pdc Line").Editable = False

        PdcGrid.Columns.Item("T3 No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("T3 No.").Editable = False
        PdcGrid.Columns.Item("T3 No.").Width = 100

        PdcGrid.Columns.Item("Invoice No.").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice No.").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice No.").Editable = False
        PdcGrid.Columns.Item("Invoice No.").Width = 100

        oColumn = PdcGrid.Columns.Item("Invoice Doc Entry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice
        PdcGrid.Columns.Item("Invoice Doc Entry").TitleObject.Sortable = True
        oColumn.Editable = False
        PdcGrid.Columns.Item("Invoice Doc Entry").Width = 100



        PdcGrid.Columns.Item("Invoice Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Date").Editable = False
        PdcGrid.Columns.Item("Invoice Date").Width = 100

        PdcGrid.Columns.Item("Due Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Due Date").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Due Date").Editable = False
        PdcGrid.Columns.Item("Due Date").Width = 100

        PdcGrid.Columns.Item("Invoice Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Invoice Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Invoice Amount").Editable = False
        PdcGrid.Columns.Item("Invoice Amount").Width = 100

        PdcGrid.Columns.Item("Collection Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Collection Amount").TitleObject.Sortable = True
        PdcGrid.Columns.Item("Collection Amount").Width = 100
        PdcGrid.Columns.Item("Collection Amount").Editable = False

        PdcGrid.Columns.Item("Pdc Doc Entry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Pdc Doc Entry").Visible = False
        PdcGrid.Columns.Item("Pdc Doc Entry").Width = 120

        PdcGrid.Columns.Item("Pdc Date").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Pdc Date").Visible = False
        PdcGrid.Columns.Item("Pdc Date").Width = 120

        PdcGrid.Columns.Item("Pdc Amount").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("Pdc Amount").Visible = False
        PdcGrid.Columns.Item("Pdc Amount").Width = 50

        PdcGrid.Columns.Item("T3 Doc Entry").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 Doc Entry").Visible = False
        PdcGrid.Columns.Item("T3 Doc Entry").Width = 50

        PdcGrid.Columns.Item("T3 Line No").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
        PdcGrid.Columns.Item("T3 Line No").Visible = False
        PdcGrid.Columns.Item("T3 Line No").Width = 50

        oForm.DataSources.UserDataSources.Item("PdcDate").Value = Format(PdcGrid.DataTable.GetValue(("Pdc Date"), PdcGrid.GetDataTableRowIndex(0).ToString), "yyyyMMdd")

        'oForm.Items.Item("PdcDate").Specific.value = oMIS_Utils.fctFormatDate(PdcGrid.DataTable.GetValue(("Pdc Date"), PdcGrid.GetDataTableRowIndex(0).ToString), oCompany, 5)
        oForm.Items.Item("PdcAmount").Specific.value = PdcGrid.DataTable.GetValue(("Pdc Amount"), PdcGrid.GetDataTableRowIndex(0).ToString)
        SumTolakCollection(oForm)

Keluar:
        oForm.Freeze(False)
        oForm = Nothing
        oColumn = Nothing
        PdcGrid = Nothing


    End Sub

    Private Sub PdcUpdate(ByVal oForm As SAPbouiCOM.Form)
        On Error GoTo Keluar
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralDataChild As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralDataLines As SAPbobsCOM.GeneralData
        Dim oGeneralDataLinesRows As SAPbobsCOM.GeneralDataCollection

        Dim oChild As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim strQuery As String
        Dim strseries As String
        Dim i As Integer
        Dim CheckInt As Integer = 0

        oCmpSrv = oCompany.GetCompanyService

        PdcGrid = oForm.Items.Item("Grid").Specific


        If oForm.Items.Item("PdcBank").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc Bank To Must fill", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If


        If oForm.Items.Item("PdcNo").Specific.value = "" Then
            SBO_Application.MessageBox("Pdc No To Must fill", 1, "OK")
            oForm.Items.Item("PdcNo").Click()
            GoTo Keluar
        End If


        If oForm.Items.Item("CustCode").Specific.value = "" Then
            SBO_Application.MessageBox("Customer Code To Must fill", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If

        oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strQuery = "SELECT TOP 1 U_CardName FROM [@MIS_T3] WHERE U_CardCode = '" & oForm.Items.Item("CustCode").Specific.value & "'"

        oRecordset.DoQuery(strQuery)
        If oRecordset.RecordCount < 1 Then
            SBO_Application.MessageBox("Customer Code Not Found ", 1, "OK")
            oForm.Items.Item("CustCode").Click()
            GoTo Keluar
        End If

        strQuery = "SELECT U_BankID FROM [@BANKGIRO] WHERE U_BankID = '" & oForm.Items.Item("PdcBank").Specific.value & "'"

        oRecordset.DoQuery(strQuery)
        If oRecordset.RecordCount < 1 Then
            SBO_Application.MessageBox("Bank Code Not Found ", 1, "OK")
            oForm.Items.Item("PdcBank").Click()
            GoTo Keluar
        End If

        strQuery = "SELECT U_PDCNo FROM [@MIS_PDC] WHERE U_PDCNo = '" & oForm.Items.Item("PdcNo").Specific.value & "'"

        oRecordset.DoQuery(strQuery)
        If oRecordset.RecordCount < 1 Then
            SBO_Application.MessageBox("PDC No Not Found ", 1, "OK")
            oForm.Items.Item("PdcNo").Click()
            GoTo Keluar
        End If


        If Not oCompany.InTransaction Then
            oCompany.StartTransaction()
        End If

        oGeneralService = oCmpSrv.GetGeneralService("PDC")

        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)




        For i = 0 To PdcGrid.Rows.Count - 1
            If i = 0 Then
                oGeneralParams.SetProperty("DocEntry", PdcGrid.DataTable.GetValue(("Pdc Doc Entry"), i))
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                oGeneralData.SetProperty("U_PDCStatus", "V")
            End If

            oGeneralDataLinesRows = oGeneralData.Child("MIS_PDCL")
            oGeneralDataLines = oGeneralDataLinesRows.Item(i)

            oGeneralDataLines.SetProperty("U_InvPaidStatus", "V")
            oGeneralService.Update(oGeneralData)


            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                GoTo Keluar
                'Else
                'strQuery = "Update [@MIS_T3L] set U_T3PDCStatus='V' " & _
                '           "where DocEntry='" & PdcGrid.DataTable.GetValue(("T3 Doc Entry"), PdcGrid.GetDataTableRowIndex(i).ToString) & "'" & _
                '           "and LineId='" & PdcGrid.DataTable.GetValue(("T3 Line No"), PdcGrid.GetDataTableRowIndex(i).ToString) & "'"
                'oRecordset.DoQuery(strQuery)
            End If

        Next


        If oCompany.InTransaction Then
            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If

        PdcTolakEmpty(oForm)
        oForm.Items.Item("CustCode").Click()
Keluar:
        If Err.Description <> "" Then
            SBO_Application.MessageBox("Exception: " & Err.Description, 1, "OK")
        End If

        If lRetCode <> 0 Or Err.Description <> "" Then
            If oCompany.InTransaction Then
                Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        End If

        oGeneralService = Nothing
        oGeneralData = Nothing
        oGeneralDataChild = Nothing
        oChild = Nothing
        oGeneralParams = Nothing
        oCmpSrv = Nothing
        PdcGrid = Nothing
        oColumn = Nothing
        oRecordset = Nothing

    End Sub

    Private Sub SumTolakCollection(ByVal oForm As SAPbouiCOM.Form)
        Dim PdcGrid As SAPbouiCOM.Grid = Nothing
        Dim DecCollection As Decimal = 0

        PdcGrid = oForm.Items.Item("Grid").Specific

        Dim i As Integer
        For i = 0 To PdcGrid.Rows.Count - 1
            DecCollection = DecCollection + CDec(PdcGrid.DataTable.GetValue(("Collection Amount"), i))
        Next

        oForm.Items.Item("TotalCol").Specific.value = DecCollection
        PdcGrid = Nothing
    End Sub

    Private Sub PdcTolakAplicationItem(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        On Error GoTo keluar

        BubbleEvent = False
        If pVal.BeforeAction = False Then
            If pVal.FormTypeEx = "PdcUpd" Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    oForm = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If pVal.ItemUID = "CustCode" Then
                            Dim oRecordset As SAPbobsCOM.Recordset
                            Dim strQuery As String
                            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            strQuery = "SELECT TOP 1 U_CardName FROM [@MIS_T3] WHERE U_CardCode = '" & oForm.Items.Item("CustCode").Specific.value & "'"

                            oRecordset.DoQuery(strQuery)
                            If oRecordset.RecordCount > 0 Then
                                oForm.Items.Item("CustName").Specific.value = oRecordset.Fields.Item("U_CardName").Value
                            Else
                                oForm.Items.Item("CustName").Specific.value = ""
                            End If

                            oRecordset = Nothing
                        End If

                        If pVal.ItemUID = "PdcBank" Then
                            Dim oRecordset As SAPbobsCOM.Recordset
                            Dim strQuery As String
                            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            strQuery = "SELECT TOP 1 U_BankName FROM [@BANKGIRO] WHERE U_BankID = '" & oForm.Items.Item("PdcBank").Specific.value & "'"

                            oRecordset.DoQuery(strQuery)
                            If oRecordset.RecordCount > 0 Then
                                oForm.Items.Item("BankNm").Specific.value = oRecordset.Fields.Item("U_BankName").Value
                            Else
                                oForm.Items.Item("BankNm").Specific.value = ""
                            End If

                            oRecordset = Nothing
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "BtnShow" Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            PdcTolakShow(oForm)
                        ElseIf pVal.ItemUID = "BtnUpdate" Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            PdcUpdate(oForm)
                        ElseIf pVal.ItemUID = "BtnCancel" Then
                            oForm = SBO_Application.Forms.Item(FormUID)
                            oForm.Close()
                        End If
                        

                End Select
            End If
        End If

keluar:

        BubbleEvent = True
    End Sub

#End Region



End Class
