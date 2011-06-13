
Option Explicit On
Option Strict Off

Imports MDS_AR.MIS_Utils

Public Class SOToMFG

    Public oCompany As SAPbobsCOM.Company
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter

    Dim oMIS_Utils As New MIS_Utils

    'Error handling variables
    Public sErrMsg As String
    Public lErrCode As Integer
    Public lRetCode As Integer

    Const ProductionIssue_MenuId As String = "4371"
    Const ProductionIssue_FormId As String = "65213"
    Const ProductionIssueUDF_FormId As String = "-65213"
    Dim objFormProductionIssue As SAPbouiCOM.Form
    Dim objFormProductionIssueUDF As SAPbouiCOM.Form
    Dim intRowProductionIssueDetail As Integer
    'karno 
    ' Production Issue
    Const Production_MenuId As String = "4369"
    Const Production_FormId As String = "65211"
    Const ProductionUDF_FormId As String = "-65211"
    Dim objFormProduction As SAPbouiCOM.Form
    Dim objFormProductionUDF As SAPbouiCOM.Form
    Dim intRowProductionDetail As Integer

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        '???
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            SBO_Application.MessageBox("MDS Production Addon now terminate...")
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

        'SetFilter()
        'subCreateMenu()
        'subCreateTable()

        Try
            LoadFromXML_Menu("MDSProdMenus.xml")

        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try

        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Private Function LoadFromXML(ByVal FileName As String) As String

        Dim oXmlDoc As Xml.XmlDocument
        Dim sPath As String

        oXmlDoc = New Xml.XmlDocument

        '// load the content of the XML File

        sPath = System.Windows.Forms.Application.StartupPath
        ''remove dir BIN
        'sPath = sPath.Remove(sPath.Length - 3, 3)

        'sPath = "E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK\"

        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString
        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString & "\"

        oXmlDoc.Load(sPath & "\" & FileName)

        '// load the form to the SBO application in one batch
        Return (oXmlDoc.InnerXml)

        'oXmlDoc = Nothing
        'sPath = Nothing
        'GC.Collect()
    End Function

    Private Sub LoadFromXML_Menu(ByVal FileName As String)
        'method Trial for adding menu using xml

        Dim oXmlDoc As Xml.XmlDocument

        oXmlDoc = New Xml.XmlDocument

        '// load the content of the XML File
        Dim sPath As String

        '        sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        ' Check build output path; remove the bin

        sPath = System.Windows.Forms.Application.StartupPath
        ' Check build output path; remove directory the "bin" to get app root path 
        '   e.g: E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK\bin
        'sPath = sPath.Remove(sPath.Length - 3, 3)


        '' Or
        '' Get Startup app path directory e.g: E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK
        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString

        '        sPath = "E:\Toin\SBO\Maruni\Production\Drobox_WIP\ProductionSDK\MaruniProductionSDK\"
        oXmlDoc.Load(sPath & "\" & FileName)

        ' e.g Adding Menu
        '// load the form to the SBO application in one batch
        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = SBO_Application.GetLastBatchResults()

        'MsgBox(GC.GetTotalMemory(True))
        'oXmlDoc = Nothing
        'sPath = Nothing

        ''not compatible to release oxmldoc using releaseComObject
        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc)
        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(sPath)

        'GC.Collect()

    End Sub

    'karno copy optim
    Private Sub CopyOptimize(ByVal oform As SAPbouiCOM.Form)
        Dim oItem As SAPbouiCOM.Item
        Dim oListOptimGrid As SAPbouiCOM.Grid = Nothing

        Try
            oform = SBO_Application.Forms.Item(ProductionIssue_FormId)
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "ListOptimize"
            fcp.UniqueID = "ListOptimize"
            fcp.XmlData = LoadFromXML("ListOptimize.srf")
            oform = SBO_Application.Forms.AddEx(fcp)

            oform.Freeze(True)

            oform.DataSources.DataTables.Add("ListOptim")
            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            'oform.DataSources.DataTables.Add("ListOptim")
            oform.DataSources.UserDataSources.Add("OptimNo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            'oItem = oform.Items.Add("BtnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Left = 145
            'oItem.Top = 318
            'oItem.Width = 150
            'oItem.Height = 19
            'oItem.Specific.caption = "Copy From Optimize"

            oItem = oform.Items.Add("myGrid2", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 80
            oItem.Width = oform.ClientWidth - 10
            oItem.Height = oform.ClientHeight - 200

            oListOptimGrid = oItem.Specific

            oListOptimGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oform.Freeze(False)

        End Try

    End Sub

    Private Sub CopyOptimizePro(ByVal oform As SAPbouiCOM.Form)
        Dim oItem As SAPbouiCOM.Item
        Dim oListOptimProGrid As SAPbouiCOM.Grid = Nothing

        Try
            oform = SBO_Application.Forms.Item(Production_FormId)
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "ListOptimizePro"
            fcp.UniqueID = "ListOptimizePro"
            fcp.XmlData = LoadFromXML("ListOptimizePro.srf")
            oform = SBO_Application.Forms.AddEx(fcp)

            oform.Freeze(True)

            oform.DataSources.DataTables.Add("ListOptimPro")
            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            'oform.DataSources.DataTables.Add("ListOptim")


            'oItem = oform.Items.Add("BtnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Left = 145
            'oItem.Top = 318
            'oItem.Width = 150
            'oItem.Height = 19
            'oItem.Specific.caption = "Copy From Optimize"

            oItem = oform.Items.Add("myGrid2", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 80
            oItem.Width = oform.ClientWidth - 10
            oItem.Height = oform.ClientHeight - 200

            oListOptimProGrid = oItem.Specific

            oListOptimProGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oform.Freeze(False)

        End Try

    End Sub

    'karno Production status
    Private Sub ProductionStatus()
        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText

        Dim oPDOStatusGrid As SAPbouiCOM.Grid

        Try
            oForm = SBO_Application.Forms.Item("PDOStatus")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "PDOStatus"
            fcp.UniqueID = "PDOStatus"
            fcp.XmlData = LoadFromXML("PDOStatus.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.DataTables.Add("PDOStatusLst")
            oForm.DataSources.UserDataSources.Add("TxtDtfrm", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("TxtDtTo", SAPbouiCOM.BoDataType.dt_DATE)


            'Default value for SO Date
            oForm.DataSources.UserDataSources.Item("TxtDtfrm").Value = oForm.Items.Item("TxtDtFrm").Specific.string

            oForm.DataSources.UserDataSources.Item("TxtDtTo").Value = oForm.Items.Item("TxtDtTo").Specific.string

            'Default setting
            ' add txtbox
            '        oEditText = oForm.Items.Add("SODate", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific


            oEditText = oForm.Items.Item("TxtDtFrm").Specific
            oEditText.DataBind.SetBound(True, "", "TxtDtFrm")
            oEditText = oForm.Items.Item("TxtDtTo").Specific
            oEditText.DataBind.SetBound(True, "", "TxtDtTo")

            '  add a GRID item to the form
            oItem = oForm.Items.Add("myGridPDO", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 80
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200



            'oSOToMFGGrid = oForm.Items.Item("myGrid").Specific
            oPDOStatusGrid = oItem.Specific

            oPDOStatusGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

        End Try
    End Sub
    'karno OutDel

    Private Sub OutDelEntry()
        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText

        Dim oDelOutGrid As SAPbouiCOM.Grid

        Try
            oForm = SBO_Application.Forms.Item("OutDel")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "OutDel"
            fcp.UniqueID = "OutDel"
            fcp.XmlData = LoadFromXML("OutDel.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.DataTables.Add("DelOutLst")
            oForm.DataSources.UserDataSources.Add("TxtDtfrm", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("TxtDtTo", SAPbouiCOM.BoDataType.dt_DATE)


            'Default value for SO Date
            oForm.DataSources.UserDataSources.Item("TxtDtfrm").Value = oForm.Items.Item("TxtDtFrm").Specific.string

            oForm.DataSources.UserDataSources.Item("TxtDtTo").Value = oForm.Items.Item("TxtDtTo").Specific.string


            'Default setting
            ' add txtbox
            '        oEditText = oForm.Items.Add("SODate", SAPbouiCOM.BoFormItemTypes.it_EDIT).Specific


            oEditText = oForm.Items.Item("TxtDtFrm").Specific
            oEditText.DataBind.SetBound(True, "", "TxtDtFrm")
            oEditText = oForm.Items.Item("TxtDtTo").Specific
            oEditText.DataBind.SetBound(True, "", "TxtDtTo")

            '  add a GRID item to the form
            oItem = oForm.Items.Add("myGrid1", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 5
            oItem.Top = 80
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            'oSOToMFGGrid = oForm.Items.Item("myGrid").Specific
            oDelOutGrid = oItem.Specific



            oDelOutGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

        End Try

    End Sub

    Private Function SOToMFGFormValid(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim isFormValid As Boolean

        isFormValid = True

        If Len(oForm.Items.Item("SODateFrom").Specific.string) = 0 Then
            isFormValid = False
        End If
        If Len(oForm.Items.Item("BPCardCode").Specific.string) = 0 Then
            isFormValid = False
        End If

        SOToMFGFormValid = isFormValid
    End Function

    Private Sub SOToMFGEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim SOToMFGQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oSOToMFGGrid As SAPbouiCOM.Grid

        Try
            oForm = SBO_Application.Forms.Item("mds_p1")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds"
            fcp.UniqueID = "mds_p1"

            fcp.XmlData = LoadFromXML("sotomfg.srf")
            'fcp.XmlData = LoadFromXML("form01.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("SODateFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("SODateTo", SAPbouiCOM.BoDataType.dt_DATE)


            'Default value for SO Date
            oForm.DataSources.UserDataSources.Item("SODateFrom").Value = DateTime.Today.ToString("yyyyMMdd")

            oForm.DataSources.UserDataSources.Item("SODateTo").Value = DateTime.Today.ToString("yyyyMMdd")

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("SoNumber", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            ''Dim bpCFL01 As MISToolbox
            ''bpCFL01 = New MISToolbox

            ''bpCFL01.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, "SOBPCFL1", "SOBPCFL2", _
            ''                "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C")

            ''bpCFL01 = Nothing
            ''GC.Collect()



            oEditText = oForm.Items.Item("BPCardCode").Specific
            oButton = oForm.Items.Item("BPButton").Specific

            oEditText.DataBind.SetBound(True, "", "BPDS")


            ''oEditText.ChooseFromListUID = "SOBPCFL1"
            ''oEditText.ChooseFromListAlias = "CardCode"
            ''oButton.ChooseFromListUID = "SOBPCFL2"

            oEditText = oForm.Items.Item("SoNumber").Specific

            oForm.Items.Item("SODateFrom").Width = 100
            oEditText = oForm.Items.Item("SODateFrom").Specific
            oEditText.DataBind.SetBound(True, "", "SODateFrom")

            oForm.Items.Item("SODateTo").Width = 100
            oEditText = oForm.Items.Item("SODateTo").Specific
            oEditText.DataBind.SetBound(True, "", "SODateTo")



            oItem = oForm.Items.Item("myGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200


            oSOToMFGGrid = oItem.Specific

            oSOToMFGGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)


            SOToMFGQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, " _
            & " 'Y' [Release PdO], " _
            & " CONVERT(VARCHAR(10), T0.DocDate, 102) [SO Date], T0.DocEntry [DocEntry], T0.DocNum [DocNum], VisOrder + 1 [SO Line]," _
            & " T3.SlpName [Sales Rep.], T0.CardCode [Cust. Code], T0.CardName [Customer Name], " _
            & " T1.ItemCode FG,T1.Dscription FGName, Quantity, " _
            & " T1.WhsCode, T2.InvntryUom UOM, T0.DocDueDate [Exp Delivery Date], " _
            & " T1.U_SO_Pcm PanjangInCm, T1.U_SO_Lcm LebarInCm, " _
            & " Case " _
            & " when T1.[U_SO_Bentuk] ='S' then 'Segi' " _
            & " when T1.[U_SO_Bentuk] ='J' then 'Jenjang' " _
            & " when T1.[U_SO_Bentuk] ='O' then 'Oval' " _
            & " when T1.[U_SO_Bentuk] ='B' then 'Bulat' " _
            & " when T1.[U_SO_Bentuk] ='G' then 'Bending' " _
            & " when T1.[U_SO_Bentuk] ='M' then 'Mal' " _
            & " when T1.[U_SO_Bentuk] ='X' then 'Others' " _
            & " Else 'Undefined' " _
            & " End SO_Bentuk " _
            & " FROM ORDR T0 " _
            & " LEFT JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " _
            & " LEFT JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " _
            & " LEFT JOIN OSLP T3 ON T1.SlpCode = T3.SlpCode " _
            & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
            & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
            & " AND T1.LineStatus = 'O' " _
            & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _
            & " AND (T1.U_MIS_ReleasePdOFlag = '' OR T1.U_MIS_ReleasePdOFlag IS NULL) " _
            & " ORDER BY T0.DocDate, T0.DocNum, VisOrder DESC "

            '& " AND T1.WhsCode = 'FG-002' " _
            '            & " AND T1.U_MIS_ReleasePdOFlag = '' AND T1.U_MIS_SupplyWith = 'M' "



            ' Grid #: 1
            oForm.DataSources.DataTables.Add("SOToMFGLst")
            oForm.DataSources.DataTables.Item("SOToMFGLst").ExecuteQuery(SOToMFGQuery)
            oSOToMFGGrid.DataTable = oForm.DataSources.DataTables.Item("SOToMFGLst")


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oSOToMFGGrid = Nothing

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try

        ''oForm.Top = 150
        ''oForm.Left = 330
        ''oForm.Width = 900


        SOToMFGQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, " _
            & " 'Y' [Release PdO], " _
            & " CONVERT(VARCHAR(10), T0.DocDate, 102) [SO Date], T0.DocEntry [DocEntry], T0.DocNum [DocNum], VisOrder + 1 [SO Line]," _
            & " T3.SlpName [Sales Rep.], T0.CardCode [Cust. Code], T0.CardName [Customer Name], " _
            & " T1.ItemCode FG,T1.Dscription FGName, Quantity, " _
            & " T1.WhsCode, T2.InvntryUom UOM, T0.DocDueDate [Exp Delivery Date], " _
            & " T1.U_SO_Pcm PanjangInCm, T1.U_SO_Lcm LebarInCm, " _
            & " Case " _
            & " when T1.[U_SO_Bentuk] ='S' then 'Segi' " _
            & " when T1.[U_SO_Bentuk] ='J' then 'Jenjang' " _
            & " when T1.[U_SO_Bentuk] ='O' then 'Oval' " _
            & " when T1.[U_SO_Bentuk] ='B' then 'Bulat' " _
            & " when T1.[U_SO_Bentuk] ='G' then 'Bending' " _
            & " when T1.[U_SO_Bentuk] ='M' then 'Mal' " _
            & " when T1.[U_SO_Bentuk] ='X' then 'Others' " _
            & " Else 'Undefined' " _
            & " End SO_Bentuk " _
            & " FROM ORDR T0 " _
            & " LEFT JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " _
            & " LEFT JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " _
            & " LEFT JOIN OSLP T3 ON T1.SlpCode = T3.SlpCode " _
            & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
            & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
            & " AND T1.LineStatus = 'O' " _
            & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _
            & " AND (T1.U_MIS_ReleasePdOFlag = '' OR T1.U_MIS_ReleasePdOFlag IS NULL) " _
            & " ORDER BY T0.DocDate, T0.DocNum, VisOrder DESC "
        '& " AND T1.WhsCode = 'FG-002' " _
        '    & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _

        '        & " JOIN MARUNI_SOTRIAL..mis_sofg002 T4 ON T0.DocNum = T4.[Document Number] AND T1.LineNum = T4.[Row Number] " _
        '            & " AND T1.U_MIS_ReleasePdOFlag = '' AND T1.U_MIS_SupplyWith = 'M' "


        oForm.DataSources.DataTables.Item(0).ExecuteQuery(SOToMFGQuery)

        oForm.Items.Item("BPCardCode").Click()



        RearrangeGrid(oForm)


        oForm.Visible = True

    End Sub

    Private Sub OptimizationEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim oCombobox As SAPbouiCOM.ComboBox = Nothing
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim lNextSeriesNumOptimization As Long

        Try
            oForm = SBO_Application.Forms.Item("mds_p3")
            SBO_Application.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds"
            fcp.UniqueID = "mds_p3"
            fcp.ObjectType = "MIS_OPTIM"
            fcp.XmlData = LoadFromXML("Optimization.srf")
            oForm = SBO_Application.Forms.AddEx(fcp)

            'oForm.DataBrowser.BrowseBy = "DocNum"

            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("OptimDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("QtyLembar", SAPbouiCOM.BoDataType.dt_QUANTITY)

            'Default value for Optimization Date
            oForm.DataSources.UserDataSources.Item("OptimDate").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("QtyLembar").Value = 2

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.DBDataSources.Add("@MIS_OPTIM")

            oForm.DataSources.DBDataSources.Add("@MIS_OPTIML")

            'oForm.Items.Item("OptimDate").Width = 100
            'oEditText = oForm.Items.Item("OptimDate").Specific
            'oEditText.DataBind.SetBound(True, "", "OptimDate")


            oForm.DataSources.UserDataSources.Add("#", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)


            'Bind Data to Form

            'Combo Series UDO
            oItem = oForm.Items.Add("SeriesOptm", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = 720
            oItem.Top = 10

            'Fill data for combo series
            oCombobox = oItem.Specific
            oCombobox.ValidValues.LoadSeries("MIS_OPTIM", SAPbouiCOM.BoSeriesMode.sf_Add)
            'New Method
            oCombobox.DataBind.SetBound(True, "@MIS_OPTIM", "SERIES")


            'oItem = oForm.Items.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Left = 300
            'oItem.Top = 10

            oEditText = oForm.Items.Item("DocNum").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "DocNum")

            lNextSeriesNumOptimization = oForm.BusinessObject.GetNextSerialNumber("SERIES")
            oEditText = oForm.Items.Item("DocNum").Specific
            oEditText.String = lNextSeriesNumOptimization

            'oEditText = oForm.Items.Item("DocNum").Specific
            'oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "DocNum")

            oEditText = oForm.Items.Item("OptimDate").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_OptDate")

            oEditText = oForm.Items.Item("OptimRef").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_OptNum")

            oEditText = oForm.Items.Item("ItemCode").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_ItemCode")

            Dim itemCFL As SBOConnection
            itemCFL = New SBOConnection


            itemCFL.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_Items, "ItemCFL1", "ItemCFL2", "ItemCode", _
                            SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "")

            oEditText.ChooseFromListUID = "ItemCFL1"
            oButton = oForm.Items.Item("ItemButton").Specific
            oButton.ChooseFromListUID = "ItemCFL2"


            itemCFL.AddCFL1(oForm, SAPbouiCOM.BoLinkedObject.lf_Items, "ItemCFL3", "ItemCFL4", "ItemCode", _
                SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "")


            oEditText = oForm.Items.Item("ItemKcSisa").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_ItemCdKacaSisa")
            oEditText.ChooseFromListUID = "ItemCFL3"
            oButton = oForm.Items.Item("ItmSisaBtn").Specific
            oButton.ChooseFromListUID = "ItemCFL4"

            itemCFL = Nothing

            oEditText = oForm.Items.Item("Dscription").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_ItemDesc")

            oEditText = oForm.Items.Item("PnjangKaca").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_Pcm")

            oEditText = oForm.Items.Item("LebarKaca").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_Lcm")

            oEditText = oForm.Items.Item("QtyLembar").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_QtyinLembar")
            oEditText = oForm.Items.Item("LuasKaca").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_LuasM2")
            oEditText = oForm.Items.Item("SisaKcUtuh").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_KcSisaUtuh")
            oEditText = oForm.Items.Item("KacaPakai").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_KacaUsed")
            oEditText = oForm.Items.Item("TotalWaste").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_TotalWaste")
            oEditText = oForm.Items.Item("ByUser").Specific
            oEditText.DataBind.SetBound(True, "@MIS_OPTIM", "U_MIS_User")


            'Set Matrix - add column from PdO & MIS_OPTIML
            'Dim oItem As SAPbouiCOM.Item
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim oColumns As SAPbouiCOM.Columns
            Dim oColumn As SAPbouiCOM.Column

            'oItem = oForm.Items.Item("OptimMtx").Specific
            oItem = oForm.Items.Item("OptimMtx")
            oItem.Width = 980
            oItem.Height = 350


            oMatrix = oItem.Specific

            oColumns = oMatrix.Columns

            'Add Column to Matrix
            'oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oColumn.TitleObject.Caption = "#"
            'oColumn.Width = 20
            'oColumn.Editable = False

            oColumn = oColumns.Item("#")
            oColumn.TitleObject.Caption = "#"
            oColumn.Width = 30
            oColumn.DataBind.SetBound(True, , "#")
            oColumn.Editable = False

            oColumn = oColumns.Item("LineId")
            oColumn.TitleObject.Caption = "LineId"
            oColumn.Width = 40
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "LineId")


            'Dim mistoolbox As MISToolbox
            'mistoolbox = New MISToolbox
            'mistoolbox.AddChooseFromListForMatrix(oForm, SAPbouiCOM.BoLinkedObject.lf_ProductionOrder, "PdOCFL1", "Status", _
            '                   SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL, "L")

            'oColumn = oColumns.Add("PdOButton", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oColumn.Width = 20
            'oColumn.ChooseFromListUID = "PdOCFL2"
            'oButton.ChooseFromListUID = "PdOCFL2"

            oColumn = oColumns.Item("PdO#")
            oColumn.TitleObject.Caption = "Pdo No."
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_PdONum")


            oColumn = oColumns.Item("SO#")
            'oColumn = oColumns.Add("SO#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "SO Num"
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_SONum")


            oColumn = oColumns.Item("SOLine")
            oColumn.TitleObject.Caption = "SOLine"
            oColumn.Width = 40
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_SOLineNum")


            oColumn = oColumns.Item("CardCode")
            oColumn.TitleObject.Caption = "Cust. Code"
            oColumn.Width = 50
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_CardCode")

            'oColumn.ChooseFromListUID = "PdOCFL1"
            'oColumn.ChooseFromListAlias = "DocNum"

            'mistoolbox = Nothing



            oColumn = oColumns.Item("CardName")
            oColumn.TitleObject.Caption = "Customer Name"
            oColumn.Width = 120
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_CardName")


            oColumn = oColumns.Item("QtyPotong")
            oColumn.TitleObject.Caption = "Jumlah Potong"
            oColumn.Width = 60
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_QtyPotong")


            oColumn = oColumns.Item("P")
            oColumn.TitleObject.Caption = "Panjang"
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_Pcm")

            oColumn = oColumns.Item("L")
            oColumn.TitleObject.Caption = "Lebar"
            oColumn.Width = 80
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_Lcm")


            oColumn = oColumns.Item("TotalABC")
            oColumn.TitleObject.Caption = "Total A x B x C"
            oColumn.Width = 100
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_TotalABC")

            oColumn = oColumns.Item("AlocWaste")
            oColumn.TitleObject.Caption = "Allocated Waste"
            oColumn.Width = 100
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_AllocatedWaste")

            oColumn = oColumns.Item("PlanPdIsue")
            oColumn.TitleObject.Caption = "Plan PdO Issue"
            oColumn.Width = 100
            oColumn.DataBind.SetBound(True, "@MIS_OPTIML", "U_MIS_QtPlanPdoIssue")



            oForm.DataBrowser.BrowseBy = "DocNum"
            'oForm.DataBrowser.BrowseBy = "U_MIS_ItemCode"



            oForm.EnableMenu("1292", True) 'Add Row
            oForm.EnableMenu("1293", True) 'Delete Row



            oForm.Freeze(False)


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing



            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))



        End Try

        'oForm.Top = 150
        'oForm.Left = 330
        'oForm.Width = 900





        'Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oOptimization As SAPbobsCOM.GeneralData
        'Dim oOptimizationParams As SAPbobsCOM.GeneralDataParams
        'Dim oOptimizationLinesParams As SAPbobsCOM.GeneralDataParams
        'Dim oOptimizationLines As SAPbobsCOM.GeneralData
        'Dim oOptimizationLinesRows As SAPbobsCOM.GeneralDataCollection
        'Dim oCompanyService As SAPbobsCOM.CompanyService



        'Dim vCompany As SAPbobsCOM.Company = Nothing
        'Dim sCookie As String
        'Dim sConnectionContext As String
        'Dim isconnect As Long
        'Dim errConnect As String = ""

        'vCompany = New SAPbobsCOM.Company
        'sCookie = vCompany.GetContextCookie
        'sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
        'vCompany.SetSboLoginContext(sConnectionContext)
        'isconnect = vCompany.Connect()
        'MsgBox("result: " & Str(isconnect))
        'Call vCompany.GetLastError(isconnect, errConnect)
        'MsgBox("lasterror: " & Str(isconnect) & "; msg: " & errConnect)


        'oCompanyService = oCompany.GetCompanyService

        ''Get General Service
        'oGeneralService = oCompanyService.GetGeneralService("MIS_OPTIM")

        ''Create Data for new row
        'oOptimization = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
        'oOptimization.SetProperty("U_MIS_ItemCode", "item01")
        'oOptimizationLinesRows = oOptimization.Child("MIS_OPTIML")
        'oOptimizationLines = oOptimizationLinesRows.Add
        'oOptimizationLines.SetProperty("U_MIS_CardCode", "bp01")
        'oOptimizationLines.SetProperty("U_MIS_SONum", 123)

        'oGeneralService.Add(oOptimization)

        ''Get UDT Optimization
        'oOptimizationParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        'oOptimizationParams.SetProperty("DocEntry", 2)
        'oOptimization = oGeneralService.GetByParams(oOptimizationParams)

        ''Update UDT Optimization
        'oOptimization.SetProperty("U_MIS_User", "dhh ruby")
        'oOptimization.SetProperty("U_MIS_TotalWaste", 108)

        ''Update UDT Document lines
        'oOptimizationLinesRows = oOptimization.Child("MIS_OPTIML")
        'oOptimizationLines = oOptimizationLinesRows.Item(0)
        ''Update UDT Optimization Lines
        ''oOptimizationLines.SetProperty("U_CardName", "railsss !!")
        ''oOptimizationLines.SetProperty("U_MIS_Pcm", 118)

        'oGeneralService.Update(oOptimization)


        'Get UDT Optimization Lines
        'oOptimizationLinesParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        'oOptimizationLinesParams.SetProperty("DocEntry", 2)
        'oOptimizationLinesParams.SetProperty("LineId", 1)

        'oOptimizationLines = oGeneralService.GetByParams(oOptimizationLinesParams)

        'Update UDT Document Lines
        'oOptimizationLines = oOptimizationLinesRows.Item(0)
        'Update UDT Optimization Lines
        'oOptimizationLines.SetProperty("U_CardName", "railsss !!")
        'oOptimizationLines.SetProperty("U_MIS_Pcm", 118)
        'oGeneralService.Update(oOptimizationLines)

        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oOptimizationLinesParams)
        'oOptimizationLinesParams = Nothing
        'GC.Collect()

        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oOptimizationParams)
        'oOptimizationParams = Nothing
        'GC.Collect()


        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oOptimization)
        'oOptimization = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oOptimizationLines)
        'oOptimizationLines = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oOptimizationLinesRows)
        'oOptimizationLinesRows = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService)
        'oGeneralService = Nothing
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompanyService)
        'oCompanyService = Nothing

        GC.Collect()


        'RearrangeGridOptimization(oForm)

        oForm.Items.Item("GTabc").Specific.value = 0
        oForm.Items.Item("GTaloc").Specific.value = 0
        oForm.Items.Item("GTplanPdO").Specific.value = 0

        oForm.Visible = True


    End Sub


    Private Sub LoadSO(ByVal oForm As SAPbouiCOM.Form)
        Dim SOToMFGQuery As String

        If oForm.Items.Item("BPCardCode").Specific.string = "" Then
            SBO_Application.SetStatusBarMessage("Customer must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Exit Sub
        End If

        If oForm.Items.Item("SoNumber").Specific.value = "" Then
            SBO_Application.SetStatusBarMessage("So Number Must Be Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Exit Sub
        End If

        If oForm.Items.Item("SODateTo").Specific.string = "" Then
            oForm.Items.Item("SODateTo").Specific.string = oForm.Items.Item("SODateFrom").Specific.string
        End If

        SOToMFGQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, " _
            & " 'Y' [Release PdO], " _
            & " CONVERT(VARCHAR(10), T0.DocDate, 102) [SO Date], T0.DocEntry [DocEntry], T0.DocNum [DocNum], VisOrder + 1 [SO Line]," _
            & " T3.SlpName [Sales Rep.], T0.CardCode [Cust. Code], T0.CardName [Customer Name], " _
            & " T1.ItemCode FG,T1.Dscription FGName, Quantity, " _
            & " T1.WhsCode, T2.InvntryUom UOM, T0.DocDueDate [Exp Delivery Date], " _
            & " T1.U_SO_Pcm PanjangInCm, T1.U_SO_Lcm LebarInCm, " _
            & " Case " _
            & " when T1.[U_SO_Bentuk] ='S' then 'Segi' " _
            & " when T1.[U_SO_Bentuk] ='J' then 'Jenjang' " _
            & " when T1.[U_SO_Bentuk] ='O' then 'Oval' " _
            & " when T1.[U_SO_Bentuk] ='B' then 'Bulat' " _
            & " when T1.[U_SO_Bentuk] ='G' then 'Bending' " _
            & " when T1.[U_SO_Bentuk] ='M' then 'Mal' " _
            & " when T1.[U_SO_Bentuk] ='X' then 'Others' " _
            & " Else 'Undefined' " _
            & " End SO_Bentuk " _
            & " FROM ORDR T0 " _
            & " LEFT JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " _
            & " LEFT JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " _
            & " LEFT JOIN OSLP T3 ON T1.SlpCode = T3.SlpCode " _
        & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
            & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
            & " AND T1.LineStatus = 'O' " _
            & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _
            & " AND (T1.U_MIS_ReleasePdOFlag = '' OR T1.U_MIS_ReleasePdOFlag IS NULL) " _
            & " AND T0.Docnum = " & oForm.Items.Item("SoNumber").Specific.value & " " _
            & " ORDER BY T0.DocDate, T0.DocNum, VisOrder DESC "
        '& " AND T1.WhsCode = 'FG-002' " _


        '& " JOIN MARUNI_SOTRIAL..mis_sofg002 T4 ON T0.DocNum = T4.[Document Number] AND T1.LineNum = T4.[Row Number] " _
        '    & " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _

        '            & " AND T1.U_MIS_ReleasePdOFlag = '' AND T1.U_MIS_SupplyWith = 'M' "


        '            & " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _

        '        & " , T1.U_SO_P1, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xJob1, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xDC1," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xIC1," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J1 WHERE Code = T1.U_SO_P1) xFOH1," _
        '& " T1.U_SO_P2, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xJob2, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xDC2," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xIC2," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J2 WHERE Code = T1.U_SO_P2) xFOH2," _
        '& " T1.U_SO_P3, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xJob3, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xDC3," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xIC3," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J3 WHERE Code = T1.U_SO_P3) xFOH3," _
        '& " T1.U_SO_P4, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xJob4, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xDC4," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xIC4," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J4 WHERE Code = T1.U_SO_P4) xFOH4," _
        '& " T1.U_SO_P5, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xJob5, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xDC5," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xIC5," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J5 WHERE Code = T1.U_SO_P5) xFOH5," _
        '& " T1.U_SO_P6, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xJob6, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xDC6," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xIC6," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J6 WHERE Code = T1.U_SO_P6) xFOH6," _
        '& " T1.U_SO_P7, (SELECT U_MIS_JobItemCode FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xJob7, " _
        '& " (SELECT U_MIS_DC FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xDC7," _
        '& " (SELECT U_MIS_IC FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xIC7," _
        '& " (SELECT U_MIS_FOH FROM [@IND_PL] J7 WHERE Code = T1.U_SO_P7) xFOH7 " _

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(SOToMFGQuery)

        RearrangeGrid(oForm)

    End Sub

    '    Private Sub GeneratePdOFromSO(ByVal oForm As SAPbouiCOM.Form)
    '        'On Error GoTo errHandler

    '        Dim oSOToMFGGrid As SAPbouiCOM.Grid

    '        Dim idx As Long



    '        Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
    '        Dim oSalesOrderLines As SAPbobsCOM.Document_Lines = Nothing

    '        Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
    '        Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines = Nothing

    '        Dim vCompany As SAPbobsCOM.Company = Nothing
    '        Dim sCookie As String
    '        Dim sConnectionContext As String

    '        Dim isconnect As Long
    '        Dim errConnect As String = ""

    '        Dim oPdODocSeriesRec As SAPbobsCOM.Recordset

    '        Dim strQry As String = ""
    '        Dim oPdODocSeries As String = ""

    '        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific



    '        'GRID - Order by column checkbox
    '        oSOToMFGGrid.Columns.Item("Release PdO").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)



    '        'Dim oSalesOrder As SAPbobsCOM.Documents
    '        'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines

    '        'Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
    '        'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines

    '        'Dim vCompany As SAPbobsCOM.Company = Nothing
    '        'Dim sCookie As String
    '        'Dim sConnectionContext As String

    '        'Dim isconnect As Long
    '        'Dim errConnect As String = ""  



    '        'Get PdO Doc. Series 
    '        'oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oPdODocSeriesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        ' FG-002: PENJUALAN JASA (SERIENAME:2011JS), FG-001: PENJUALAN ORDER (SERIENAME:2011) 

    '        If oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) = "FG-002" Then
    '            'oProd1.Series = 45
    '            strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) = 'JS' AND Indicator = YEAR(GETDATE()) "
    '        Else
    '            'oProd1.Series = 27
    '            strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) <> 'JS' AND Indicator = YEAR(GETDATE()) "
    '        End If


    '        oPdODocSeriesRec.DoQuery(strQry)
    '        '??? 
    '        If oPdODocSeriesRec.RecordCount <> 0 Then
    '            oPdODocSeries = oPdODocSeriesRec.Fields.Item("Series").Value
    '        Else
    '            MsgBox("Production Order Document Series Tidak ada, Mohon Setup PdO Document Series!")
    '            Exit Sub
    '        End If


    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPdODocSeriesRec)
    '        oPdODocSeriesRec = Nothing

    '        If oPdODocSeries <> "" Then

    '            'If oSOToMFGGrid.Rows.Count > 5 Then
    '            '    SBO_Application.MessageBox("Minimal 5 To Generate So", 1, "OK")
    '            'Else
    '            'Loop only selected/checked in grid rows and exit.
    '            For idx = oSOToMFGGrid.Rows.Count - 1 To 0 Step -1
    '                SBO_Application.SetStatusBarMessage("Generating PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

    '                If oSOToMFGGrid.DataTable.GetValue(1, oSOToMFGGrid.GetDataTableRowIndex(idx)) = "Y" Then


    '                    'MsgBox("line: " & oSOToMFGGrid.DataTable.GetValue(0, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
    '                    '       & "; Release PdO?#: " & oSOToMFGGrid.DataTable.GetValue(1, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
    '                    '       & "; so#: " & oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
    '                    ')


    '                    'SBO_Application.SetStatusBarMessage("Generating PdO....", SAPbouiCOM.BoMessageTime.bmt_Short, False)

    '                    'MsgBox(GC.GetTotalMemory(True))

    '                    'Try


    '                    'Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
    '                    'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines = Nothing

    '                    'Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
    '                    'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines = Nothing

    '                    'Dim vCompany As SAPbobsCOM.Company = Nothing
    '                    'Dim sCookie As String
    '                    'Dim sConnectionContext As String

    '                    'Dim isconnect As Long
    '                    'Dim errConnect As String = ""



    '                    Try
    '                        SetApplication()

    '                        vCompany = New SAPbobsCOM.Company
    '                        'Dim sCookie As String = vCompany.GetContextCookie
    '                        'Dim sConnectionContext As String
    '                        sCookie = vCompany.GetContextCookie
    '                        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
    '                        vCompany.SetSboLoginContext(sConnectionContext)
    '                        isconnect = vCompany.Connect()

    '                        'If vCompany.Connect() <> 0 Then
    '                        If isconnect <> 0 Then
    '                            End
    '                        End If
    '                    Catch ex As Exception
    '                        End
    '                    End Try


    '                    'vCompany = New SAPbobsCOM.Company
    '                    'sCookie = vCompany.GetContextCookie
    '                    'sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
    '                    'vCompany.SetSboLoginContext(sConnectionContext)

    '                    ''isconnect = vCompany.Connect()
    '                    ''MsgBox("result: " & Str(isconnect))
    '                    'Call vCompany.GetLastError(isconnect, errConnect)
    '                    'MsgBox("lasterror: " & Str(isconnect) & "; msg: " & errConnect)

    '                    'vCompany = oCompany

    '                    vCompany.StartTransaction()
    '                    'oCompany.StartTransaction()


    '                    'oSO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
    '                    'oSO.GetByKey(252)
    '                    'oSO.UserFields.Fields.Item("U_NBS_Range").Value = "123Tes321"
    '                    'oSO.Update()

    '                    ' by Toin 2011-02-09 Check Duplicate PdO before Generate PdO
    '                    '???

    '                    Dim oRS As SAPbobsCOM.Recordset

    '                    'vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                    oRS = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    strQry = "SELECT DocNum FROM OWOR WHERE OriginNum =  " & oSOToMFGGrid.DataTable.GetValue(4, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
    '                        & " AND ItemCode = '" & oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) & "' "
    '                    oRS.DoQuery(strQry)
    '                    'oRS.DoQuery("UPDATE RDR1 SET U_bacthNum = 'b321' where docentry = 249 and linenum = 0")


    '                    'If oRS.RecordCount <> 0 Then
    '                    '    MsgBox("ada record!")
    '                    'End If

    '                    If oRS.RecordCount = 0 Then

    '                        'oProd1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
    '                        oProd1 = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

    '                        'oprod1.ItemNo = "S10000"
    '                        'oprod1.DueDate = Today
    '                        oProd1.PlannedQuantity = 2

    '                        ''Fill oPdO properties...oProductionOrder
    '                        ''oProdOrder.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        ''oProdOrder.ItemNo = "LM4029"
    '                        'oprod1.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        'oprod1.ItemNo = "S10000"

    '                        ' Series PdO JS = 202 (NNM1) objectCode = 202 (OWOR PdO) series id = 45

    '                        ' IMPORTANT !!!
    '                        ' PdO SERIES YEAR 2011, 2011JS PdO JASA, SERIES# = 45
    '                        ' PdO SERIES YEAR 2011, 2011   PdO KACA SERIES# = 27 

    '                        If oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) = "FG-001" Then
    '                            'oProd1.Series = 27
    '                            oProd1.Series = oPdODocSeries
    '                        Else
    '                            'oProd1.Series = 45
    '                            oProd1.Series = oPdODocSeries
    '                        End If

    '                        'oProd1.ItemNo = "KTF12CLXX589"
    '                        oProd1.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

    '                        oProd1.PlannedQuantity = oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

    '                        ''oProdOrder.DueDate = oSOToMFGGrid.DataTable.GetValue(13, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

    '                        'PdO Posting Date = SO Posting Date
    '                        oProd1.PostingDate = oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        oProd1.PostingDate = Format(Now, "yyyy-MM-dd")

    '                        Dim dueDt As DateTime
    '                        Dim sodt As DateTime
    '                        Dim sodelivdt As DateTime
    '                        Dim dtdiff As Integer

    '                        sodt = oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        sodelivdt = oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
    '                        dtdiff = DateDiff(DateInterval.Day, CDate(oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)), CDate(oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)))
    '                        'sodelivdt = ""
    '                        dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
    '                        dueDt = DateAdd(DateInterval.Day, IIf(dtdiff < 0, 0, dtdiff), Now)

    '                        'PdO Due Date = SO Deliv. Date
    '                        'oProd1.DueDate = Today + n days (so date - so deliv date)
    '                        ''oProd1.DueDate = oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        oProd1.DueDate = dueDt

    '                        'oprod1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
    '                        oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
    '                        oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
    '                        'oprod1.Warehouse = "01"
    '                        'oProd1.Warehouse = "FG-001"

    '                        oProd1.Warehouse = oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

    '                        oProd1.CustomerCode = oSOToMFGGrid.DataTable.GetValue(7, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
    '                        ' so docnum
    '                        oProd1.ProductionOrderOriginEntry = oSOToMFGGrid.DataTable.GetValue(3, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

    '                        oProd1.UserFields.Fields.Item("U_PoD_Pcm").Value = oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString
    '                        oProd1.UserFields.Fields.Item("U_PdO_Lcm").Value = oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString
    '                        oProd1.UserFields.Fields.Item("U_PdO_Bentuk").Value = oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

    '                        'Dim QTY_LUASM2 As Double
    '                        'QTY_LUASM2 = _
    '                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
    '                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
    '                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString))


    '                        oProd1.UserFields.Fields.Item("U_SO_Luas_M2").Value = _
    '                        CStr( _
    '                            CDbl(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
    '                            CDbl(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
    '                            CDbl(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)) / 10000 _
    '                            )

    '                        'oprod1.UserFields.Fields.Item("U_PoD_Pcm").Value = "p100cm"
    '                        'oprod1.UserFields.Fields.Item("U_PdO_Lcm").Value = "L90cm"
    '                        'oprod1.UserFields.Fields.Item("U_PdO_Bentuk").Value = "segi"
    '                        'oprod1.UserFields.Fields.Item("U_NBS_OnHoldReason").Value = "test123"

    '                        oProdLine1 = oProd1.Lines

    '                        ' Generate one line - Dummy item
    '                        oProdLine1.ItemNo = "XDUMMY"
    '                        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
    '                        oProdLine1.Warehouse = "SRV-DL"


    '                        ' In Case of JOB 01 Exists
    '                        'If oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
    '                        '    oProdLine1.Add()
    '                        '    ' Job 01 - sample: XTP1 
    '                        '    'oprodLine1.ItemNo = "F12CLXXMM51003048"
    '                        '    oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        '    'oProdLine1.ItemNo = "XSE1"

    '                        '    'oProdLine1.Warehouse = "RM-PRD"
    '                        '    'oProdLine1.Warehouse = "SRV-DL"

    '                        '    oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush

    '                        '    'SO JOb 01 - DC dummy code (Direct Cost)
    '                        '    If oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
    '                        '        oProdLine1.Add()
    '                        '        oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        '        oProdLine1.Warehouse = "SRV-DL"
    '                        '        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
    '                        '    End If
    '                        '    'SO JOb 01 - IC dummy code (Indirect Cost)
    '                        '    If oSOToMFGGrid.DataTable.GetValue(18, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
    '                        '        oProdLine1.Add()
    '                        '        oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(18, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        '        oProdLine1.Warehouse = "SRV-DL"
    '                        '        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
    '                        '    End If
    '                        '    'SO JOb 01 - FOH dummy code (FOH)
    '                        '    If oSOToMFGGrid.DataTable.GetValue(19, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
    '                        '        oProdLine1.Add()
    '                        '        oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(19, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
    '                        '        oProdLine1.Warehouse = "SRV-DL"
    '                        '        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
    '                        '    End If

    '                        'End If





    '                        'MsgBox(GC.GetTotalMemory(True))

    '                        lRetCode = oProd1.Add()


    '                        If lRetCode <> 0 Then
    '                            'oCompany.GetLastError(lErrCode, sErrMsg)
    '                            vCompany.GetLastError(lErrCode, sErrMsg)
    '                            SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
    '                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                        Else

    '                            Dim PdOno As String = ""
    '                            Dim tmpKey As Double
    '                            Dim vSOLine As Long

    '                            vCompany.GetNewObjectCode(tmpKey)
    '                            vCompany.GetNewObjectCode(PdOno)
    '                            tmpKey = Convert.ToInt32(PdOno)

    '                            ' !!!! Make sure before create another object type-> clear previous/current object type.
    '                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
    '                            oProdLine1 = Nothing

    '                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
    '                            oProd1 = Nothing

    '                            'GC.Collect()

    '                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)


    '                            oSalesOrder = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
    '                            'oSalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
    '                            oSalesOrder.GetByKey(oSOToMFGGrid.DataTable.GetValue(3, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString))
    '                            'oSalesOrder.UserFields.Fields.Item("U_MIS_ReasonCode").Value = "T1"

    '                            vSOLine = oSOToMFGGrid.DataTable.GetValue(5, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) - 1
    '                            oSalesOrderLines = oSalesOrder.Lines
    '                            'oSalesOrderLines.SetCurrentLine(oSOToMFGGrid.DataTable.GetValue(5, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) - 1)
    '                            oSalesOrderLines.SetCurrentLine(vSOLine)
    '                            'oSalesOrderLines.SetCurrentLine(10)

    '                            oSalesOrderLines.UserFields.Fields.Item("U_MIS_SupplyWith").Value = "M"
    '                            oSalesOrderLines.UserFields.Fields.Item("U_MIS_ReleasePdOFlag").Value = "Y"
    '                            oSalesOrderLines.UserFields.Fields.Item("U_MIS_PdONum").Value = PdOno  '"1010025702"

    '                            oSalesOrder.Update()

    '                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrderLines)
    '                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrder)
    '                            oSalesOrder = Nothing
    '                            'oSalesOrderLines = Nothing

    '                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                            vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                            'oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

    '                        End If

    '                        SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

    '                    End If

    '                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
    '                    oRS = Nothing

    '                    ' by Toin 2011-02-09 Check Duplicate PdO before Generate PdO

    '                    vCompany.Disconnect()
    '                    System.Runtime.InteropServices.Marshal.ReleaseComObject(vCompany)
    '                    vCompany = Nothing

    '                    GC.Collect()
    '                    'MsgBox(GC.GetTotalMemory(True))


    '                    'MsgBox("generating... PdO; DONE!!!")

    '                Else
    '                    Exit For
    '                End If
    '            Next

    '        End If  ' Checking PdO Series

    '        'End If
    '        'Catch ex As Exception
    '        '    MsgBox(oCompany.GetLastErrorDescription)
    '        '    MsgBox(ex.Message)
    '        '    Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '        'End Try

    '        'Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '        'Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

    '        'MsgBox("Begin trx: generating... PdO")
    '        SBO_Application.MessageBox("Generating PdO.... Finished !!! ", 1, "Ok")

    '        'Begin Trxs

    '        'Call oCompany.StartTransaction()

    '        'Dim oSalesOrder As SAPbobsCOM.Documents
    '        'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines

    '        'oSalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
    '        'oSalesOrder.GetByKey(252)
    '        'oSalesOrder.UserFields.Fields.Item("U_NBS_Range").Value = "123Tes321"

    '        'oSalesOrder.Lines.SetCurrentLine(0)
    '        'oSalesOrder.Lines.UserFields.Fields.Item("U_BacthNum").Value = "223344"

    '        ''### How to update UDF using recordset
    '        'Dim oRS As SAPbobsCOM.Recordset
    '        ''Dim vCompany As SAPbobsCOM.Company
    '        ''vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        'oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        ''oRS.DoQuery("UPDATE RDR1 SET U_BatchNum = 'b 321' WHERE DocEntry = 252 and LineNum = 0")
    '        'oRS.DoQuery("UPDATE RDR1 SET U_bacthNum = 'b321' where docentry = 249 and linenum = 0")

    '        'If oRS.RecordCount <> 0 Then
    '        '    MsgBox("ada record!")
    '        'End If
    '        '### How to update UDF using recordset

    '        'oSalesOrderLines = oSalesOrder.Lines
    '        'oSalesOrderLines.UserFields.Fields.Item("U_BacthNum").Value = "b 321"

    '        'lRetCode = oSalesOrder.Update()
    '        'If lRetCode <> 0 Then
    '        '    oCompany.GetLastError(lErrCode, sErrMsg)
    '        '    SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)

    '        '    Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '        'End If

    '        'For idx = 0 To oSOToMFGGrid.Rows.SelectedRows.Count - 1
    '        '    MsgBox("selected row#:" & idx.ToString & _
    '        '           "; selectedrow->row#: " & oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
    '        '           & "DocEntry: " & oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

    '        'Next

    '        'Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

    '        'MsgBox("generating... PdO; DONE!!!")


    '        'Disconnect Company Object & Release Resource
    '        'Call oCompany.Disconnect()
    '        'oCompany = Nothing

    '        Exit Sub


    'errHandler:
    '        MsgBox("Exception: " & Err.Description)
    '        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '    End Sub

    Private Sub GeneratePdOFromSO(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler

        Dim oSOToMFGGrid As SAPbouiCOM.Grid

        Dim idx As Long



        Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
        Dim oSalesOrderLines As SAPbobsCOM.Document_Lines = Nothing

        Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
        Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines = Nothing

        Dim vCompany As SAPbobsCOM.Company = Nothing
        Dim sCookie As String
        Dim sConnectionContext As String

        Dim isconnect As Long
        Dim errConnect As String = ""

        Dim oPdODocSeriesRec As SAPbobsCOM.Recordset

        Dim strQry As String = ""
        Dim oPdODocSeries As String = ""

        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific



        'GRID - Order by column checkbox
        oSOToMFGGrid.Columns.Item("Release PdO").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)



        'Dim oSalesOrder As SAPbobsCOM.Documents
        'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines

        'Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
        'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines

        'Dim vCompany As SAPbobsCOM.Company = Nothing
        'Dim sCookie As String
        'Dim sConnectionContext As String

        'Dim isconnect As Long
        'Dim errConnect As String = ""  



        'Get PdO Doc. Series 
        'oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPdODocSeriesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ' FG-002: PENJUALAN JASA (SERIENAME:2011JS), FG-001: PENJUALAN ORDER (SERIENAME:2011) 

        If oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) = "FG-002" Then
            'oProd1.Series = 45
            strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) = 'JS' AND Indicator = YEAR(GETDATE()) "
        Else
            'oProd1.Series = 27
            strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) <> 'JS' AND Indicator = YEAR(GETDATE()) "
        End If


        oPdODocSeriesRec.DoQuery(strQry)
        '??? 
        If oPdODocSeriesRec.RecordCount <> 0 Then
            oPdODocSeries = oPdODocSeriesRec.Fields.Item("Series").Value
        Else
            MsgBox("Production Order Document Series Tidak ada, Mohon Setup PdO Document Series!")
            Exit Sub
        End If


        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPdODocSeriesRec)
        oPdODocSeriesRec = Nothing
        GC.Collect()

        If oPdODocSeries <> "" Then

            'If oSOToMFGGrid.Rows.Count > 5 Then
            '    SBO_Application.MessageBox("Minimal 5 To Generate So", 1, "OK")
            'Else
            'Loop only selected/checked in grid rows and exit.
            For idx = oSOToMFGGrid.Rows.Count - 1 To 0 Step -1
                SBO_Application.SetStatusBarMessage("Generating PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                If oSOToMFGGrid.DataTable.GetValue(1, oSOToMFGGrid.GetDataTableRowIndex(idx)) = "Y" Then


                    'MsgBox("line: " & oSOToMFGGrid.DataTable.GetValue(0, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
                    '       & "; Release PdO?#: " & oSOToMFGGrid.DataTable.GetValue(1, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
                    '       & "; so#: " & oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
                    ')


                    'SBO_Application.SetStatusBarMessage("Generating PdO....", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    'MsgBox(GC.GetTotalMemory(True))

                    'Try


                    'Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
                    'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines = Nothing

                    'Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
                    'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines = Nothing

                    'Dim vCompany As SAPbobsCOM.Company = Nothing
                    'Dim sCookie As String
                    'Dim sConnectionContext As String

                    'Dim isconnect As Long
                    'Dim errConnect As String = ""



                    Try
                        SetApplication()

                        vCompany = New SAPbobsCOM.Company
                        'Dim sCookie As String = vCompany.GetContextCookie
                        'Dim sConnectionContext As String
                        sCookie = vCompany.GetContextCookie
                        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
                        vCompany.SetSboLoginContext(sConnectionContext)
                        isconnect = vCompany.Connect()

                        'If vCompany.Connect() <> 0 Then
                        If isconnect <> 0 Then
                            End
                        End If
                    Catch ex As Exception
                        End
                    End Try


                    'vCompany = New SAPbobsCOM.Company
                    'sCookie = vCompany.GetContextCookie
                    'sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
                    'vCompany.SetSboLoginContext(sConnectionContext)

                    ''isconnect = vCompany.Connect()
                    ''MsgBox("result: " & Str(isconnect))
                    'Call vCompany.GetLastError(isconnect, errConnect)
                    'MsgBox("lasterror: " & Str(isconnect) & "; msg: " & errConnect)

                    'vCompany = oCompany

                    vCompany.StartTransaction()
                    'oCompany.StartTransaction()


                    'oSO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    'oSO.GetByKey(252)
                    'oSO.UserFields.Fields.Item("U_NBS_Range").Value = "123Tes321"
                    'oSO.Update()

                    ' by Toin 2011-02-09 Check Duplicate PdO before Generate PdO
                    '???

                    Dim oRS As SAPbobsCOM.Recordset

                    'vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oRS = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strQry = "SELECT DocNum FROM OWOR WHERE Status <> 'C' AND OriginNum =  " & oSOToMFGGrid.DataTable.GetValue(4, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) _
                        & " AND ItemCode = '" & oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) & "' "
                    oRS.DoQuery(strQry)
                    'oRS.DoQuery("UPDATE RDR1 SET U_bacthNum = 'b321' where docentry = 249 and linenum = 0")


                    'If oRS.RecordCount <> 0 Then
                    '    MsgBox("ada record!")
                    'End If

                    If oRS.RecordCount = 0 Then

                        'oProd1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        oProd1 = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                        'oprod1.ItemNo = "S10000"
                        'oprod1.DueDate = Today
                        oProd1.PlannedQuantity = 2

                        ''Fill oPdO properties...oProductionOrder
                        ''oProdOrder.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        ''oProdOrder.ItemNo = "LM4029"
                        'oprod1.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        'oprod1.ItemNo = "S10000"

                        ' Series PdO JS = 202 (NNM1) objectCode = 202 (OWOR PdO) series id = 45

                        ' IMPORTANT !!!
                        ' PdO SERIES YEAR 2011, 2011JS PdO JASA, SERIES# = 45
                        ' PdO SERIES YEAR 2011, 2011   PdO KACA SERIES# = 27 

                        If oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) = "FG-001" Then
                            'oProd1.Series = 27
                            oProd1.Series = oPdODocSeries
                        Else
                            'oProd1.Series = 45
                            oProd1.Series = oPdODocSeries
                        End If

                        'oProd1.ItemNo = "KTF12CLXX589"
                        oProd1.ItemNo = oSOToMFGGrid.DataTable.GetValue(9, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        oProd1.PlannedQuantity = oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        ''oProdOrder.DueDate = oSOToMFGGrid.DataTable.GetValue(13, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        'PdO Posting Date = SO Posting Date
                        oProd1.PostingDate = oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        oProd1.PostingDate = Format(Now, "yyyy-MM-dd")

                        Dim dueDt As DateTime
                        Dim sodt As DateTime
                        Dim sodelivdt As DateTime
                        Dim dtdiff As Integer

                        sodt = oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        sodelivdt = oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
                        dtdiff = DateDiff(DateInterval.Day, CDate(oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)), CDate(oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)))
                        'sodelivdt = ""
                        dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
                        dueDt = DateAdd(DateInterval.Day, IIf(dtdiff < 0, 0, dtdiff), Now)

                        'PdO Due Date = SO Deliv. Date
                        'oProd1.DueDate = Today + n days (so date - so deliv date)
                        ''oProd1.DueDate = oSOToMFGGrid.DataTable.GetValue(14, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        oProd1.DueDate = dueDt

                        'oprod1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                        oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                        oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                        'oprod1.Warehouse = "01"
                        'oProd1.Warehouse = "FG-001"

                        oProd1.Warehouse = oSOToMFGGrid.DataTable.GetValue(12, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        oProd1.CustomerCode = oSOToMFGGrid.DataTable.GetValue(7, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
                        ' so docnum
                        oProd1.ProductionOrderOriginEntry = oSOToMFGGrid.DataTable.GetValue(3, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        oProd1.UserFields.Fields.Item("U_PoD_Pcm").Value = oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString
                        oProd1.UserFields.Fields.Item("U_PdO_Lcm").Value = oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString
                        oProd1.UserFields.Fields.Item("U_PdO_Bentuk").Value = oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)

                        'Dim QTY_LUASM2 As Double
                        'QTY_LUASM2 = _
                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString) * _
                        '    CDbl(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString))


                        oProd1.UserFields.Fields.Item("U_SO_Luas_M2").Value = _
                        Left(CStr( _
                            Math.Round( _
                              (IIf(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oSOToMFGGrid.DataTable.GetValue(15, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
                              IIf(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
                              IIf(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oSOToMFGGrid.DataTable.GetValue(11, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString).ToString)) / 10000) _
                              , 4) _
                            ) _
                        , 10)

                        'oprod1.UserFields.Fields.Item("U_PoD_Pcm").Value = "p100cm"
                        'oprod1.UserFields.Fields.Item("U_PdO_Lcm").Value = "L90cm"
                        'oprod1.UserFields.Fields.Item("U_PdO_Bentuk").Value = "segi"
                        'oprod1.UserFields.Fields.Item("U_NBS_OnHoldReason").Value = "test123"

                        oProdLine1 = oProd1.Lines

                        ' Generate one line - Dummy item
                        oProdLine1.ItemNo = "XDUMMY"
                        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                        oProdLine1.Warehouse = "SRV-DL"


                        ' In Case of JOB 01 Exists
                        'If oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
                        '    oProdLine1.Add()
                        '    ' Job 01 - sample: XTP1 
                        '    'oprodLine1.ItemNo = "F12CLXXMM51003048"
                        '    oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(16, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        '    'oProdLine1.ItemNo = "XSE1"

                        '    'oProdLine1.Warehouse = "RM-PRD"
                        '    'oProdLine1.Warehouse = "SRV-DL"

                        '    oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush

                        '    'SO JOb 01 - DC dummy code (Direct Cost)
                        '    If oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
                        '        oProdLine1.Add()
                        '        oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(17, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        '        oProdLine1.Warehouse = "SRV-DL"
                        '        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                        '    End If
                        '    'SO JOb 01 - IC dummy code (Indirect Cost)
                        '    If oSOToMFGGrid.DataTable.GetValue(18, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
                        '        oProdLine1.Add()
                        '        oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(18, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        '        oProdLine1.Warehouse = "SRV-DL"
                        '        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                        '    End If
                        '    'SO JOb 01 - FOH dummy code (FOH)
                        '    If oSOToMFGGrid.DataTable.GetValue(19, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) <> "" Then
                        '        oProdLine1.Add()
                        '        oProdLine1.ItemNo = oSOToMFGGrid.DataTable.GetValue(19, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        '        oProdLine1.Warehouse = "SRV-DL"
                        '        oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                        '    End If

                        'End If





                        'MsgBox(GC.GetTotalMemory(True))

                        lRetCode = oProd1.Add()


                        If lRetCode <> 0 Then
                            'oCompany.GetLastError(lErrCode, sErrMsg)
                            vCompany.GetLastError(lErrCode, sErrMsg)
                            SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        Else

                            Dim PdOno As String = ""
                            Dim tmpKey As Double
                            Dim vSOLine As Long

                            vCompany.GetNewObjectCode(tmpKey)
                            vCompany.GetNewObjectCode(PdOno)
                            tmpKey = Convert.ToInt32(PdOno)

                            ' !!!! Make sure before create another object type-> clear previous/current object type.
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
                            oProdLine1 = Nothing

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
                            oProd1 = Nothing

                            'GC.Collect()

                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)


                            oSalesOrder = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                            'oSalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                            oSalesOrder.GetByKey(oSOToMFGGrid.DataTable.GetValue(3, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString))
                            'oSalesOrder.UserFields.Fields.Item("U_MIS_ReasonCode").Value = "T1"

                            vSOLine = oSOToMFGGrid.DataTable.GetValue(5, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) - 1
                            oSalesOrderLines = oSalesOrder.Lines
                            'oSalesOrderLines.SetCurrentLine(oSOToMFGGrid.DataTable.GetValue(5, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString) - 1)
                            oSalesOrderLines.SetCurrentLine(vSOLine)
                            'oSalesOrderLines.SetCurrentLine(10)

                            oSalesOrderLines.UserFields.Fields.Item("U_MIS_SupplyWith").Value = "M"
                            oSalesOrderLines.UserFields.Fields.Item("U_MIS_ReleasePdOFlag").Value = "Y"
                            oSalesOrderLines.UserFields.Fields.Item("U_MIS_PdONum").Value = PdOno  '"1010025702"

                            oSalesOrder.Update()

                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrderLines)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrder)
                            oSalesOrder = Nothing
                            'oSalesOrderLines = Nothing

                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            'oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

                        End If

                        SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                    oRS = Nothing

                    ' by Toin 2011-02-09 Check Duplicate PdO before Generate PdO

                    vCompany.Disconnect()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(vCompany)
                    vCompany = Nothing

                    GC.Collect()
                    'MsgBox(GC.GetTotalMemory(True))


                    'MsgBox("generating... PdO; DONE!!!")

                Else
                    Exit For
                End If
            Next

        End If  ' Checking PdO Series

        'End If
        'Catch ex As Exception
        '    MsgBox(oCompany.GetLastErrorDescription)
        '    MsgBox(ex.Message)
        '    Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

        'End Try

        'Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        'Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

        'MsgBox("Begin trx: generating... PdO")
        SBO_Application.MessageBox("Generating PdO.... Finished !!! ", 1, "Ok")

        'Begin Trxs

        'Call oCompany.StartTransaction()

        'Dim oSalesOrder As SAPbobsCOM.Documents
        'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines

        'oSalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        'oSalesOrder.GetByKey(252)
        'oSalesOrder.UserFields.Fields.Item("U_NBS_Range").Value = "123Tes321"

        'oSalesOrder.Lines.SetCurrentLine(0)
        'oSalesOrder.Lines.UserFields.Fields.Item("U_BacthNum").Value = "223344"

        ''### How to update UDF using recordset
        'Dim oRS As SAPbobsCOM.Recordset
        ''Dim vCompany As SAPbobsCOM.Company
        ''vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        'oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ''oRS.DoQuery("UPDATE RDR1 SET U_BatchNum = 'b 321' WHERE DocEntry = 252 and LineNum = 0")
        'oRS.DoQuery("UPDATE RDR1 SET U_bacthNum = 'b321' where docentry = 249 and linenum = 0")

        'If oRS.RecordCount <> 0 Then
        '    MsgBox("ada record!")
        'End If
        '### How to update UDF using recordset

        'oSalesOrderLines = oSalesOrder.Lines
        'oSalesOrderLines.UserFields.Fields.Item("U_BacthNum").Value = "b 321"

        'lRetCode = oSalesOrder.Update()
        'If lRetCode <> 0 Then
        '    oCompany.GetLastError(lErrCode, sErrMsg)
        '    SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)

        '    Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

        'End If

        'For idx = 0 To oSOToMFGGrid.Rows.SelectedRows.Count - 1
        '    MsgBox("selected row#:" & idx.ToString & _
        '           "; selectedrow->row#: " & oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
        '           & "DocEntry: " & oSOToMFGGrid.DataTable.GetValue(2, oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

        'Next

        'Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

        'MsgBox("generating... PdO; DONE!!!")


        'Disconnect Company Object & Release Resource
        'Call oCompany.Disconnect()
        'oCompany = Nothing

        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

    'Karno Generate PDO Status
    Private Sub GeneratePdOStatus(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler

        Dim oPDOStatusGrid As SAPbouiCOM.Grid

        Dim idx As Long

        oPDOStatusGrid = oForm.Items.Item("myGridPDO").Specific

        'GRID - Order by column checkbox
        oPDOStatusGrid.Columns.Item("Release PdO").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        'Loop only selected/checked in grid rows and exit.
        For idx = oPDOStatusGrid.Rows.Count - 1 To 0 Step -1

            If oPDOStatusGrid.DataTable.GetValue(0, oPDOStatusGrid.GetDataTableRowIndex(idx)) = "Y" Then

                Dim oPDOStatus As SAPbobsCOM.ProductionOrders = Nothing

                Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
                Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines

                Dim vCompany As SAPbobsCOM.Company = Nothing
                Dim sCookie As String
                Dim sConnectionContext As String

                Dim isconnect As Long
                Dim errConnect As String = ""

                Try
                    vCompany = New SAPbobsCOM.Company
                    'Dim sCookie As String = vCompany.GetContextCookie
                    'Dim sConnectionContext As String
                    sCookie = vCompany.GetContextCookie
                    sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
                    vCompany.SetSboLoginContext(sConnectionContext)
                    isconnect = vCompany.Connect()

                    'If vCompany.Connect() <> 0 Then
                    If isconnect <> 0 Then
                        End
                    End If
                Catch ex As Exception
                    End
                End Try

                vCompany.StartTransaction()

                oProd1 = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                'oProd1.PlannedQuantity = 2

                oProd1.ItemNo = oPDOStatusGrid.DataTable.GetValue(7, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                oProd1.PlannedQuantity = oPDOStatusGrid.DataTable.GetValue(8, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                oProd1.PostingDate = oPDOStatusGrid.DataTable.GetValue(1, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                oProd1.DueDate = oPDOStatusGrid.DataTable.GetValue(1, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased

                'oprod1.Warehouse = "01"
                'oProd1.Warehouse = "FG-001"
                'oProd1.Warehouse = oPDOStatusGrid.DataTable.GetValue(12, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                oProd1.CustomerCode = oPDOStatusGrid.DataTable.GetValue(14, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)
                oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
                ' so docnum
                oProd1.ProductionOrderOriginEntry = oPDOStatusGrid.DataTable.GetValue(4, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString)

                oProdLine1 = oProd1.Lines

                'lRetCode = oProd1.Add()

                Dim PdOno As String = ""

                If lRetCode <> 0 Then
                    vCompany.GetLastError(lErrCode, sErrMsg)
                    SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                Else

                    ' !!!! Make sure before create another object type-> clear previous/current object type.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
                    oProdLine1 = Nothing

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
                    oProd1 = Nothing

                    oPDOStatus = vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    oPDOStatus.GetByKey(oPDOStatusGrid.DataTable.GetValue(2, oPDOStatusGrid.GetDataTableRowIndex(idx).ToString))


                    oPDOStatus.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                    'oPDOStatus.UserFields.Fields.Item(12).Value = "12"
                    oPDOStatus.UserFields.Fields.Item("U_MIS_Progress").Value = "Released"

                    oPDOStatus.Update()


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPDOStatus)
                    oPDOStatus = Nothing

                    'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)

                End If


                vCompany.Disconnect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(vCompany)
                vCompany = Nothing

                GC.Collect()
            Else
                Exit For
            End If
        Next

        'MsgBox("Begin trx: generating... PdO")
        SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)


        'MsgBox("generating... PdO; DONE!!!")

        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub


    Private Sub RearrangeGrid(ByVal oForm As SAPbouiCOM.Form)

        Dim oColumn As SAPbouiCOM.EditTextColumn

        Dim oSOToMFGGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)

        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

        oSOToMFGGrid.RowHeaders.Width = 50

        'Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        oColumn = oSOToMFGGrid.Columns.Item("Cust. Code")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner '"2" -> BP Master
        oColumn.Editable = False

        oColumn = oSOToMFGGrid.Columns.Item("DocEntry")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
        oColumn.Editable = False

        oColumn = oSOToMFGGrid.Columns.Item("FG")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items '"2"
        oColumn.Editable = False


        oSOToMFGGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oSOToMFGGrid.Columns.Item("Release PdO").TitleObject.Sortable = True

        oSOToMFGGrid.Columns.Item("DocEntry").Width = 60
        oSOToMFGGrid.Columns.Item("Cust. Code").Width = 130


        oColumn = oSOToMFGGrid.Columns.Item("Customer Name")
        oColumn.Editable = False

        oSOToMFGGrid.Columns.Item("SO Date").Width = 80
        oSOToMFGGrid.Columns.Item("SO Date").TitleObject.Sortable = True


        oSOToMFGGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        ' Set Total Row count in colum title/header
        'oSOToMFGGrid.Columns.Item(0).TitleObject.Caption = oSOToMFGGrid.Rows.Count.ToString


        If oForm.DataSources.DataTables.Item(0).Rows.Count <> 0 _
        And oSOToMFGGrid.DataTable.GetValue(0, 0) <> "" Then
            oForm.Items.Item("cmdGenPdO").Enabled = True
        Else
            oForm.Items.Item("cmdGenPdO").Enabled = False
        End If

        oSOToMFGGrid.Columns.Item("#").Editable = False
        oSOToMFGGrid.Columns.Item(1).Editable = True
        oSOToMFGGrid.Columns.Item("SO Date").Editable = False
        oSOToMFGGrid.Columns.Item("DocEntry").Editable = False
        oSOToMFGGrid.Columns.Item("DocNum").Editable = False
        oSOToMFGGrid.Columns.Item("SO Line").Editable = False
        oSOToMFGGrid.Columns.Item("Sales Rep.").Editable = False
        oSOToMFGGrid.Columns.Item("Cust. Code").Editable = False
        oSOToMFGGrid.Columns.Item("FG").Editable = False
        oSOToMFGGrid.Columns.Item("FGName").Editable = False
        oSOToMFGGrid.Columns.Item("Quantity").Editable = False
        oSOToMFGGrid.Columns.Item("UOM").Editable = False
        oSOToMFGGrid.Columns.Item("Exp Delivery Date").Editable = False
        oSOToMFGGrid.Columns.Item("WhsCode").Editable = False
        oSOToMFGGrid.Columns.Item("PanjangInCm").Editable = False
        oSOToMFGGrid.Columns.Item("LebarInCm").Editable = False
        oSOToMFGGrid.Columns.Item("SO_Bentuk").Editable = False


        oSOToMFGGrid.RowHeaders.Width = 20
        oSOToMFGGrid.Columns.Item("#").Width = 30
        oSOToMFGGrid.Columns.Item(1).Width = 20
        oSOToMFGGrid.Columns.Item("SO Date").Width = 60
        oSOToMFGGrid.Columns.Item("DocEntry").Width = 60
        oSOToMFGGrid.Columns.Item("DocNum").Width = 60
        oSOToMFGGrid.Columns.Item("SO Line").Width = 30
        oSOToMFGGrid.Columns.Item("Cust. Code").Width = 80
        oSOToMFGGrid.Columns.Item("FG").Width = 100
        oSOToMFGGrid.Columns.Item("Exp Delivery Date").Width = 80
        oSOToMFGGrid.Columns.Item("WhsCode").Width = 50
        oSOToMFGGrid.Columns.Item("PanjangInCm").Width = 50
        oSOToMFGGrid.Columns.Item("LebarInCm").Width = 50
        oSOToMFGGrid.Columns.Item("SO_Bentuk").Width = 80



        Dim sboDate As String
        Dim dDate As DateTime

        'dDate = DateTime.Now

        'sbo formatdate
        sboDate = oMis_Utils.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oSOToMFGGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Private Sub SetApplication()

        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        SboGuiApi = New SAPbouiCOM.SboGuiApi

        '// by following the steps specified above, the following
        '// statment should be suficient for either development or run mode

        'sConnectionString = Command()
        ' 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056

        '#If DEBUG Then
        '    sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        '#Else
        '   sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        '#End If

        'sConnectionString = Environment.GetCommandLineArgs.GetValue(1) '

        'If Environment.GetCommandLineArgs.Length = 1 Then
        '    sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        'Else
        '    sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        'End If


        'sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"


        'sConnectionString = "5645523035496D706C656D656E746174696F6E3A59313931303035313531383699469FA92C3C9A964A219C5862952A90D911E9" 'Environment.GetCommandLineArgs.GetValue(1)'
        Try
            SboGuiApi.Connect(sConnectionString)
            '// connect to a running SBO Application
            '// get an initialized application object
            SBO_Application = SboGuiApi.GetApplication()
        Catch ex As Exception
            MsgBox("Make Sure That SAP Business One Application is running!!! ", MsgBoxStyle.Information)
            End
        End Try

        ''// connect to a running SBO Application

        'SboGuiApi.Connect(sConnectionString)

        ''// get an initialized application object

        'SBO_Application = SboGuiApi.GetApplication()

        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String
        '        Dim lRetCode As Integer

        '// First initialize the Company object

        oCompany = New SAPbobsCOM.Company

        '// Acquire the connection context cookie from the DI API.
        sCookie = oCompany.GetContextCookie

        '// Retrieve the connection context string from the UI API using the
        '// acquired cookie.
        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        '// before setting the SBO Login Context make sure the company is not
        '// connected

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If

        '// Set the connection context information to the DI API.
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function



    Private Sub AddCFL1(ByVal oForm As SAPbouiCOM.Form, ByVal oLinkedObject As SAPbouiCOM.BoLinkedObject, _
                    ByVal CFLtxt As String, ByVal CFLbtn As String, _
                    ByVal CFLCondField As String, ByVal CFLCondFieldValue As String)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFLConds As SAPbouiCOM.Conditions
            Dim oCFLCond As SAPbouiCOM.Condition


            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Add 2 CFL
            'one for button (windows popup) & one for edit textbox
            oCFLCreationParams.MultiSelection = False
            '            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            Dim oLinkedObjectType As SAPbouiCOM.BoLinkedObject
            oLinkedObjectType = oLinkedObject
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner ' "2"-> BP Master
            oCFLCreationParams.UniqueID = CFLtxt ' "CFL1" -> txtbox cfl Field

            oCFL = oCFLs.Add(oCFLCreationParams)

            'Add conditions to CFL1
            oCFLConds = oCFL.GetConditions()

            oCFLCond = oCFLConds.Add()
            oCFLCond.Alias = CFLCondField ' "CardType" -> BP Master where CardType = ??
            oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCFLCond.CondVal = CFLCondFieldValue ' "C" -> CardType value = C -> BP Customer data 
            oCFL.SetConditions(oCFLConds)

            oCFLCreationParams.UniqueID = CFLbtn ' "CFL2" -> button CFL field
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub



    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        'karno optimize
        If pVal.BeforeAction = False Then
            If pVal.FormTypeEx = "ListOptimize" Then

                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    Dim oForm As SAPbouiCOM.Form = Nothing
                    oForm = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Select Case pVal.ItemUID
                            Case "OptimNo"
                                Dim oForm As SAPbouiCOM.Form = Nothing
                                oForm = SBO_Application.Forms.Item("ListOptimize")

                                Dim ListOptimizeQuery As String
                                'Dim oListOptimGrid As SAPbouiCOM.Grid = Nothing
                                'Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

                                'Dim oMis_Utils As MIS_Utils

                                'oMis_Utils = New MIS_Utils

                                'oListOptimGrid = oForm.Items.Item("myGrid2").Specific

                                'ListOptimizeQuery = "select T0.DocNum, T2.Visorder + 1 RowNo, T2.ItemCode,  T0.U_MIS_OPTNUM,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                '                    "From [@MIS_OPTIM] T0 INNER JOIN OWOR T1 ON T0.U_MIS_OPTNUM = T1.Docnum INNER JOIN WOR1 T2 ON T1.Docentry = T2.Docentry " & _
                                '                    "where T1.STATUS = 'R' AND LEFT(T2.ItemCode,1) <> 'X'"

                                If oForm.Items.Item("OptimNo").Specific.string = "" Then
                                    ListOptimizeQuery = "select T0.DocNum [Number Optimize], T2.DocNum [PDO Number], T3.Visorder + 1 [Row No], T3.ItemCode [Item Code],  " & _
                                                        "T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                                        "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                                        "From [@MIS_OPTIM] T0 " & _
                                                        "join [@MIS_OPTIML] T1 " & _
                                                        "ON T0.DocEntry = T1.DocEntry " & _
                                                        "INNER JOIN OWOR T2 " & _
                                                        "ON T2.DocNum = T1.U_MIS_PdONum " & _
                                                        "INNER JOIN WOR1 T3 " & _
                                                        "ON T2.Docentry = T3.Docentry " & _
                                                        "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X'"

                                Else
                                    ListOptimizeQuery = "select T0.DocNum [Number Optimize], T2.DocNum [PDO Number], T3.Visorder + 1 [Row No], T3.ItemCode [Item Code],  " & _
                                                        "T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                                        "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                                        "From [@MIS_OPTIM] T0 " & _
                                                        "join [@MIS_OPTIML] T1 " & _
                                                        "ON T0.DocEntry = T1.DocEntry " & _
                                                        "INNER JOIN OWOR T2 " & _
                                                        "ON T2.DocNum = T1.U_MIS_PdONum " & _
                                                        "INNER JOIN WOR1 T3 " & _
                                                        "ON T2.Docentry = T3.Docentry " & _
                                                        "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' AND T0.DocNum LIKE '" & oForm.Items.Item("OptimNo").Specific.string & "%'"
                                End If

                                oForm.DataSources.DataTables.Item("ListOptim").ExecuteQuery(ListOptimizeQuery)
                                'oListOptimGrid.DataTable = oForm.DataSources.DataTables.Item("ListOptim")
                        End Select


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If pVal.ItemUID = "BtnChoose" Then
                            Dim oForm As SAPbouiCOM.Form = Nothing
                            oForm = SBO_Application.Forms.Item("ListOptimize")
                            Dim oGrid As SAPbouiCOM.Grid = Nothing
                            Dim oColumns As SAPbouiCOM.Columns = Nothing
                            Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                            Dim StrSql As String
                            Dim DocNum As String
                            Dim Row As Integer

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                            oMatrix = objFormProductionIssue.Items.Item("13").Specific
                            oColumns = oMatrix.Columns
                            'karno not yet

                            oGrid = oForm.Items.Item("myGrid2").Specific


                            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            For i As Integer = 0 To oGrid.Rows.Count - 1
                                'For i As Integer = 0 To oGrid.Rows.SelectedRows.Count

                                If oGrid.Rows.IsSelected(i) = True Then
                                    DocNum = oGrid.DataTable.GetValue("PDO Number", oGrid.GetDataTableRowIndex(i)).ToString

                                    'StrSql = "SELECT T0.DocNum DocNum, T0.U_MIS_OPTNUM OrderNo, T0.U_MIS_QtyInLembar Qty FROM [@MIS_OPTIM] T0 INNER JOIN [@MIS_OPTIM] T1 ON T0.DocEntry = T1.DocEntry " & _
                                    '        "INNER JOIN OWOR T2 ON T0.U_MIS_OPTNUM = T2.DocNum WHERE T0.U_MIS_OPTNUM = '" & DocNum & "'"

                                    'StrSql = "select T0.DocNum, T2.Visorder + 1 RowNo, T2.itemcode, T0.U_MIS_QtyInLembar Qty, T0.U_MIS_OPTNUM OrderNo,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                    '                    "From [@MIS_OPTIM] T0 INNER JOIN OWOR T1 ON T0.U_MIS_OPTNUM = T1.Docnum INNER JOIN WOR1 T2 ON T1.Docentry = T2.Docentry " & _
                                    '                    "where T1.STATUS = 'R' AND LEFT(T2.ItemCode,1) <> 'X' AND T0.U_MIS_OPTNUM = '" & DocNum & "' "

                                    StrSql = "select T2.DocNum OrderNo, T3.linenum + 1  RowNo, T3.ItemCode, T1.U_MIS_QtPlanPdoIssue Qty,  " & _
                                            "T0.U_MIS_QtyInLembar Lembar, T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                                            "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                                            "From [@MIS_OPTIM] T0 " & _
                                            "join [@MIS_OPTIML] T1 " & _
                                            "ON T0.DocEntry = T1.DocEntry " & _
                                            "INNER JOIN OWOR T2 " & _
                                            "ON T2.DocNum = T1.U_MIS_PdONum " & _
                                            "INNER JOIN WOR1 T3 " & _
                                            "ON T2.Docentry = T3.Docentry " & _
                                            "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' AND T2.DocNum = '" & DocNum & "' "




                                    objRecSet.DoQuery(StrSql)

                                    If objRecSet.RecordCount > 0 Then
                                        For Row = 1 To objRecSet.RecordCount
                                            If objRecSet.Fields.Item("Qty").Value = 0.0 Then
                                                SBO_Application.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                oForm.Close()
                                                Exit Sub
                                            Else
                                                oColumns.Item("61").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("OrderNo").Value
                                                oColumns.Item("60").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("RowNo").Value
                                                oColumns.Item("9").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Qty").Value
                                            End If

                                        Next
                                    End If

                                End If
                            Next

                            oForm.Close()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                        End If
                End Select

            End If

            ' karno optimize production
            If pVal.FormTypeEx = "ListOptimizePro" Then

                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    Dim oForm As SAPbouiCOM.Form = Nothing
                    oForm = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        If pVal.ItemUID = "BtnChoose" Then
                            Dim oForm As SAPbouiCOM.Form = Nothing
                            oForm = SBO_Application.Forms.Item("ListOptimizePro")

                            Dim oGrid As SAPbouiCOM.Grid = Nothing
                            Dim oColumns As SAPbouiCOM.Columns = Nothing
                            Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                            Dim StrSql As String
                            Dim DocNum As String
                            Dim Row As Integer

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                            oMatrix = objFormProduction.Items.Item("37").Specific
                            oColumns = oMatrix.Columns
                            'karno not yet

                            oGrid = oForm.Items.Item("myGrid2").Specific


                            For i As Integer = 0 To oGrid.Rows.SelectedRows.Count
                                If oGrid.Rows.IsSelected(i) = True Then
                                    DocNum = oGrid.DataTable.GetValue("U_MIS_OPTNUM", oGrid.GetDataTableRowIndex(i)).ToString

                                    StrSql = "SELECT T0.U_MIS_ItemCode Item, T0.DocNum DocNum, T0.U_MIS_OPTNUM OrderNo, T0.U_MIS_QtyInLembar Qty FROM [@MIS_OPTIM] T0 INNER JOIN [@MIS_OPTIM] T1 ON T0.DocEntry = T1.DocEntry " & _
                                            "INNER JOIN OWOR T2 ON T0.U_MIS_OPTNUM = T2.DocNum WHERE T0.U_MIS_OPTNUM = '" & DocNum & "'"
                                    objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRecSet.DoQuery(StrSql)

                                    If objRecSet.RecordCount > 0 Then
                                        For Row = 1 To objRecSet.RecordCount

                                            oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("Item").Value
                                        Next
                                    End If
                                End If
                            Next

                            oForm.Close()

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                        End If
                End Select

            End If

            If pVal.FormTypeEx = ProductionUDF_FormId Then
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    'objFormProductionUDF = SBO_Application.Forms.Item(pVal.FormUID)
                End If
            End If

            ' karno Production
            If pVal.FormTypeEx = Production_FormId Then
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    objFormProduction = SBO_Application.Forms.Item(pVal.FormUID)
                End If


                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Select Case pVal.ColUID
                            Case "U_NBS_MatlQty"
                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing

                                oMatrix = objFormProduction.Items.Item("37").Specific
                                oColumns = oMatrix.Columns

                                If oColumns.Item("U_MIS_OptNum").Cells.Item(pVal.Row).Specific.value = "" Then
                                    If oColumns.Item("4").Cells.Item(pVal.Row).Specific.value = "" Then
                                        SBO_Application.SetStatusBarMessage("Please Fill Item Code!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        Exit Sub
                                    Else

                                        If oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                            SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        ElseIf oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value = "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value = "0.0" Then
                                            SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                            SBO_Application.SetStatusBarMessage("Material Quantity Must Blank AND Run Time Must Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                            oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                            Exit Sub
                                        Else
                                            If Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) <> "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                            Else
                                                SBO_Application.SetStatusBarMessage("Material Quantity Must Fill AND Run Time Must Blank!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)

                            Case "U_NBS_RunTime"
                                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                                Dim oColumns As SAPbouiCOM.Columns = Nothing

                                oMatrix = objFormProduction.Items.Item("37").Specific
                                oColumns = oMatrix.Columns

                                If oColumns.Item("U_MIS_OptNum").Cells.Item(pVal.Row).Specific.value = "" Then
                                    If oColumns.Item("4").Cells.Item(pVal.Row).Specific.value = "" Then
                                        SBO_Application.SetStatusBarMessage("Please Fill Item Code!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        BubbleEvent = False
                                        Exit Sub

                                    Else
                                        If oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                            SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        ElseIf oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value = "0.0" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value = "0.0" Then
                                            SBO_Application.SetStatusBarMessage("Please Fill Material Quantity OR Run Time!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                            SBO_Application.SetStatusBarMessage("Material Quantity Must Blank AND Run Time Must Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                            Exit Sub
                                        ElseIf Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) = "X" And oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                            oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_RunTime").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                            Exit Sub
                                        Else
                                            If Left(oColumns.Item("4").Cells.Item(pVal.Row).Specific.value, 1) <> "X" And oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value <> "0.0" Then
                                                oColumns.Item("14").Cells.Item(pVal.Row).Specific.value = oColumns.Item("U_NBS_MatlQty").Cells.Item(pVal.Row).Specific.value * objFormProduction.Items.Item("12").Specific.value
                                            Else
                                                SBO_Application.SetStatusBarMessage("Material Quantity Must Fill AND Run Time Must Blank!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)



                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Dim oItem As SAPbouiCOM.Item

                        oItem = objFormProduction.Items.Add("BtnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oItem.Left = 145
                        oItem.Top = 395
                        oItem.Width = 150
                        oItem.Height = 19
                        oItem.Specific.caption = "Copy From Optimize"



                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "BtnCopy" Then
                            Dim StrSql As String
                            'Dim DocNum As String
                            Dim PlannedQty As Double
                            Dim PdoNumber As String


                            PdoNumber = objFormProduction.Items.Item("18").Specific.value
                            'DocNum = objFormProductionUDF.Items.Item("U_MIS_OptNum").Specific.string
                            PlannedQty = objFormProduction.Items.Item("12").Specific.value

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                            Dim oColumns As SAPbouiCOM.Columns = Nothing
                            Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                            oMatrix = objFormProduction.Items.Item("37").Specific
                            oColumns = oMatrix.Columns
                            'karno not yet

                            'oGrid = oForm.Items.Item("myGrid2").Specific
                            '201100146 1011001880

                            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'StrSql = "Select T0.U_MIS_ItemCode ItemCode, " & _
                            '        "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyinLembar	PlannedQty, " & _
                            '        "'Y' Flag " & _
                            '        "FROM [@MIS_OPTIM] T0 " & _
                            '        "JOIN [@MIS_OPTIML] T1 " & _
                            '        "ON T0.docentry = T1.docentry " & _
                            '        "WHERE T0.DocNum = '" & DocNum & "' " & _
                            '        "AND T1.U_MIS_PdONum = " & PdoNumber & " "

                            StrSql = "Select T0.U_MIS_ItemCode ItemCode, " & _
                                    "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyinLembar	PlannedQty, " & _
                                    "'Y' Flag, T0.DocNum OptimNumber " & _
                                    "FROM [@MIS_OPTIM] T0 " & _
                                    "JOIN [@MIS_OPTIML] T1 " & _
                                    "ON T0.docentry = T1.docentry " & _
                                    "WHERE T1.U_MIS_PdONum = " & PdoNumber & " "

                            objRecSet.DoQuery(StrSql)

                            If objRecSet.RecordCount > 0 Then
                                objRecSet.MoveFirst()
                                For Row = 1 To objRecSet.RecordCount
                                    If objRecSet.Fields.Item("PlannedQty").Value = 0.0 Then
                                        SBO_Application.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        objFormProductionIssue.Close()
                                        Exit Sub
                                    Else
                                        'oColumns.Item("U_MIS_OptNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("OptimNumber").Value
                                        'oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = Math.Round(objRecSet.Fields.Item("PlannedQty").Value / PlannedQty, 4)
                                        oColumns.Item("4").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("ItemCode").Value
                                        oColumns.Item("14").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("PlannedQty").Value
                                        oColumns.Item("U_MIS_PdOGenFlag").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Flag").Value
                                        oColumns.Item("U_NBS_MatlQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = Math.Round(objRecSet.Fields.Item("PlannedQty").Value / PlannedQty, 4)

                                        'Reassignment Again! Dari Optimize data harus masuk dulu ke Planned Qty baru nanti kalkulasi dptkan nilai material qty. Jadi Optimize Planned qty = PdO Planned Qty 
                                        oColumns.Item("14").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("PlannedQty").Value
                                        oColumns.Item("U_MIS_OptNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("OptimNumber").Value
                                    End If
                                    objRecSet.MoveNext()
                                Next
                            Else
                                SBO_Application.SetStatusBarMessage("Please Check Optimazation Number", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'End If
                            'Next

                            'objFormProductionIssue.Close()

                            objFormProduction.Items.Item("BtnCopy").Enabled = False

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProduction)
                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProductionUDF)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)

                        End If
                End Select
            End If
            'karno Copy Optim
            If pVal.FormTypeEx = ProductionIssueUDF_FormId Then
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    objFormProductionIssueUDF = SBO_Application.Forms.Item(pVal.FormUID)
                End If
            End If

            If pVal.FormTypeEx = ProductionIssue_FormId Then
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    objFormProductionIssue = SBO_Application.Forms.Item(pVal.FormUID)
                End If

                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Dim oItem As SAPbouiCOM.Item
                        'Dim oListOptimizeGrid As SAPbouiCOM.Grid = Nothing

                        oItem = objFormProductionIssue.Items.Add("BtnCopy", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oItem.Left = 145
                        oItem.Top = 318
                        oItem.Width = 150
                        oItem.Height = 19
                        oItem.Specific.caption = "Copy From Optimize"

                        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oListOptimizeGrid)

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "BtnCopy" Then
                            Dim StrSql As String
                            Dim DocNum As String = ""
                            Dim OptimLembar As Integer
                            Dim Row As Integer

                            If objFormProductionIssueUDF.Items.Item("U_MIS_OptNum").Specific.string = "" Then
                                SBO_Application.SetStatusBarMessage("Optimize Number Must Fill", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            Else
                                If objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.value = "" Then
                                    objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.value = 1
                                    OptimLembar = objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.value
                                Else
                                    OptimLembar = objFormProductionIssueUDF.Items.Item("U_MIS_OptLmbr").Specific.string
                                    DocNum = objFormProductionIssueUDF.Items.Item("U_MIS_OptNum").Specific.string
                                End If
                            End If

                            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                            Dim oColumns As SAPbouiCOM.Columns = Nothing
                            Dim objRecSet As SAPbobsCOM.Recordset = Nothing

                            oMatrix = objFormProductionIssue.Items.Item("13").Specific
                            oColumns = oMatrix.Columns

                            objRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            StrSql = "select T2.DocNum OrderNo, T3.linenum + 1 RowNo, T3.ItemCode, " & _
                            "CASE T0.U_MIS_ItemCode WHEN T3.ItemCode THEN T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & " " & _
                            "ELSE (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty END Qty,  T0.U_MIS_QtyInLembar Lembar, " & _
                            "T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                            "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar MaximumQtyLembarPdoIssue, " & _
                            "SUM(ISNULL(T4.Quantity,0)) QuantityPdoIssue " & _
                            "From [@MIS_OPTIM] T0 join [@MIS_OPTIML] T1  " & _
                            "ON T0.DocEntry = T1.DocEntry INNER JOIN OWOR T2  " & _
                            "ON T2.DocNum = T1.U_MIS_PdONum INNER JOIN WOR1 T3  " & _
                            "ON T2.Docentry = T3.Docentry  AND T0.DocNum = T3.U_MIS_OptNum " & _
                            "AND ROUND((1 / T0.U_MIS_QtyInLembar) * T3.PlannedQty, 4) = ROUND(T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar / T0.U_MIS_QtyInLembar, 4) " & _
                            "LEFT JOIN IGE1 T4 ON T2.DocNum = T4.BaseRef " & _
                            "AND T3.ItemCode = T4.ItemCode " & _
                            "AND T3.LineNum = T4.BaseLine " & _
                            "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' " & _
                            "AND T0.DocNum = '" & DocNum & "' " & _
                            "GROUP BY T2.DocNum, T3.linenum + 1, T3.ItemCode, " & _
                            "T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ", (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty,  T0.U_MIS_QtyInLembar,  " & _
                            "T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                            "T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar " & _
                            "HAVING (T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ") + SUM(ISNULL(T4.Quantity,0)) <= T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar "


                            'StrSql = "select T2.DocNum OrderNo, T3.linenum + 1 RowNo, T3.ItemCode, " & _
                            '"CASE T0.U_MIS_ItemCode WHEN T3.ItemCode THEN T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & " " & _
                            '"ELSE (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty END Qty,  T0.U_MIS_QtyInLembar Lembar, " & _
                            '"T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            '"T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                            '"T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar MaximumQtyLembarPdoIssue, " & _
                            '"SUM(ISNULL(T4.Quantity,0)) QuantityPdoIssue " & _
                            '"From [@MIS_OPTIM] T0 join [@MIS_OPTIML] T1  " & _
                            '"ON T0.DocEntry = T1.DocEntry INNER JOIN OWOR T2  " & _
                            '"ON T2.DocNum = T1.U_MIS_PdONum INNER JOIN WOR1 T3  " & _
                            '"ON T2.Docentry = T3.Docentry  AND T0.DocNum = T3.U_MIS_OptNum " & _
                            '"AND ROUND((1 / T0.U_MIS_QtyInLembar) * T3.PlannedQty, 4) = ROUND(T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar, 4) " & _
                            '"LEFT JOIN IGE1 T4 ON T2.DocNum = T4.BaseRef " & _
                            '"AND T3.ItemCode = T4.ItemCode " & _
                            '"AND T3.LineNum = T4.BaseLine " & _
                            '"where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' " & _
                            '"AND T0.DocNum = '" & DocNum & "' " & _
                            '"GROUP BY T2.DocNum, T3.linenum + 1, T3.ItemCode, " & _
                            '"T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ", (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty,  T0.U_MIS_QtyInLembar,  " & _
                            '"T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            '"T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                            '"T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar " & _
                            '"HAVING (T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ") + SUM(ISNULL(T4.Quantity,0)) <= T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar "

                            'StrSql = "select T2.DocNum OrderNo, T3.linenum + 1 RowNo, T3.ItemCode, " & _
                            '"CASE T0.U_MIS_ItemCode WHEN T3.ItemCode THEN T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & " " & _
                            '"ELSE (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty END Qty,  T0.U_MIS_QtyInLembar Lembar, " & _
                            '"T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            '"T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                            '"T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar MaximumQtyLembarPdoIssue, " & _
                            '"SUM(ISNULL(T4.Quantity,0)) QuantityPdoIssue " & _
                            '"From [@MIS_OPTIM] T0 join [@MIS_OPTIML] T1  " & _
                            '"ON T0.DocEntry = T1.DocEntry INNER JOIN OWOR T2  " & _
                            '"ON T2.DocNum = T1.U_MIS_PdONum INNER JOIN WOR1 T3  " & _
                            '"ON T2.Docentry = T3.Docentry  AND T0.DocNum = T3.U_MIS_OptNum " & _
                            '"AND T3.ItemCode = T4.ItemCode " & _
                            '"AND T3.LineNum = T4.BaseLine " & _
                            '"where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' " & _
                            '"AND T0.DocNum = '" & DocNum & "' " & _
                            '"GROUP BY T2.DocNum, T3.linenum + 1, T3.ItemCode, " & _
                            '"T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ", (" & OptimLembar & " / T0.U_MIS_QtyInLembar) * T3.PlannedQty,  T0.U_MIS_QtyInLembar,  " & _
                            '"T0.U_MIS_OPTNUM, T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            '"T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2, " & _
                            '"T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar " & _
                            '"HAVING (T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & ") + SUM(ISNULL(T4.Quantity,0)) <= T1.U_MIS_QtPlanPdoIssue * T0.U_MIS_QtyInLembar "

                            'StrSql = "select T2.DocNum OrderNo, T3.linenum + 1  RowNo, T3.ItemCode, T1.U_MIS_QtPlanPdoIssue * " & OptimLembar & " Qty,  " & _
                            '                                    "T0.U_MIS_QtyInLembar Lembar, T0.U_MIS_OPTNUM ,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, " & _
                            '                                    "T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                            '                                    "From [@MIS_OPTIM] T0 " & _
                            '                                    "join [@MIS_OPTIML] T1 " & _
                            '                                    "ON T0.DocEntry = T1.DocEntry " & _
                            '                                    "INNER JOIN OWOR T2 " & _
                            '                                    "ON T2.DocNum = T1.U_MIS_PdONum " & _
                            '                                    "INNER JOIN WOR1 T3 " & _
                            '                                    "ON T2.Docentry = T3.Docentry " & _
                            '                                    "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X' AND T0.U_MIS_ItemCode = T3.ItemCode AND T1.U_MIS_QtPlanPdoIssue = (T3.PlannedQty / T0.U_MIS_QtyInLembar) AND T0.DocNum = '" & DocNum & "' " & _
                            '                                    " AND NOT EXISTS( " & _
                            '                                    "SELECT T4.baseref FROM IGE1 T4 " & _
                            '                                    "WHERE T4.BaseRef = T1.U_MIS_PdONum " & _
                            '                                    "AND T3.ItemCode = T4.ItemCode) "
                            objRecSet.DoQuery(StrSql)




                            If objRecSet.RecordCount > 0 Then
                                objRecSet.MoveFirst()
                                For Row = 1 To objRecSet.RecordCount
                                    If objRecSet.Fields.Item("Qty").Value = 0.0 Then
                                        SBO_Application.SetStatusBarMessage("In (Quantity) column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        objFormProductionIssue.Close()
                                        Exit Sub
                                    Else
                                        oColumns.Item("61").Cells.Item(oMatrix.VisualRowCount).Specific.string = objRecSet.Fields.Item("OrderNo").Value
                                        oColumns.Item("60").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("RowNo").Value
                                        oColumns.Item("9").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string = objRecSet.Fields.Item("Qty").Value
                                    End If
                                    objRecSet.MoveNext()
                                Next
                            Else
                                SBO_Application.SetStatusBarMessage("Please Check ItemCode, Planned Qty Production Order Not Same With Optimization Or Production Order Status Not Release", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'End If
                            'Next

                            'objFormProductionIssue.Close()
                            objFormProductionIssue.Items.Item("BtnCopy").Enabled = False
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objFormProductionIssue)
                            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumns)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecSet)
                        End If


                        'If pVal.ItemUID = "BtnCopy" Then
                        '    CopyOptimize(objFormProductionIssue)

                        '    Dim oForm As SAPbouiCOM.Form = Nothing
                        '    oForm = SBO_Application.Forms.Item("ListOptimize")

                        '    Dim ListOptimizeQuery As String
                        '    Dim oListOptimGrid As SAPbouiCOM.Grid = Nothing
                        '    Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

                        '    Dim oMis_Utils As MIS_Utils

                        '    oMis_Utils = New MIS_Utils

                        '    oListOptimGrid = oForm.Items.Item("myGrid2").Specific

                        '    'ListOptimizeQuery = "select T0.DocNum, T2.Visorder + 1 RowNo, T2.ItemCode,  T0.U_MIS_OPTNUM,  T0.U_MIS_ITEMCODE, T0.U_MIS_ITEMDESC, T0.U_MIS_PCM, T0.U_MIS_LCM, T0.U_MIS_LUASM2 " & _
                        '    '                    "From [@MIS_OPTIM] T0 INNER JOIN OWOR T1 ON T0.U_MIS_OPTNUM = T1.Docnum INNER JOIN WOR1 T2 ON T1.Docentry = T2.Docentry " & _
                        '    '                    "where T1.STATUS = 'R' AND LEFT(T2.ItemCode,1) <> 'X'"

                        '    ListOptimizeQuery = "select T0.DocNum [Number Optimize], T2.DocNum [PDO Number], T3.linenum + 1 [PDO Row No],  " & _
                        '                        "T1.U_MIS_QtPlanPdoIssue [Quantity Plan Optim], T0.U_MIS_QtyinLembar Lembar, T0.U_MIS_ITEMCODE [Item Code Line Optim], T0.U_MIS_ITEMDESC [Description Line Optim], " & _
                        '                        "T0.U_MIS_PCM [Panjang], T0.U_MIS_LCM [Lebar], T0.U_MIS_LUASM2 [Luas M2]" & _
                        '                        "From [@MIS_OPTIM] T0 " & _
                        '                        "join [@MIS_OPTIML] T1 " & _
                        '                        "ON T0.DocEntry = T1.DocEntry " & _
                        '                        "INNER JOIN OWOR T2 " & _
                        '                        "ON T2.DocNum = T1.U_MIS_PdONum " & _
                        '                        "INNER JOIN WOR1 T3 " & _
                        '                        "ON T2.Docentry = T3.Docentry " & _
                        '                        "where T2.STATUS = 'R' AND LEFT(T3.ItemCode,1) <> 'X'"

                        '    oForm.DataSources.DataTables.Item("ListOptim").ExecuteQuery(ListOptimizeQuery)
                        '    oListOptimGrid.DataTable = oForm.DataSources.DataTables.Item("ListOptim")

                        'End If
                End Select
            End If


            Select Case FormUID
                ' karno Prodution Status
                Case "PDOStatus"
                    If pVal.ItemUID = "BtnRelease" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oPDOStatusGrid As SAPbouiCOM.Grid

                        Dim dt As SAPbouiCOM.DataTable

                        dt = oForm.DataSources.DataTables.Item("PDOStatusLst")

                        oPDOStatusGrid = oForm.Items.Item("myGridPDO").Specific


                        GeneratePdOStatus(oForm)

                    End If

                    If ((pVal.ItemUID = "BtnShow") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)) Then

                        Dim oForm As SAPbouiCOM.Form = Nothing
                        oForm = SBO_Application.Forms.Item(FormUID)

                        Dim PDOStatusQuery As String
                        Dim oPDOStatusGrid As SAPbouiCOM.Grid = Nothing
                        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

                        Dim oMis_Utils As MIS_Utils

                        oMis_Utils = New MIS_Utils

                        oPDOStatusGrid = oForm.Items.Item("myGridPDO").Specific



                        'DelOutQuery = "select T0.DocEntry DocEntry, T0.DocNum SoDocNum, T0.DocDate Sodate, '' SalesRep, T0.CardCode SoCustCode, " & _
                        ' "T0.CardName SOCustName, T0.TrnspCode ShippingType, T1.TrnspName ShippingName, T0.DocStatus SoStatus " & _
                        ' "from ordr T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode " '& _
                        ''" Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("TxtDtFrm").Specific.string), "yyyyMMdd") & "' " & _
                        ''" AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("TxtDtTo").Specific.string), "yyyyMMdd") & "' "

                        If oForm.Items.Item("TxtDtFrm").Specific.string = "" Or oForm.Items.Item("TxtDtTo").Specific.string = "" Then
                            PDOStatusQuery = "SELECT 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                                            "T2.PlannedQty PdoQty, T2.CmpltQty PdoReceiptQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                                            "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                                            "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                                            "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                                            "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                                            "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                                            "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O'  " & _
                                            " AND T2.Status = 'P' " & _
                                            "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "
                        Else
                            PDOStatusQuery = "SELECT '1' gbr, 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                                           "T2.PlannedQty PdoQty, T2.CmpltQty PdoReceiptQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                                           "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                                           "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                                           "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                                           "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                                           "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                                           "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O' " & _
                                           " AND T2.Status = 'P' " & _
                                           " AND T2.PostDate >= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                                           " AND T2.PostDate <= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' " & _
                                           "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "

                            '    PDOStatusQuery = "SELECT 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                            '                    "T2.PlannedQty PdoQty, T1.Quantity SOQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                            '                    "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                            '                    "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                            '                    "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                            '                    "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                            '                    "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                            '                    "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O'  " & _
                            '                    "AND (T1.U_MIS_ReleasePdOFlag IS NULL OR T1.U_MIS_ReleasePdOFlag = '') AND T1.U_MIS_SupplyWith IS NULL " & _
                            '                    "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "
                            'Else
                            '    PDOStatusQuery = "SELECT 'N' [Release PdO], T2.PostDate PdoDate, T2.DocEntry PdoEntry, T2.DocNum Pdo#,T0.Docentry SOEntry, T2.OriginNum SoNumber, T0.DocDate SoDate, T2.ItemCode FGItemCode, " & _
                            '                   "T2.PlannedQty PdoQty, T1.Quantity SOQty, T2.Uom UM, T2.DueDate ExpDelDate, DATEDIFF(day, T2.DueDate, GETDATE()) Delayed, T2.U_MIS_Progress Progress, " & _
                            '                   "T3.CardCode CustomerCode, T3.CardName Customer, T1.VisOrder + 1 SoLine, T4.SlpName SalesRep " & _
                            '                   "FROM ORDR T0 INNER JOIN RDR1 T1 " & _
                            '                   "ON T0.DocEntry = T1.DocEntry LEFT JOIN OWOR T2 " & _
                            '                   "ON T0.DocNum = T2.OriginNum AND T2.itemcode = T1.ItemCode Inner Join OCRD T3 " & _
                            '                   "ON T2.CardCode = T3.CardCode INNER JOIN OSLP T4 " & _
                            '                   "ON T1.SlpCode = T4.SlpCode WHERE T0.DocStatus = 'O' " & _
                            '                   "AND (T1.U_MIS_ReleasePdOFlag IS NULL OR T1.U_MIS_ReleasePdOFlag = '') AND T1.U_MIS_SupplyWith IS NULL " & _
                            '                   " AND T2.PostDate >= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                            '                   " AND T2.PostDate <= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' " & _
                            '                   "ORDER BY T0.DocDate, T2.DueDate, T2.U_MIS_Progress, T2.OriginNum "

                        End If

                        ' Grid #: 1
                        'oForm.DataSources.DataTables.Add("DelOutLst")
                        oForm.DataSources.DataTables.Item("PDOStatusLst").ExecuteQuery(PDOStatusQuery)
                        oPDOStatusGrid.DataTable = oForm.DataSources.DataTables.Item("PDOStatusLst")

                        oPDOStatusGrid.Columns.Item("gbr").Type = SAPbouiCOM.BoGridColumnType.gct_Picture


                        oPDOStatusGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                        oPDOStatusGrid.Columns.Item("Release PdO").TitleObject.Sortable = True
                        'oPDOStatusGrid.Columns.Item("Release PdO").BackColor = 7

                        oColumn = oPDOStatusGrid.Columns.Item("PdoDate")
                        oPDOStatusGrid.Columns.Item("PdoDate").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("PdoEntry")
                        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_ProductionOrder
                        oPDOStatusGrid.Columns.Item("PdoEntry").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("Pdo#")
                        oPDOStatusGrid.Columns.Item("Pdo#").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("SOEntry")
                        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
                        oPDOStatusGrid.Columns.Item("SOEntry").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("SoNumber")
                        oPDOStatusGrid.Columns.Item("SoNumber").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("SoDate")
                        oPDOStatusGrid.Columns.Item("SoDate").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("FGItemCode")
                        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
                        oPDOStatusGrid.Columns.Item("FGItemCode").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("PdoQty")
                        oPDOStatusGrid.Columns.Item("PdoQty").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("PdoReceiptQty")
                        oPDOStatusGrid.Columns.Item("PdoReceiptQty").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("UM")
                        oPDOStatusGrid.Columns.Item("UM").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("ExpDelDate")
                        oPDOStatusGrid.Columns.Item("ExpDelDate").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("Delayed")
                        'oPDOStatusGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                        oPDOStatusGrid.Columns.Item("Delayed").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("Progress")
                        oPDOStatusGrid.Columns.Item("Progress").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("CustomerCode")
                        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
                        oPDOStatusGrid.Columns.Item("CustomerCode").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("Customer")
                        oPDOStatusGrid.Columns.Item("Customer").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("SoLine")
                        oPDOStatusGrid.Columns.Item("SoLine").TitleObject.Sortable = True
                        oColumn.Editable = False

                        oColumn = oPDOStatusGrid.Columns.Item("SalesRep")
                        oPDOStatusGrid.Columns.Item("SalesRep").TitleObject.Sortable = True
                        oColumn.Editable = False

                        'Dim idx As Integer
                        'Dim oColumns As SAPbouiCOM.Columns
                        ''Dim oColumn As SAPbouiCOM.Column

                        'oForm.DataSources.UserDataSources.Add("gbr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                        'oForm.DataSources.UserDataSources.Item("gbr").ValueEx = "E:\karno\Maruni_Production_100\Maruni_Production_100\CFL.bmp"

                        'For idx = oPDOStatusGrid.Rows.Count - 1 To 0 Step -1
                        '    If oPDOStatusGrid.DataTable.GetValue(13, oPDOStatusGrid.GetDataTableRowIndex(idx)) > 0 Then
                        '        oPDOStatusGrid.Columns.Item("Release PdO").BackColor = 2
                        '        oColumns = oPDOStatusGrid.Columns
                        '        oColumn = oColumns.Item("gbr")
                        '        oColumn.DataBind.SetBound(True, , "gbr")
                        '    Else
                        '        oPDOStatusGrid.Columns.Item("Release PdO").Type = SAPbouiCOM.BoGridColumnType.gct_Picture
                        '        oPDOStatusGrid.Columns.Item("Release PdO").BackColor = 2
                        '    End If
                        'Next

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPDOStatusGrid)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

                    End If
                    ' Karno Out Del
                Case "OutDel"
                    If ((pVal.ItemUID = "BtnShow") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)) Then

                        Dim oForm As SAPbouiCOM.Form = Nothing
                        oForm = SBO_Application.Forms.Item(FormUID)

                        Dim DelOutQuery As String
                        Dim oDelOutGrid As SAPbouiCOM.Grid = Nothing
                        Dim oColumn As SAPbouiCOM.EditTextColumn = Nothing

                        Dim oMis_Utils As MIS_Utils

                        oMis_Utils = New MIS_Utils

                        oDelOutGrid = oForm.Items.Item("myGrid1").Specific

                        'DelOutQuery = "select T0.DocEntry DocEntry, T0.DocNum SoDocNum, T0.DocDate Sodate, '' SalesRep, T0.CardCode SoCustCode, " & _
                        ' "T0.CardName SOCustName, T0.TrnspCode ShippingType, T1.TrnspName ShippingName, T0.DocStatus SoStatus " & _
                        ' "from ordr T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode " '& _
                        ''" Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("TxtDtFrm").Specific.string), "yyyyMMdd") & "' " & _
                        ''" AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("TxtDtTo").Specific.string), "yyyyMMdd") & "' "

                        If oForm.Items.Item("TxtDtFrm").Specific.string = "" Or oForm.Items.Item("TxtDtTo").Specific.string = "" Then

                            DelOutQuery = "select DISTINCT T0.DocEntry DocEntry, T0.DocNum So_DocNum, T0.DocDate So_Date, " & _
                            "T4.SlpName Sales_Rep, T0.CardCode Customer_Code, T0.CardName Customer_Name, " & _
                            "T1.TrnspName Shipping_Type, CASE WHEN(select COUNT(P1.itemcode) from ORDR P0 " & _
                            "INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry WHERE P0.DocEntry = T0.DocEntry " & _
                            ") > (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry " & _
                            "LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode WHERE P0.DocEntry = T0.Docentry " & _
                            "AND P2.OnHand > 0) THEN 'Partialy Ready' " & _
                            "WHEN (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.Docentry " & _
                            "WHERE P0.DocEntry = T0.Docentry) = (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1  " & _
                            "ON P0.DocEntry = P1.DocEntry LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode " & _
                            "WHERE P0.DocEntry = T0.Docentry AND P2.OnHand > 0) THEN 'Completely Ready' END Status " & _
                            "From ORDR T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode  LEFT JOIN OSLP T4 ON T0.SlpCode = T4.SlpCode INNER JOIN RDR1 T2 " & _
                            "ON T0.DocEntry = T2.DocEntry LEFT JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " & _
                            "AND T2.WhsCode = T3.WhsCode WHERE T3.OnHand >= T2.Quantity AND T0.DocStatus = 'O' ORDER BY T0.DocDate "
                        Else
                            DelOutQuery = "select DISTINCT T0.DocEntry DocEntry, T0.DocNum So_DocNum, T0.DocDate So_Date, " & _
                            "T4.SlpName Sales_Rep, T0.CardCode Customer_Code, T0.CardName Customer_Name, " & _
                            "T1.TrnspName Shipping_Type, CASE WHEN(select COUNT(P1.itemcode) from ORDR P0 " & _
                            "INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry WHERE P0.DocEntry = T0.DocEntry " & _
                            ") > (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.DocEntry " & _
                            "LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode WHERE P0.DocEntry = T0.Docentry " & _
                            "AND P2.OnHand > 0) THEN 'Partialy Ready' " & _
                            "WHEN (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1 ON P0.DocEntry = P1.Docentry " & _
                            "WHERE P0.DocEntry = T0.Docentry) = (select COUNT(P1.itemcode) from ORDR P0 INNER JOIN RDR1 P1  " & _
                            "ON P0.DocEntry = P1.DocEntry LEFT JOIN OITW P2 ON P1.ItemCode = P2.ItemCode AND P1.WhsCode = P2.WhsCode " & _
                            "WHERE P0.DocEntry = T0.Docentry AND P2.OnHand > 0) THEN 'Completely Ready' END Status " & _
                            "From ORDR T0 LEFT JOIN OSHP T1 ON T0.TrnspCode = T1.TrnspCode  LEFT JOIN OSLP T4 ON T0.SlpCode = T4.SlpCode INNER JOIN RDR1 T2 " & _
                            "ON T0.DocEntry = T2.DocEntry LEFT JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " & _
                            "AND T2.WhsCode = T3.WhsCode WHERE T3.OnHand >= T2.Quantity AND T0.DocStatus = 'O' " & _
                            " AND T0.DocDate >= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtFrm").Specific.string, oCompany) & "' " & _
                            " AND T0.DocDate <= '" & oMis_Utils.fctFormatDate(oForm.Items.Item("TxtDtTo").Specific.string, oCompany) & "' ORDER BY T0.DocDate "

                        End If


                        ' Grid #: 1
                        'oForm.DataSources.DataTables.Add("DelOutLst")
                        oForm.DataSources.DataTables.Item("DelOutLst").ExecuteQuery(DelOutQuery)
                        oDelOutGrid.DataTable = oForm.DataSources.DataTables.Item("DelOutLst")

                        oColumn = oDelOutGrid.Columns.Item("DocEntry")
                        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("So_DocNum")
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("So_Date")
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("Sales_Rep")
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("Customer_Code")
                        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("Customer_Name")
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("Shipping_Type")
                        oColumn.Editable = False

                        oColumn = oDelOutGrid.Columns.Item("Status")
                        oColumn.Editable = False

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDelOutGrid)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn)

                    End If


                Case "mds_p1"
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal

                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID

                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)

                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)



                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects

                            Dim xval As String


                            xval = oDataTable.GetValue(0, 0)

                            If pVal.ItemUID = "BPCardCode" Or pVal.ItemUID = "BPButton" Then

                                oForm.DataSources.UserDataSources.Item("BPDS").ValueEx = xval
                            End If

                            oCFL = Nothing
                            oDataTable = Nothing
                        End If

                        'oForm = Nothing
                        'oCFLEvento = Nothing
                        'GC.Collect()

                    End If


                    ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                    ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadSO") Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)

                        'If Not SOToMFGFormValid(oForm) Then
                        '    SBO_Application.SetStatusBarMessage("Form invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '    BubbleEvent = False

                        'Else
                        '    LoadSO(oForm)
                        'End If

                        LoadSO(oForm)


                    End If

                    If (pVal.ItemUID = "SODateFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)

                        'Dim vdate As MISToolbox
                        'vdate = New MISToolbox
                        'Dim validDate As Boolean


                        If Len(oForm.Items.Item("SODateFrom").Specific.string) = 0 Then
                            SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Exit Sub
                        End If

                        'validDate = vdate.SBODateisValid("2010918")

                        'validDate = vdate.SBODateisValid(oForm.Items.Item("SODateFrom").Specific.string)
                        'If validDate = False Then
                        '    SBO_Application.SetStatusBarMessage("SO Date From is invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        If Len(oForm.Items.Item("SODateFrom").Specific.string) < 8 Then
                            SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Exit Sub
                        End If

                        If Len(oForm.Items.Item("SODateFrom").Specific.string) = 8 Then
                            oForm.Items.Item("SODateFrom").Specific.string = _
                                CDate(Left(oForm.Items.Item("SODateFrom").Specific.string, 4) & "/" & _
                                    Mid(oForm.Items.Item("SODateFrom").Specific.string, 5, 2) & "/" & _
                                    Right(oForm.Items.Item("SODateFrom").Specific.string, 2))
                        End If

                        If oForm.Items.Item("SODateFrom").Specific.string = "" Then
                            oForm.Items.Item("SODateFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                        End If

                        If oForm.Items.Item("SODateTo").Specific.string = "" Then
                            oForm.Items.Item("SODateTo").Specific.string = oForm.Items.Item("SODateFrom").Specific.string
                        End If

                        'vdate = Nothing

                        'oForm.Items.Item("SODateFrom").Click()
                        '                        BubbleEvent = False
                    End If

                    If pVal.ItemUID = "SODateTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)

                        If Len(oForm.Items.Item("SODateTo").Specific.string) = 0 Then
                            SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            'Exit Sub
                        End If

                        If Len(oForm.Items.Item("SODateTo").Specific.string) < 8 Then
                            SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                        End If

                        If Len(oForm.Items.Item("SODateTo").Specific.string) = 8 Then
                            oForm.Items.Item("SODateTo").Specific.string = _
                                CDate(Left(oForm.Items.Item("SODateTo").Specific.string, 4) & "/" & _
                                    Mid(oForm.Items.Item("SODateTo").Specific.string, 5, 2) & "/" & _
                                    Right(oForm.Items.Item("SODateTo").Specific.string, 2))
                        End If
                        'BubbleEvent = True
                        'oForm.Items.Item("SODateTo").Click('')

                        'oForm = Nothing
                        'GC.Collect()

                    End If

                    If pVal.ItemUID = "cmdGenPdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oSOToMFGGrid As SAPbouiCOM.Grid

                        Dim dt As SAPbouiCOM.DataTable

                        dt = oForm.DataSources.DataTables.Item("SOToMFGLst")

                        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

                        'get total row count selected
                        'oSOToMFGGrid.Rows.SelectedRows.Count.ToString()


                        'selection rows -> e.g: user select row# by order respectively: 1, 3, 2, 5

                        'get row index of selected grid, has two method:
                        'method# 1: ot_RowOrder (value=1)
                        'result row selected: 1, 2, 3, 5

                        'method# 2: ot_SelectionOrder (value=0)
                        'result row selected: 1, 3, 2, 5

                        'For idx = 0 To oSOToMFGGrid.Rows.SelectedRows.Count - 1
                        '    MsgBox("selected row#:" & idx.ToString & _
                        '           "; selectedrow->row#: " & oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
                        '           & "docnum: " & oSOToMFGGrid.DataTable.GetValue(0, oSOToMFGGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

                        'Next

                        'Dim oPdO As SAPbobsCOM.ProductionOrders
                        'oPdO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                        ''Fill PdO properties...
                        'oPdO.ItemNo = "LM4029"
                        ''oPdO.DueDate = oSOToMFGGrid.DataTable.GetValue(13, oSOToMFGGrid.GetDataTableRowIndex(idx).ToString)
                        'oPdO.DueDate = DateTime.Today.ToString("yyyyMMdd")
                        'oPdO.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                        'oPdO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                        'oPdO.PlannedQuantity = 188
                        'oPdO.PostingDate = DateTime.Today 'DateTime.Today.ToString("yyyyMMdd")
                        'oPdO.Add()

                        GeneratePdOFromSO(oForm)

                        LoadSO(oForm)


                    End If

                    'toggle select/unselect all
                    If pVal.ColUID = "Release PdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oSOToMFGGrid As SAPbouiCOM.Grid

                        Dim idx As Long
                        Dim dt As SAPbouiCOM.DataTable

                        dt = oForm.DataSources.DataTables.Item("SOToMFGLst")

                        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

                        'get total row count selected
                        'oSOToMFGGrid.Rows.SelectedRows.Count.ToString()


                        oSOToMFGGrid = oForm.Items.Item("myGrid").Specific

                        If oSOToMFGGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                            'select/check all
                            oForm.Freeze(True)

                            For idx = 0 To oSOToMFGGrid.Rows.Count - 1
                                dt.SetValue("Release PdO", idx, "Y")
                            Next
                            oSOToMFGGrid.Columns.Item(1).TitleObject.Caption = "Reset All"
                            oForm.Freeze(False)
                        Else
                            'unselect/uncheck all
                            oForm.Freeze(True)
                            For idx = 0 To oSOToMFGGrid.Rows.Count - 1
                                dt.SetValue("Release PdO", idx, "N")
                            Next
                            oSOToMFGGrid.Columns.Item(1).TitleObject.Caption = "Select All"
                            oForm.Freeze(False)
                        End If

                        'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                    End If

                Case "mds_p3"

                    If pVal.ItemUID = "SeriesOptm" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then ' 3 = fm_ADD_MODE 
                        Dim lNextSeriesNumOptimization As Long
                        Dim Series As String
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)

                        Dim cmbSeries As SAPbouiCOM.ComboBox
                        cmbSeries = oForm.Items.Item("SeriesOptm").Specific
                        Series = cmbSeries.Selected.Value
                        lNextSeriesNumOptimization = oForm.BusinessObject.GetNextSerialNumber(Series)

                        Dim oItem As SAPbouiCOM.EditText
                        oItem = oForm.Items.Item("DocNum").Specific
                        oItem.Value = lNextSeriesNumOptimization

                        oItem = oForm.Items.Item("ByUser").Specific
                        oItem.Value = oCompany.UserName.ToString

                        oForm.Items.Item("KcSisaPctg").Specific.value = 0
                        oForm.Items.Item("TotWastPct").Specific.value = 0

                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = oForm.Items.Item("OptimMtx").Specific
                        oMatrix.AddRow()
                        oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount
                        oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount

                        ''????
                        'Dim txtColor As SAPbouiCOM.EditText
                        'txtColor = oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific
                        'txtColor.BackColor = 3

                        oForm.Items.Item("GTabc").Specific.value = 0
                        oForm.Items.Item("GTaloc").Specific.value = 0
                        oForm.Items.Item("GTplanPdO").Specific.value = 0

                        oForm.Items.Item("QtyLembar").Specific.value = 1
                        oForm.Items.Item("OptimDate").Specific.value = DateTime.Today.ToString("yyyyMMdd")

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(cmbSeries)


                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal

                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID

                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)

                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)

                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable = Nothing
                            oDataTable = oCFLEvento.SelectedObjects

                            If Not oDataTable Is Nothing Then
                                Dim xVal As String
                                xVal = oDataTable.GetValue(0, 0)

                                Dim oDBDataSource As SAPbouiCOM.DBDataSource
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                                If pVal.ItemUID = "ItemCode" Or pVal.ItemUID = "ItemButton" Then
                                    oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_ItemCode", oDBDataSource.Offset, xVal)
                                    oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_ItemDesc", oDBDataSource.Offset, oDataTable.GetValue(1, 0))
                                    oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_Pcm", oDBDataSource.Offset, oDataTable.GetValue("SHeight1", 0))
                                    oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_Lcm", oDBDataSource.Offset, oDataTable.GetValue("SWidth1", 0))

                                    Dim oRecLengthWidth As SAPbobsCOM.Recordset = Nothing
                                    Dim StrQuery As String

                                    oRecLengthWidth = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    StrQuery = "SELECT UnitCode, SizeInMM FROM OITM T0 JOIN OLGT T1 ON T0.SHght1Unit = T1.UnitCode WHERE ItemCode = '" & xVal & "' "

                                    oRecLengthWidth.DoQuery(StrQuery)

                                    Dim inMM As Double
                                    Dim LuasInM2 As Double

                                    inMM = oRecLengthWidth.Fields.Item("SizeInMM").Value
                                    LuasInM2 = (oDataTable.GetValue("SHeight1", 0) * inMM) * (oDataTable.GetValue("SWidth1", 0) * inMM) / 1000000
                                    oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_LuasM2", oDBDataSource.Offset, LuasInM2)

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecLengthWidth)
                                    oRecLengthWidth = Nothing

                                End If
                                If pVal.ItemUID = "ItemKcSisa" Or pVal.ItemUID = "ItmSisaBtn" Then
                                    oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_ItemCdKacaSisa", oDBDataSource.Offset, xVal)
                                End If

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable)
                                GC.Collect()

                            End If

                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLEvento)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)

                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then

                        Dim oForm As SAPbouiCOM.Form = Nothing
                        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                        oForm = SBO_Application.Forms.Item(FormUID)

                        'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then
                            oMatrix = oForm.Items.Item("OptimMtx").Specific

                            oForm.Freeze(True)


                            Dim idx As Long
                            Dim gtabc As Double
                            Dim gtaloc As Double
                            Dim gtplanpdo As Double

                            gtabc = 0
                            gtaloc = 0
                            gtplanpdo = 0
                            For idx = 1 To oMatrix.RowCount
                                gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                                'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                                gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))
                                'oForm.Items.Item("#").Specific.value = idx
                                'oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.value = oMatrix.VisualRowCount
                                oMatrix.Columns.Item("#").Cells.Item(CInt(idx)).Specific.value = idx


                            Next

                            Dim oDBDataSource As SAPbouiCOM.DBDataSource = Nothing
                            oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                            Dim i As Integer
                            gtabc = 0
                            Dim oDBDataSource_OptimL As SAPbouiCOM.DBDataSource = Nothing
                            oDBDataSource_OptimL = oForm.DataSources.DBDataSources.Item("@MIS_OPTIML")
                            For i = 0 To oDBDataSource_OptimL.Size - 1
                                gtabc += oDBDataSource_OptimL.GetValue("U_MIS_TotalABC", i)
                            Next

                            Dim docnum As Integer
                            Dim LuasKaca As Double
                            Dim SisaKacaUtuh As Double
                            Dim TotalWaste As Double

                            docnum = oDBDataSource.GetValue("docnum", 0)
                            LuasKaca = oDBDataSource.GetValue("U_MIS_LuasM2", 0)
                            SisaKacaUtuh = oDBDataSource.GetValue("U_MIS_KcSisaUtuh", 0)
                            TotalWaste = oDBDataSource.GetValue("U_MIS_TotalWaste", 0)

                            ' by Toin 2011-02-10
                            oForm.Items.Item("TotWastPct").Specific.value = IIf(LuasKaca = 0, 0, Math.Round(TotalWaste / LuasKaca * 100, 2))
                            ' by Toin 2011-02-10
                            ' by Toin 2011-03-01
                            oForm.Items.Item("KcSisaPctg").Specific.value = IIf(LuasKaca = 0, 0, Math.Round(SisaKacaUtuh / LuasKaca * 100, 2))
                            ' by Toin 2011-03-01

                            gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                            oForm.Items.Item("GTabc").Specific.value = gtabc
                            'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                            oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo



                            ' by Toin 2011-02-28 = Grand Total Waste = Total Waste + Sisa Kaca Utuh
                            'LuasKaca(-SisaKacaUtuh - gtabc)
                            oForm.Items.Item("GTaloc").Specific.value = _
                            TotalWaste + SisaKacaUtuh
                            'TotalWaste

                            BubbleEvent = False

                            oForm.Freeze(False)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDBDataSource)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDBDataSource_OptimL)

                            oMatrix = Nothing
                            oDBDataSource = Nothing
                            oDBDataSource_OptimL = Nothing
                            oForm = Nothing

                            GC.Collect()


                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD _
                             Then

                        End If

                    End If

                    'karno haha
                    'If pVal.ItemUID = "1" Then
                    '    Dim oForm As SAPbouiCOM.Form

                    '    oForm = SBO_Application.Forms.Item(FormUID)

                    '    'Validation - Item Code Kaca RM harus idem Item Code Kaca Sisa, 1st 7 Digit Item code harus sama!
                    '    If Left(oForm.Items.Item("ItemCode").Specific.value, 7) <> Left(oForm.Items.Item("ItemKcSisa").Specific.value, 7) Then
                    '        SBO_Application.SetStatusBarMessage("Item Code kaca sisa harus idem dgn Item Code Kaca RM!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    '    End If

                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                    '    'oDBDatasource = Nothing

                    'End If


                    'If pVal.ItemUID = "OptimMtx" And _
                    '    pVal.ColUID = "AlocWaste" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then

                    '    Dim oForm As SAPbouiCOM.Form = Nothing
                    '    Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                    '    oForm = SBO_Application.Forms.Item(FormUID)
                    '    oMatrix = oForm.Items.Item("OptimMtx").Specific

                    '    oForm.Freeze(True)

                    '    'Plan PdO Issue = Total AxBxC + Allocated Waste
                    '    oMatrix.Columns.Item("PlanPdIsue").Cells.Item(pVal.Row).Specific.value = _
                    '    IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value)) + _
                    '    IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value))


                    '    Dim idx As Long
                    '    Dim gtabc As Double
                    '    Dim gtaloc As Double
                    '    Dim gtplanpdo As Double

                    '    gtabc = 0
                    '    gtaloc = 0
                    '    gtplanpdo = 0

                    '    If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then

                    '        For idx = 1 To oMatrix.RowCount
                    '            gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                    '            'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                    '            gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))
                    '        Next
                    '    End If

                    '    gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value


                    '    oForm.Items.Item("GTabc").Specific.value = gtabc
                    '    'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                    '    oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo
                    '    oForm.Items.Item("GTaloc").Specific.value = _
                    '    IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)

                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                    '    oMatrix = Nothing

                    '    oForm.Freeze(False)
                    'End If

                    If (pVal.ItemUID = "LebarKaca" Or pVal.ItemUID = "PnjangKaca") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                        Dim oForm As SAPbouiCOM.Form = Nothing
                        Dim oItem As SAPbouiCOM.Item
                        Dim oMatrix As SAPbouiCOM.Matrix
                        Dim oEditText As SAPbouiCOM.EditText

                        'Dim sb As String

                        oForm = SBO_Application.Forms.Item(FormUID)
                        oItem = oForm.Items.Item("OptimMtx")
                        oMatrix = oItem.Specific
                        'oEditText = oMatrix.GetCellSpecific(5, 1)
                        'sb = oEditText.Value

                        Dim oRecLengthWidth As SAPbobsCOM.Recordset = Nothing
                        Dim StrQuery As String

                        oRecLengthWidth = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        StrQuery = "SELECT UnitCode, SizeInMM FROM OITM T0 JOIN OLGT T1 ON T0.SHght1Unit = T1.UnitCode WHERE ItemCode = '" & oForm.Items.Item("ItemCode").Specific.value & "' "

                        oRecLengthWidth.DoQuery(StrQuery)

                        Dim inMM As Double
                        'Dim LuasInM2 As Double

                        inMM = oRecLengthWidth.Fields.Item("SizeInMM").Value
                        'LuasInM2 = oDataTable.GetValue("SHeight1", 0) * inMM * oDataTable.GetValue("SWidth1", 0) / 1000
                        'oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_LuasM2", oDBDataSource.Offset, LuasInM2)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecLengthWidth)
                        oRecLengthWidth = Nothing
                        GC.Collect()

                        oEditText = oForm.Items.Item("LuasKaca").Specific
                        oEditText.Value = _
                            IIf(oForm.Items.Item("PnjangKaca").Specific.value = "", 0, oForm.Items.Item("PnjangKaca").Specific.value) _
                            * inMM * _
                            IIf(oForm.Items.Item("LebarKaca").Specific.value = "", 0, oForm.Items.Item("LebarKaca").Specific.value) _
                            * inMM _
                            / 1000000

                        'oEditText = oForm.Items.Item("KacaPakai").Specific
                        'oEditText.Value = oForm.Items.Item("LuasKaca").Specific.value - oForm.Items.Item("SisaKcUtuh").Specific.value

                        'oEditText = oForm.Items.Item("GTaloc").Specific
                        'oEditText.Value = _
                        '    IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) _
                        '    - IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) _
                        '    - IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)

                        '2011-02-28
                        oEditText = oForm.Items.Item("GTaloc").Specific
                        oEditText.Value = _
                            CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))
                        'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)

                        'oEditText = oForm.Items.Item("TotalABC").Specific

                        oForm.Items.Item("KacaPakai").Specific.value = _
                        CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                        - CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) _
                        - CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        GC.Collect()

                    End If

                    If pVal.ItemUID = "KcSisaPctg" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                        Dim oForm As SAPbouiCOM.Form = Nothing
                        Dim oEditText As SAPbouiCOM.EditText

                        oForm = SBO_Application.Forms.Item(FormUID)
                        oEditText = oForm.Items.Item("SisaKcUtuh").Specific

                        If oForm.Items.Item("SisaKcUtuh").Specific.value = "" Then
                            oEditText.Value = (IIf(oForm.Items.Item("KcSisaPctg").Specific.value = "", 0, oForm.Items.Item("KcSisaPctg").Specific.value) * IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) / 100
                        End If

                        oForm.Items.Item("KacaPakai").Specific.value = _
                        IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) - _
                        IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) - _
                        IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
                        GC.Collect()

                    End If

                    If pVal.ItemUID = "SisaKcUtuh" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                        Dim oForm As SAPbouiCOM.Form = Nothing
                        oForm = SBO_Application.Forms.Item(FormUID)

                        oForm.Items.Item("KacaPakai").Specific.value = _
                        CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                        - CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) _
                        - CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        'GC.Collect()

                        Dim TotalWaste As Double
                        Dim Kolom As Double

                        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                        oForm.Items.Item("TotalWaste").Specific.Value = (CDbl(IIf(oForm.Items.Item("TotWastPct").Specific.value = "", 0, oForm.Items.Item("TotWastPct").Specific.value)) * CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value))) / 100

                        ' 2011-02-28
                        TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                        Dim idx As Long
                        Dim gtabc As Double
                        Dim gtplanpdo As Double

                        'oForm = SBO_Application.Forms.Item(FormUID)
                        oMatrix = oForm.Items.Item("OptimMtx").Specific

                        oForm.Freeze(True)

                        gtabc = 0

                        'gtplanpdo = 0
                        'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                        For idx = 1 To oMatrix.RowCount
                            gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                            'gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                        Next
                        'End If

                        'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And oMatrix.RowCount = 0 Then
                        '    oForm.Items.Item("GTabc").Specific.value = 0
                        'Else
                        oForm.Items.Item("GTabc").Specific.value = gtabc
                        'End If

                        'oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                        ' 2011-02-28
                        oForm.Items.Item("GTaloc").Specific.value = _
                        CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                        + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                        'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) _
                        '+ IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, CDbl(oForm.Items.Item("SisaKcUtuh").Specific.value))

                        'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value))

                        Kolom = oForm.Items.Item("GTabc").Specific.value

                        'oMatrix = oForm.Items.Item("OptimMtx").Specific

                        'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then

                        For Row = 1 To oMatrix.RowCount
                            'Allocated Waste
                            oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = _
                            Math.Round( _
                                (CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, _
                                oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                / Kolom) * TotalWaste, _
                            4)


                            'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                            oMatrix.Columns.Item("PlanPdIsue").Cells.Item(Row).Specific.value = _
                            Math.Round( _
                                CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) + _
                                CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value)) _
                            , 4)


                        Next
                        'End If

                        gtplanpdo = 0
                        'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                        For idx = 1 To oMatrix.RowCount
                            'gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                            gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                        Next
                        'End If

                        oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                        ' 2011-02-28
                        oForm.Items.Item("GTaloc").Specific.value = _
                        CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                        + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                        ' by Toin 2011-03-01
                        If CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) = 0 Or _
                            CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) = 0 Then
                            oForm.Items.Item("KcSisaPctg").Specific.value = 0
                        Else
                            oForm.Items.Item("KcSisaPctg").Specific.value = _
                                (IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, _
                                    Math.Round(CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, _
                                                   oForm.Items.Item("SisaKcUtuh").Specific.value)) / _
                                               CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, _
                                                   oForm.Items.Item("LuasKaca").Specific.value)) * 100, 2)))
                        End If
                        ' by Toin 2011-03-01

                        oForm.Items.Item("KacaPakai").Specific.value = _
                        CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) - _
                        CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) - _
                        CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))


                        oForm.Freeze(False)
                        'oForm.Refresh()

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        GC.Collect()

                    End If

                    'If pVal.ItemUID = "TotWastPct" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                    '    Dim oForm As SAPbouiCOM.Form = Nothing
                    '    Dim oEditText As SAPbouiCOM.EditText

                    '    oForm = SBO_Application.Forms.Item(FormUID)
                    '    oEditText = oForm.Items.Item("TotalWaste").Specific

                    '    

                    '    oForm.Items.Item("KacaPakai").Specific.value = _
                    '    IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) - _
                    '    IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) - _
                    '    IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)

                    'End If


                    If (pVal.ItemUID = "TotalWaste" Or pVal.ItemUID = "TotWastPct") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                        Dim TotalWaste As Double
                        Dim Kolom As Double
                        Dim oForm As SAPbouiCOM.Form = Nothing

                        Dim oMatrix As SAPbouiCOM.Matrix = Nothing

                        oForm = SBO_Application.Forms.Item(FormUID)

                        oForm.Items.Item("TotalWaste").Specific.value = _
                            Math.Round( _
                                CDbl(IIf(oForm.Items.Item("TotWastPct").Specific.value = "", 0.0, oForm.Items.Item("TotWastPct").Specific.value)) / 100 * _
                                CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                            , 4)


                        Dim totwastpct As Double
                        Dim luaskaca As Double
                        totwastpct = CDbl(IIf(oForm.Items.Item("TotWastPct").Specific.value = "", 0.0, oForm.Items.Item("TotWastPct").Specific.value))
                        luaskaca = CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value))

                        ' 2011-02-28
                        TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                        'oForm.Items.Item("GTaloc").Specific.value = TotalWaste

                        Dim idx As Long
                        Dim gtabc As Double
                        Dim gtplanpdo As Double

                        'oForm = SBO_Application.Forms.Item(FormUID)
                        oMatrix = oForm.Items.Item("OptimMtx").Specific

                        oForm.Freeze(True)

                        gtabc = 0

                        'gtplanpdo = 0
                        'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                        For idx = 1 To oMatrix.RowCount
                            gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                            'gtplanpdo += IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                        Next
                        'End If
                        'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And oMatrix.RowCount = 0 Then
                        '    oForm.Items.Item("GTabc").Specific.value = 0
                        'Else
                        oForm.Items.Item("GTabc").Specific.value = gtabc
                        'End If
                        'oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                        ' 2011-02-28
                        oForm.Items.Item("GTaloc").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))
                        'IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value))

                        Kolom = oForm.Items.Item("GTabc").Specific.value

                        'oMatrix = oForm.Items.Item("OptimMtx").Specific

                        'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then

                        For Row = 1 To oMatrix.RowCount
                            'Allocated Waste
                            oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = _
                            Math.Round( _
                                (CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, _
                                oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                / Kolom) * TotalWaste, _
                            4)


                            'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                            oMatrix.Columns.Item("PlanPdIsue").Cells.Item(Row).Specific.value = _
                            Math.Round( _
                                CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) + _
                                CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value)) _
                            , 4)


                        Next
                        'End If

                        gtplanpdo = 0
                        'If IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, CDbl(oForm.Items.Item("TotalWaste").Specific.value)) <> 0 Then
                        For idx = 1 To oMatrix.RowCount
                            'gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                            gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                        Next
                        'End If

                        oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                        ' 2011-02-28
                        oForm.Items.Item("GTaloc").Specific.value = _
                            CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))


                        oForm.Items.Item("KacaPakai").Specific.value = _
                        CDbl(IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value)) _
                        - CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)) _
                        - CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                        oForm.Freeze(False)
                        'oForm.Refresh()

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        GC.Collect()

                    End If

                    '???
                    If pVal.ItemUID = "OptimMtx" And _
                        pVal.ColUID = "PdO#" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then

                        Dim oForm As SAPbouiCOM.Form = Nothing

                        Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                        oForm = SBO_Application.Forms.Item(FormUID)

                        oMatrix = oForm.Items.Item("OptimMtx").Specific

                        oForm.Freeze(True)

                        'Total AxBxC = Jumlah Potong x P x L
                        'oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = _

                        If oMatrix.Columns.Item("PdO#").Cells.Item(pVal.Row).Specific.value <> "" Then
                            Dim oRecPdo As SAPbobsCOM.Recordset = Nothing
                            Dim StrQuery As String

                            oRecPdo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            StrQuery = "SELECT T0.DocNum, OriginNum, T2.LineNum, T0.CardCode, T1.CardName FROM OWOR T0 " & _
                                " LEFT JOIN ORDR T1 ON T1.DocNum = OriginNum " & _
                                " LEFT JOIN RDR1 T2 ON T2.DocEntry = T1.DocEntry " & _
                                " WHERE T0.DocNum = " & oMatrix.Columns.Item("PdO#").Cells.Item(pVal.Row).Specific.value
                            '" & oForm.Items.Item("ItemCode").Specific.value & "' "

                            oRecPdo.DoQuery(StrQuery)


                            oMatrix.Columns.Item("SO#").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("OriginNum").Value
                            oMatrix.Columns.Item("SOLine").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("LineNum").Value
                            oMatrix.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("CardCode").Value
                            oMatrix.Columns.Item("CardName").Cells.Item(pVal.Row).Specific.value = oRecPdo.Fields.Item("CardName").Value


                            'LuasInM2 = oDataTable.GetValue("SHeight1", 0) * inMM * oDataTable.GetValue("SWidth1", 0) / 1000
                            'oForm.DataSources.DBDataSources.Item("@MIS_OPTIM").SetValue("U_MIS_LuasM2", oDBDataSource.Offset, LuasInM2)

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecPdo)
                            oRecPdo = Nothing

                            'GC.WaitForPendingFinalizers()
                            GC.Collect()

                            'Dim oColumn As SAPbouiCOM.Column
                            'Dim oCell As SAPbouiCOM.Cell

                            'oColumn = oMatrix.Columns.Item("CardName")
                            'oCell = oColumn.Cells.Item(pVal.Row)
                            'oCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            'oMatrix.Columns.Item(4).Cells.Item(lstRowIndex + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            'oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            'BubbleEvent = False
                        End If

                        'oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)

                        'BubbleEvent = False
                        oForm.Freeze(False)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        GC.Collect()

                    End If

                    'If pVal.ItemUID = "OptimMtx" And _
                    '    pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_GOT_FOCUS Then
                    '    MsgBox("Got Focus! Qty Potong isi ya")
                    '    ''????
                    '    Dim oForm As SAPbouiCOM.Form = Nothing

                    '    Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                    '    oForm = SBO_Application.Forms.Item(FormUID)

                    '    oMatrix = oForm.Items.Item("OptimMtx").Specific
                    '    oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Click()
                    '    BubbleEvent = False

                    'End If


                    'If pVal.ItemUID = "OptimMtx" And _
                    'pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                    '    MsgBox("Lost Focus!! Qty Potong isi ya")
                    '    ''????
                    '    Dim oForm As SAPbouiCOM.Form = Nothing

                    '    Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                    '    oForm = SBO_Application.Forms.Item(FormUID)

                    '    oMatrix = oForm.Items.Item("OptimMtx").Specific

                    '    oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)


                    '    BubbleEvent = False
                    'End If

                    'If pVal.ItemUID = "OptimMtx" And _
                    'pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.Before_Action = False Then
                    '    MsgBox("Validate!! Qty Potong isi ya")
                    '    ''????
                    '    Dim oForm As SAPbouiCOM.Form = Nothing

                    '    Dim oMatrix As SAPbouiCOM.Matrix = Nothing


                    '    oForm = SBO_Application.Forms.Item(FormUID)

                    '    oMatrix = oForm.Items.Item("OptimMtx").Specific

                    '    oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Click()


                    '    BubbleEvent = False
                    'End If

                    If pVal.ItemUID = "OptimMtx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK Then
                        Dim oForm As SAPbouiCOM.Form
                        oForm = SBO_Application.Forms.Item(FormUID)
                        oForm.EnableMenu("1292", True) 'Add Row
                        oForm.EnableMenu("1293", True) 'Delete Row

                        oForm.EnableMenu("1287", True) 'Duplicate

                    End If

                    'If pVal.ItemUID = "OptimMtx" And _
                    '    ((pVal.ColUID = "P" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                    '     (pVal.ColUID = "L" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)) Then

                    If pVal.ItemUID = "OptimMtx" And _
                        ((pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                        (pVal.ColUID = "P" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                         (pVal.ColUID = "QtyPotong" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) Or _
                         (pVal.ColUID = "L" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)) Then

                        Dim oForm As SAPbouiCOM.Form = Nothing
                        Dim oEditText As SAPbouiCOM.EditText = Nothing

                        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim idx As Long
                        Dim gtabc As Double
                        Dim gtaloc As Double
                        Dim gtplanpdo As Double

                        oForm = SBO_Application.Forms.Item(FormUID)
                        oMatrix = oForm.Items.Item("OptimMtx").Specific

                        oForm.Freeze(True)

                        'Total AxBxC = Jumlah Potong x P x L
                        oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = _
                        Math.Round( _
                            CDbl(IIf(oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("QtyPotong").Cells.Item(pVal.Row).Specific.value)) _
                            * CDbl(IIf(oMatrix.Columns.Item("P").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("P").Cells.Item(pVal.Row).Specific.value)) _
                            * CDbl(IIf(oMatrix.Columns.Item("L").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("L").Cells.Item(pVal.Row).Specific.value)) _
                            / 1000000 _
                        , 4)

                        gtabc = 0
                        gtaloc = 0
                        gtplanpdo = 0

                        If CDbl(IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, oForm.Items.Item("GTaloc").Specific.value)) <> 0 Then
                            For idx = 1 To oMatrix.RowCount
                                gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                                'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                                gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                            Next
                        End If

                        'gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                        oForm.Items.Item("GTabc").Specific.value = gtabc
                        'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                        'oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo

                        'oEditText.Value = _
                        '    IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) _
                        '    - IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) _
                        '    - IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)

                        ' 2011-02-28
                        oEditText = oForm.Items.Item("GTaloc").Specific
                        oEditText.Value = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                        + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                        'Allocated Waste
                        oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = _
                        Math.Round( _
                            CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value)) _
                            / CDbl(IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)) _
                            * CDbl(IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, oForm.Items.Item("GTaloc").Specific.value)) _
                        , 4)


                        'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                        oMatrix.Columns.Item("PlanPdIsue").Cells.Item(pVal.Row).Specific.value = _
                        Math.Round( _
                            CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value)) + _
                            CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value)) _
                        , 4)

                        'Dim totalABC As Double
                        'Dim alocatedWaste As Double

                        'totalABC = IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("TotalABC").Cells.Item(pVal.Row).Specific.value))
                        'alocatedWaste = IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("AlocWaste").Cells.Item(pVal.Row).Specific.value))

                        Dim TotalWaste As Double
                        Dim Kolom As Double

                        'oForm = SBO_Application.Forms.Item(FormUID)

                        TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value))

                        ' 2011-02-28
                        TotalWaste = CDbl(IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value)) _
                            + CDbl(IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value))

                        Kolom = oForm.Items.Item("GTabc").Specific.value

                        'oMatrix = oForm.Items.Item("OptimMtx").Specific

                        If CDbl(IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, oForm.Items.Item("GTaloc").Specific.value)) <> 0 Then

                            For Row = 1 To oMatrix.RowCount
                                'Allocated Waste
                                oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = _
                                Math.Round( _
                                    (CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, _
                                    oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                    / Kolom) * TotalWaste _
                                , 4)


                                'Qty Plan PdO Issue = Total AxBxC + Allocated Waste
                                oMatrix.Columns.Item("PlanPdIsue").Cells.Item(Row).Specific.value = _
                                Math.Round( _
                                    CDbl(IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value)) _
                                    + CDbl(IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value)) _
                                , 4)

                                'totalABC = IIf(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("TotalABC").Cells.Item(Row).Specific.value))
                                'alocatedWaste = IIf(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value = "", 0, CDbl(oMatrix.Columns.Item("AlocWaste").Cells.Item(Row).Specific.value))

                            Next

                        End If

                        gtplanpdo = 0
                        'If IIf(oForm.Items.Item("GTaloc").Specific.value = "", 0, CDbl(oForm.Items.Item("GTaloc").Specific.value)) <> 0 Then
                        For idx = 1 To oMatrix.RowCount
                            'gtabc += IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("TotalABC", idx).string))
                            'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                            gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))

                        Next
                        'End If

                        'gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                        'oForm.Items.Item("GTabc").Specific.value = gtabc
                        'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                        oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo


                        'BubbleEvent = False

                        'oForm.Refresh()
                        oForm.Freeze(False)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText)
                        GC.Collect()

                    End If

                    'If pVal.ItemUID = "OptimMtx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK Then
                    '    Dim oForm As SAPbouiCOM.Form
                    '    Dim oEditText As SAPbouiCOM.EditText
                    '    Dim oMatrix As SAPbouiCOM.Matrix

                    '    oForm = SBO_Application.Forms.Item(FormUID)
                    '    oMatrix = oForm.Items.Item("OptimMtx").Specific

                    '    Dim idx As Integer
                    '    Dim gtabc As Double

                    '    For idx = 1 To oMatrix.RowCount
                    '        gtabc = gtabc + IIf(oMatrix.GetCellSpecific(11, idx).string = "", 0, CDbl(oMatrix.GetCellSpecific(11, idx).string))
                    '        'gtabc = gtabc + IIf(oMatrix.GetCellSpecific(7, idx).string = "", 0, CDbl(oMatrix.GetCellSpecific(7, idx).string))
                    '        'oMatrix.Columns.Item("TotalABC").Cells.Item(idx).Specific.value = 5
                    '    Next

                    '    oEditText = oForm.Items.Item("GTabc").Specific
                    '    oEditText.Value = gtabc

                    'End If

                    'If pVal.ItemUID = "OptimMtx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    '    Dim oForm As SAPbouiCOM.Form
                    '    Dim oItem As SAPbouiCOM.Item
                    '    Dim oMatrix As SAPbouiCOM.Matrix
                    '    Dim lrowcount As Long
                    '    'Dim oEditText As SAPbouiCOM.EditText

                    '    oForm = SBO_Application.Forms.Item(FormUID)
                    '    oItem = oForm.Items.Item("OptimMtx")
                    '    oMatrix = oItem.Specific
                    '    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

                    '    If pVal.CharPressed = 13 Then
                    '        'oForm.DataSources.DBDataSources.Item(oForm.PaneLevel + 1).Clear()
                    '        lrowcount = oMatrix.RowCount
                    '        oMatrix.AddRow(1, lrowcount)

                    '        'oMatrix.SelectRow(1 + lrowcount, True, False)
                    '        'oForm.EnableMenu("1292", True) 'Add Row
                    '        'oForm.EnableMenu("1293", True) 'Delete Row

                    '        'Reset blank for new ROW
                    '        oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").Clear()
                    '        oMatrix.SetLineData(oMatrix.RowCount)
                    '        oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount
                    '        oMatrix.Columns.Item("PdO#").Cells.Item(oMatrix.VisualRowCount).Specific.string = "111"
                    '        'oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount
                    '        'oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = _
                    '        'oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount - 1).Specific.string + 1

                    '        'oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").SetValue(5, 5, 18)
                    '        'oEditText.Value = "135"
                    '        'oMatrix.FlushToDataSource()
                    '        'oMatrix.LoadFromDataSource()

                    '    End If
                    'End If

            End Select
        End If

    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        If pVal.BeforeAction = False Then
            Select Case pVal.MenuUID
                Case "PROD01_01"
                    SOToMFGEntry()
                Case "PROD01_03"
                    OptimizationEntry()
                Case "PROD01_04"
                    OutDelEntry()
                Case "PROD01_05"
                    ProductionStatus()
            End Select
        End If
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

                    If oForm.UniqueID = "mds_p3" Then
                        'Dim oForm As SAPbouiCOM.Form = Nothing
                        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim oDBDataSource As SAPbouiCOM.DBDataSource

                        'oForm = SBO_Application.Forms.Item(FormUID)

                        oMatrix = oForm.Items.Item("OptimMtx").Specific

                        oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                        oForm.Freeze(True)

                        Dim docnum As Integer

                        docnum = oDBDataSource.GetValue("docnum", 0)

                        Dim idx As Long
                        Dim gtabc As Double
                        Dim gtaloc As Double
                        Dim gtplanpdo As Double

                        gtabc = 0
                        gtaloc = 0
                        gtplanpdo = 0
                        For idx = 1 To oMatrix.RowCount
                            gtabc += CDbl(IIf(oMatrix.GetCellSpecific("TotalABC", idx).string = "", 0, oMatrix.GetCellSpecific("TotalABC", idx).string))
                            'gtaloc += IIf(oMatrix.GetCellSpecific("AlocWaste", idx).string = "", 0, CDbl(oMatrix.GetCellSpecific("AlocWaste", idx).string))
                            gtplanpdo += CDbl(IIf(oMatrix.GetCellSpecific("PlanPdIsue", idx).string = "", 0, oMatrix.GetCellSpecific("PlanPdIsue", idx).string))
                            'oForm.Items.Item("#").Specific.value = idx
                            'oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.value = oMatrix.VisualRowCount
                            oMatrix.Columns.Item("#").Cells.Item(CInt(idx)).Specific.value = idx

                        Next

                        gtplanpdo += oForm.Items.Item("SisaKcUtuh").Specific.value
                        oForm.Items.Item("GTabc").Specific.value = gtabc
                        'oForm.Items.Item("GTaloc").Specific.value = gtaloc
                        oForm.Items.Item("GTplanPdO").Specific.value = gtplanpdo


                        'IIf(oForm.Items.Item("LuasKaca").Specific.value = "", 0, oForm.Items.Item("LuasKaca").Specific.value) _
                        '- IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value) _
                        '- IIf(oForm.Items.Item("GTabc").Specific.value = "", 0, oForm.Items.Item("GTabc").Specific.value)

                        ' 2011-02-28
                        oForm.Items.Item("GTaloc").Specific.value = _
                        IIf(oForm.Items.Item("TotalWaste").Specific.value = "", 0, oForm.Items.Item("TotalWaste").Specific.value) _
                        + IIf(oForm.Items.Item("SisaKcUtuh").Specific.value = "", 0, oForm.Items.Item("SisaKcUtuh").Specific.value)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDBDataSource)
                        oMatrix = Nothing
                        oDBDataSource = Nothing

                        GC.Collect()


                        'BubbleEvent = False

                        oForm.Freeze(False)


                    End If

                Case "1291" ' Last Record

                Case "1292" ' Add a row
                    'form "mds_p3" = Optimization Entry
                    If oForm.UniqueID = "mds_p3" Then

                        ''MsgBox("Add a row optimization entry")

                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = oForm.Items.Item("OptimMtx").Specific
                        oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").Clear()
                        oMatrix.AddRow()
                        oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount
                        oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                        oMatrix = Nothing
                        GC.Collect()

                    End If
                    'MsgBox("Add a row")

                Case "1293" ' Delete a row
                    'Dim oMatrix As SAPbouiCOM.Matrix
                    'oMatrix = oForm.Items.Item("OptimMtx").Specific
                    'oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").Clear()
                    'oMatrix.DeleteRow(oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific.string)
                    '' = oMatrix.VisualRowCount
                    ''oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.string = oMatrix.VisualRowCount

                    'If oForm.UniqueID = "mds_p3" Then

                    '    Dim oMatrix As SAPbouiCOM.Matrix
                    '    oMatrix = oForm.Items.Item("OptimMtx").Specific
                    '    'Dim oDBDataSource As SAPbouiCOM.DBDataSource


                    '    'oMatrix = oForm.Items.Item("OptimMtx").Specific

                    '    oForm.Freeze(True)
                    '    oForm.DataSources.DBDataSources.Item("@MIS_OPTIML").RemoveRecord(oMatrix.VisualRowCount)
                    '    oMatrix.FlushToDataSource()
                    '    'oMatrix.LoadFromDataSource()

                    '    'oDBDataSource = oForm.DataSources.DBDataSources.Item("@MIS_OPTIM")

                    '    oForm.Freeze(False)


                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                    '    BubbleEvent = False

                    'End If
                    'MsgBox("Delete a row")
                Case "1282"

                    'MsgBox("add new doc!")
                    If oForm.UniqueID = "mds_p3" Then
                        'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'Dim oForm As SAPbouiCOM.Form
                        'oForm = SBO_Application.Forms.Item(FormUID)

                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                        oForm.Items.Item("GTabc").Specific.value = 0
                        oForm.Items.Item("GTaloc").Specific.value = 0
                        oForm.Items.Item("GTplanPdO").Specific.value = 0
                        oForm.Items.Item("QtyLembar").Specific.value = 1
                        oForm.Items.Item("TotalWaste").Specific.value = 0
                        oForm.Items.Item("TotWastPct").Specific.value = 0
                        oForm.Items.Item("SisaKcUtuh").Specific.value = 0
                        oForm.Items.Item("KcSisaPctg").Specific.value = 0

                        'End If
                    End If

            End Select

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
            GC.Collect()

        End If
    End Sub

End Class
