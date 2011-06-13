Option Strict Off
Option Explicit On

Public Class SBOConnection

    Public Shared SBOApplication As Object
    Public Shared SBOCompany As Object
    Public CCompany As SAPbobsCOM.Company
    Public WithEvents CApplication As SAPbouiCOM.Application


    Public Sub New()
        MyBase.New()

        Try
            SetApplication()

            '//*************************************************************
            '// Connect to DI
            '//*************************************************************

            CCompany = New SAPbobsCOM.Company

            '//get DI company (via UI)
            CCompany = CApplication.Company.GetDICompany()
            SBOCompany = CCompany


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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

        'sConnectionString = Environment.GetCommandLineArgs.GetValue(1)

        sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs.GetValue(1))

        '// connect to a running SBO Application
        SboGuiApi.Connect(sConnectionString)

        '// get an initialized application object
        CApplication = SboGuiApi.GetApplication()
        SBOApplication = CApplication

    End Sub

    Public Function SBODateisValid(ByVal pDate As String) As Boolean
        Dim vYear As String
        Dim vMonth As String
        Dim vDay As String

        SBODateisValid = False

        Select Case Len(pDate)
            Case Is < 8
                Return False
            Case 8
                vYear = Left(pDate, 4)
                vMonth = Mid(pDate, 5, 2)
                vDay = Right(pDate, 2)
                If IsDate(vYear & "/" & vMonth & "/" & vDay) Then
                    Return True
                End If
                If Mid(pDate, 5, 2) > 12 Or Mid(pDate, 5, 2) < 1 Then
                    Return False
                End If

            Case Else
                Return True
        End Select

    End Function

    Public Sub AddCFL1(ByVal oForm As SAPbouiCOM.Form, ByVal oLinkedObject As SAPbouiCOM.BoLinkedObject, _
            ByVal CFLtxt As String, ByVal CFLbtn As String, _
            ByVal CFLCondField As String, _
            ByVal CFLCondOperator As SAPbouiCOM.BoConditionOperation, _
            ByVal CFLCondFieldValue As String)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFLConds As SAPbouiCOM.Conditions
            Dim oCFLCond As SAPbouiCOM.Condition


            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Add 2 CFL
            'one for button (windows popup) & one for edit textbox
            oCFLCreationParams.MultiSelection = False
            '            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            Dim oLinkedObjectType As SAPbouiCOM.BoLinkedObject
            oLinkedObjectType = oLinkedObject
            oCFLCreationParams.ObjectType = oLinkedObject ' "2"-> BP Master
            oCFLCreationParams.UniqueID = CFLtxt ' "CFL1" -> txtbox cfl Field

            oCFL = oCFLs.Add(oCFLCreationParams)

            'Add conditions to CFL1
            oCFLConds = oCFL.GetConditions()

            oCFLCond = oCFLConds.Add()
            oCFLCond.Alias = CFLCondField ' "CardType" -> BP Master where CardType = ??
            'oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCFLCond.Operation = CFLCondOperator
            oCFLCond.CondVal = CFLCondFieldValue ' "C" -> CardType value = C -> BP Customer data 
            oCFL.SetConditions(oCFLConds)

            oCFLCreationParams.UniqueID = CFLbtn ' "CFL2" -> button CFL field
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Public Sub AddChooseFromListForMatrix(ByVal oForm As SAPbouiCOM.Form, ByVal oLinkedObject As SAPbouiCOM.BoLinkedObject, _
        ByVal CFLtxt As String, _
        ByVal CFLCondField As String, _
        ByVal CFLCondOperator As SAPbouiCOM.BoConditionOperation, _
        ByVal CFLCondFieldValue As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFLConds As SAPbouiCOM.Conditions
            Dim oCFLCond As SAPbouiCOM.Condition

            oCFLs = SBOApplication.Forms.ActiveForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = oLinkedObject
            oCFLCreationParams.UniqueID = CFLtxt

            oCFL = oCFLs.Add(oCFLCreationParams)

            'Add conditions to CFL1
            oCFLConds = oCFL.GetConditions()

            oCFLCond = oCFLConds.Add()
            oCFLCond.Alias = CFLCondField ' "CardType" -> BP Master where CardType = ??
            'oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCFLCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCFLCond.Operation = CFLCondOperator
            oCFLCond.CondVal = CFLCondFieldValue ' "C" -> CardType value = C -> BP Customer data 
            oCFL.SetConditions(oCFLConds)

        Catch ex As Exception
            ' app.MessageBox(ex.Message)
        End Try

    End Sub

End Class
