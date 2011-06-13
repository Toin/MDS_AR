Option Strict Off
Option Explicit On

Public Class MIS_Utils
    Public Function fctFormatDate(ByVal pdate As Date, ByVal oCompany As SAPbobsCOM.Company, Optional ByVal sngFormat As Integer = 5) As String
        Dim strSeparator As String
        Dim oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        fctFormatDate = ""

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate
        strSeparator = oAdminInfo.DateSeparator

        Select Case sngFormat
            Case 0
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "yy")
            Case 1
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MM") + strSeparator + "20" + Format(pdate, "yy")
            Case 2
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + Format(pdate, "yy")
            Case 3
                fctFormatDate = Format(pdate, "MM") + strSeparator + Format(pdate, "dd") + strSeparator + "20" + Format(pdate, "yy")
            Case 4
                fctFormatDate = "20" + Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
            Case 5
                fctFormatDate = Format(pdate, "dd") + strSeparator + Format(pdate, "MMMM") + strSeparator + Format(pdate, "yyyy")
            Case 6
                fctFormatDate = Format(pdate, "yy") + strSeparator + Format(pdate, "MM") + strSeparator + Format(pdate, "dd")
        End Select

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
    End Function

    Public Function fctFormatDateSave(ByVal oCompany As SAPbobsCOM.Company, ByVal pdate As String, ByVal sngFormat As Integer) As String
        Dim strFormat As String
        Dim strMonth As String
        Dim intLength As Integer
        Static oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        On Error GoTo ErrorHandler

        strMonth = "JANUARY01FEBRUARY02MARCH03APRIL04MAY05JUNE06JULY07AUGUST08SEPTEMBER09OCTOBER10NOVEMBER11DECEMBER12"

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate

        If pdate = "" Then
            GoTo ErrorHandler
        End If

        Select Case sngFormat
            Case 0
                fctFormatDateSave = "20" + Right(pdate, 2) + "/" + Mid(pdate, 4, 2) + "/" + Left(pdate, 2)
            Case 1
                fctFormatDateSave = Right(pdate, 4) + "/" + Mid(pdate, 4, 2) + "/" + Left(pdate, 2)
            Case 2
                fctFormatDateSave = "20" + Right(pdate, 2) + "/" + Left(pdate, 2) + "/" + Mid(pdate, 4, 2)
            Case 3
                fctFormatDateSave = Right(pdate, 4) + "/" + Left(pdate, 2) + "/" + Mid(pdate, 4, 2)
            Case 4
                fctFormatDateSave = Left(pdate, 4) + "/" + Mid(pdate, 6, 2) + "/" + Right(pdate, 2)
            Case 5
                intLength = InStr(1, strMonth, UCase(Mid(pdate, 4, Len(pdate) - 8))) + Len(Mid(pdate, 4, Len(pdate) - 8))
                fctFormatDateSave = Left(pdate, 4) + "/" + Mid(pdate, 4, 2) + "/" + Right(pdate, 2)
        End Select

        GoTo SetNothing

ErrorHandler:
        fctFormatDateSave = ""

SetNothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
        oGetCompanyService = Nothing
        oAdminInfo = Nothing
    End Function


    Public Function fctFormatNumSBO(ByVal pNum As String, ByVal oCompany As SAPbobsCOM.Company) As String
        Dim objGetCompanyService As SAPbobsCOM.CompanyService
        Dim objAdminInfo As SAPbobsCOM.AdminInfo
        Dim strDecSep As String
        Dim strThousSep As String

        objGetCompanyService = oCompany.GetCompanyService
        objAdminInfo = objGetCompanyService.GetAdminInfo

        strDecSep = objAdminInfo.DecimalSeparator
        strThousSep = objAdminInfo.ThousandsSeparator

        'no need use thousand separator
        pNum = Replace(pNum, strThousSep, "")

        objGetCompanyService = Nothing
        objAdminInfo = Nothing

        fctFormatNumSBO = (pNum)
    End Function

End Class
