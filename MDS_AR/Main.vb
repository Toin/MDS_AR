
Option Explicit On
Option Strict Off

Module Main

    Public Sub Main()

        Dim SBOConn As SBOConnection
        Dim ClsName As String
        ClsName = "Main"

        Try

            SBOConn = New SBOConnection
            'MsgBox("Connected! hi!")

            Dim MDS_AR As MDS_T3
            MDS_AR = New MDS_T3

            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            MsgBox(ClsName & " Addon failed! SBO Not Running!")
        End Try

    End Sub

End Module
