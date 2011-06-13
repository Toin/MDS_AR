Imports System.IO
Imports System.Reflection
Imports Microsoft.Win32

<System.ComponentModel.RunInstaller(True)> _
Public Class CustomInstallerClass
  Inherits System.Configuration.Install.Installer

  ' Declaring the functions inside "AddOnInstallAPI.dll"

  'SetAddOnFolder - Notify to B1 that installation of Add-On is done in another folder. Return 0 if ok.
  Declare Function SetAddOnFolder Lib "AddOnInstallAPI.dll" (ByVal strPath As String) As Int32

  'RestartNeeded - Use it if your installation requires a restart, it will cause
  'the SBO application to close itself after the installation is complete.
  Declare Function RestartNeeded Lib "AddOnInstallAPI.dll" () As Int32

  ' New functions added from 2005 version
#If Version = "2005" Then
  'EndInstallEx - Notify to B1 that installation of Add-On is finished, 
  ' the isSucceed flag notify if installation finished ok or canceled by some reason. Return 0 if ok.
  Declare Function EndInstallEx Lib "AddOnInstallAPI.dll" (ByVal strPath As String, ByVal isSucceed As Boolean) As Int32

  'EndUninstall - Notify to B1 that uninstallation of Add-On is finished, 
  ' the isSucceed flag notify is installation finished ok or canceled by some reason. Return 0 if ok.
  Declare Function EndUninstall Lib "AddOnInstallAPI.dll" (ByVal strPath As String, ByVal isSucceed As Boolean) As Int32

  'B1Info - Returns B1 version.
  Declare Function B1Info Lib "AddOnInstallAPI.dll" (ByVal strB1Info As String, ByVal maxLen As Int32) As Int32
#Else
  'EndInstall - Notify to B1 that installation of Add-On is finished. Return 0 if ok. 
  Declare Function EndInstall Lib "AddOnInstallAPI.dll" () As Int32
#End If

  ' Called at installation time
  ' No need to implement different behavior than the basic 
  Public Overrides Sub Install(ByVal stateSaver As IDictionary)
    MyBase.Install(stateSaver)
  End Sub

  ' Called just at the end of the installation
  ' Alert B1 that addon installation has finished 
  Public Overrides Sub Commit(ByVal savedState As IDictionary)

    Dim strCmdLineElements(2) As String
    Dim targetDir As String
    Dim ret As Int32

    Try
      ' First commit
      MyBase.Commit(savedState)

      ' Get target dir from Setup 
      targetDir = Context.Parameters.Item("target")

      '' Set current directory where AddOnInstallAPI.dll is located
      Environment.CurrentDirectory = GetStrDllPath() ' For Dll function calls will work

      '' Send B1 the AddOnPath where the addon has been installed
      ret = SetAddOnFolder(targetDir)
      If (ret <> 0) Then
        MsgBox("Error (" + ret + ") at SetAddOnFolder function call: " + targetDir)
      End If

      ' Alert B1 if restart is needed after installation
      If (IsRestartNeeded()) Then
        ret = RestartNeeded()
        If (ret <> 0) Then
          MsgBox("Error (" + ret + ") at RestartNeeded function call")
        End If
      End If

#If Version = "2005" Then
      '' Call B1 EndInstallEx, succeeded true
      ret = EndInstallEx(targetDir, True)
      if ((IsRestartNeeded() and ret <> -1) or _
          (Not IsRestartNeeded() and ret <> 0)) Then
        MsgBox("Error (" + ret.ToString() + ") at EndInstallEx(True) call at the end of the AddOn installation: " + targetDir)
      End If
#Else
      '' Call B1 EndInstall (no parameters)
      ret = EndInstall()
      If (ret <> 0) Then
        MsgBox("Error (" + ret.ToString() + ") at EndInstallEx() call at the end of the AddOn installation")
      End If
#End If

    Catch ex As Exception
      MsgBox("Error during AddOn installation " + ex.Message)

#If Version = "2005" Then
      ' Alert B1 if installation failed
      ' No 2004 equivalent
      ret = EndInstallEx("", False)
      If (ret <> -1 And ret <> 0) Then
        MsgBox("Error (" + ret.ToString() + ") at EndInstallEx(False) at the end of the AddOn installation")
      End If
#End If

    End Try

  End Sub 'Commit

  Public Overrides Sub UnInstall(ByVal stateSaver As IDictionary)

    MyBase.Uninstall(stateSaver)

#If Version = "2005" Then
    ' Alert B1 that addon uninstall has finished
    ' No equivalent for 2004 version

    '' Set current directory where AddOnInstallAPI.dll is located
    Environment.CurrentDirectory = GetStrDllPath() ' For Dll function calls will work

    ' call AddOnInstallAPI.dll functions 
    Dim ret As Int32
    ret = EndUninstall("", True)
    If (ret <> 0) Then
      MsgBox("Error during EndInstallEx("", False) at the end of the AddOn uninstallation")
    End If
#End If
  End Sub

  Public Overrides Sub Rollback(ByVal stateSaver As IDictionary)

    MyBase.Rollback(stateSaver)

#If Version = "2005" Then
    ' Alert B1 that addon installation has failed
    ' No equivalent for 2004 version

    '' Set current directory where AddOnInstallAPI.dll is located
    Environment.CurrentDirectory = GetStrDllPath() ' For Dll function calls will work

    ' call EndInstallEx
    Dim ret As Int32
    ret = EndInstallEx("", False)
    If (ret <> -1 And ret <> 0) Then
      MsgBox("Error during EndInstallEx(False) at the end of the AddOn installation rollback")
    End If
#End If
  End Sub

  ' Read from registry the path where AddOnInstallAPI.dll is located
  ' The path is saved during addon installation by this installer application
  Private Function GetStrDllPath() As String
    Dim strArgs As String
    Dim strCmdLineElements As String()
    Dim addOnLibPath As String

    strArgs = GetParamsFromRegistry()

    strCmdLineElements = strArgs.Split("|")

    ' Get the "AddOnInstallAPI.dll" path
    addOnLibPath = strCmdLineElements(1)
    addOnLibPath = addOnLibPath.Remove((addOnLibPath.Length - 19), 19) ' Only the path is needed

    Return addOnLibPath

  End Function

  Private Function GetParamsFromRegistry() As String

    Dim regParam As RegistryKey
    Dim keyValue As String
    Dim subKeyValue As String
    Dim strParams As String = ""

    Dim addOnInstallInfo As AddOnInstaller.AddOnInstallInfo
    addOnInstallInfo = New AddOnInstaller.AddOnInstallInfo

    keyValue = "SOFTWARE\\SAP\\SAP Manage\\SAP Business One\\InstallAddOn\\" _
      + addOnInstallInfo.PartnerName + "\\" + addOnInstallInfo.AddOnName
    subKeyValue = "Params"
    regParam = Registry.CurrentUser.OpenSubKey(keyValue, False)

    If (Not regParam Is Nothing) Then
      strParams = regParam.GetValue(subKeyValue, "")
      regParam.Close()
    End If
    Return strParams

  End Function

  Private Function IsRestartNeeded() As Boolean

    Dim regParam As RegistryKey
    Dim keyValue As String
    Dim subKeyValue As String
    Dim strRestart As String

    Dim addOnInstallInfo As AddOnInstaller.AddOnInstallInfo
    addOnInstallInfo = New AddOnInstaller.AddOnInstallInfo

    keyValue = "SOFTWARE\\SAP\\SAP Manage\\SAP Business One\\InstallAddOn\\" _
      + addOnInstallInfo.PartnerName + "\\" + addOnInstallInfo.AddOnName
    subKeyValue = "RestartNeeded"
    regParam = Registry.CurrentUser.OpenSubKey(keyValue, False)

    If (Not regParam Is Nothing) Then
      strRestart = regParam.GetValue(subKeyValue, "N")
      regParam.Close()
      If (strRestart = "Y") Then
        Return True
      Else
        Return False
      End If
    End If

    Return False

  End Function

End Class
