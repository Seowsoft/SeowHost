'********************************************************************
'* Copyright 2023 Seowsoft
'*
'* Licensed under the Apache License, Version 2.0 (the "License");
'* you may Not use this file except in compliance with the License.
'* You may obtain a copy of the License at
'*
'*     http://www.apache.org/licenses/LICENSE-2.0
'*
'* Unless required by applicable law Or agreed to in writing, software
'* distributed under the License Is distributed on an "AS IS" BASIS,
'* WITHOUT WARRANTIES Or CONDITIONS OF ANY KIND, either express Or implied.
'* See the License for the specific language governing permissions And
'* limitations under the License.
'********************************************************************
'* Name: Seow主机|SeowHost
'* Author: Seowsoft
'* Describe: 主机信息处理|Host information processing
'* Home Url: https://www.seowsoft.com
'* Version: 1.0
'* Create Time: 3/9/2023
'********************************************************************
Imports PigCmdLib
Imports PigToolsLiteLib

Public Class SeowHost
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.0.6"
    Private ReadOnly Property mPigSysCmd As New PigSysCmd
    Private ReadOnly Property mPigFunc As New PigFunc
    Public ReadOnly Property HostName As String
    Public ReadOnly Property HostID As String
    Public ReadOnly Property UUID As String

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Dim LOG As New PigStepLog("New")
        Try
            Dim strUUID As String = ""
            LOG.StepName = "GetUUID"
            LOG.Ret = Me.mPigSysCmd.GetUUID(strUUID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If strUUID = "" Then Throw New Exception("Unable to get UUID")
            Me.UUID = strUUID
            Dim strHostID As String = ""
            LOG.StepName = "GetTextPigMD5(HostID)"
            LOG.Ret = Me.mPigFunc.GetTextPigMD5(Me.UUID, PigMD5.enmTextType.UTF8, strHostID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.HostID = UCase(strHostID)
            Me.HostName = Me.mPigFunc.GetComputerName
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub
    Property mOS As SeowOS
    Public ReadOnly Property OS As SeowOS
        Get
            Try
                If mOS Is Nothing Then
                    mOS = New SeowOS
                    If mOS.LastErr <> "" Then Throw New Exception(mOS.LastErr)
                End If
                Return mOS
            Catch ex As Exception
                Me.SetSubErrInf("OS.Get", ex)
                Return Nothing
            End Try
        End Get
    End Property


End Class
