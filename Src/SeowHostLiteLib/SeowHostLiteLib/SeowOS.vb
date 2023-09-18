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
'* Name: Seow操作系统|SeowOS
'* Author: Seowsoft
'* Describe: 操作系统信息处理|OS information processing
'* Home Url: https://www.seowsoft.com
'* Version: 1.0
'* Create Time: 3/9/2023
'********************************************************************
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.Logging
Imports PigCmdLib
Imports PigToolsLiteLib

Public Class SeowOS
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.0.6"

    Public Enum EnmOSType
        Unknow = 0
        Windows = 1
        Linux = 2
        OSX = 3
        FreeBSD = 4
    End Enum

    Private ReadOnly Property mPigSysCmd As New PigSysCmd
    Public ReadOnly Property OSCaption As String
    Public Sub New()
        MyBase.New(CLS_VERSION)
        Dim LOG As New PigStepLog("New")
        Try
            Dim strOSCaption As String = ""
            LOG.StepName = "GetOSCaption"
            LOG.Ret = Me.mPigSysCmd.GetOSCaption(strOSCaption)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.OSCaption = strOSCaption
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Private mOSType As EnmOSType = EnmOSType.Unknow
    Public ReadOnly Property OSType As EnmOSType
        Get
            Try
                If mOSType = EnmOSType.Unknow Then
#If NETFRAMEWORK Then
                    mOSType = EnmOSType.Windows
#Else
                    If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) = True Then
                        mOSType = EnmOSType.Windows
                    ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.Linux) = True Then
                        mOSType = EnmOSType.Linux
                    ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.OSX) = True Then
                        mOSType = EnmOSType.OSX
                    ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.FreeBSD) = True Then
                        mOSType = EnmOSType.FreeBSD
                    End If
#End If
                End If
                Return mOSType
            Catch ex As Exception
                mOSType = EnmOSType.Unknow
                Me.SetSubErrInf("OSType.Get", ex)
                Return mOSType
            End Try
        End Get
    End Property

    Public ReadOnly Property IsWindowsServer As Boolean
        Get
            Try
                If Me.OSType = EnmOSType.Windows Then
                    If InStr(LCase(Me.OSCaption), " server ") > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                Me.SetSubErrInf("IsWinServer.Get", ex)
                Return False
            End Try
        End Get
    End Property

End Class
