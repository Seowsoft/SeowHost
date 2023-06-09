﻿'**********************************
'* Name: 我的用户|MyUser
'* Author: Seowsoft
'* License: Copyright (c) 2021-2022 Seowsoft, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 我的用户信息|My user information
'* Home Url: https://www.seowsoft.com
'* Version: 1.5
'* Create Time: 30/10/2021
'* 1.1    2/4/2022   Add 
'* 1.2    24/7/2022   Change name and initialization
'* 1.3    29/7/2022   Modify Imports
'* 1.4    30/7/2022   Modify Imports
'* 1.5    4/10/2022   Modify Imports
'**********************************
Imports PigToolsLiteLib
Imports PigCmdLib

Friend Class PigUser
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.3.2"

    Public ReadOnly Property UserName As String
    Private Property mPigFunc As New PigFunc

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.UserName = Me.mPigFunc.GetUserName
    End Sub


End Class
