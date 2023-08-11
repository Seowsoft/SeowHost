'********************************************************************
'* Copyright 2021-2023 Seowsoft
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
'* Name: 表扩展信息|Table Extension Information
'* Author: Seowsoft
'* Describe: 主机信息处理|Table Extension Information Processing
'* Home Url: https://www.seowsoft.com
'* Version: 1.2
'* Create Time: 7/28/2023
'* 1.1    29/7/2023   Modify New
'* 1.2    30/7/2023   Add RefSQL,SetExtValue, modify mGetExtValue
'********************************************************************
Imports PigToolsLiteLib
Imports PigSQLSrvLib
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If

Friend Class fTabExt
	Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.2.10"

	Public Enum EnmWhatTab
		HostInf = 1
		HFFolderInf = 2
		HFFileInf = 3
		HFDirInf = 5
	End Enum
	Private ReadOnly Property mMainTabName As String
	Private ReadOnly Property mExtTabName As String
	Private ReadOnly Property mMainPKeyName As String
	Private Property mConnSQLSrv As ConnSQLSrv
	Private mIsDirtyRead As Boolean = True

	Private Property mPigFunc As New PigFunc
	Public Property IsDirtyRead As Boolean
		Get
			Return mIsDirtyRead
		End Get
		Set(value As Boolean)
			mIsDirtyRead = value
			Me.RefSQL()
		End Set
	End Property

	Private Property mSQL_GetValue As String
	Private Property mSQL_GetLargeValue As String
	Private Property mSQL_SetValue As String
	Private Property mSQL_SetLargeValue As String

	Private Sub RefSQL()
		Try
			Me.mSQL_GetValue = "SELECT ExtValue FROM " & Me.mExtTabName
			If Me.IsDirtyRead = True Then Me.mSQL_GetValue &= " WITH (NOLOCK)"
			Me.mSQL_GetValue &= " WHERE " & Me.mMainPKeyName & "=@MainTabID AND ExtKey=@ExtKey"
			'--------------
			Me.mSQL_GetLargeValue = "SELECT le.LargeExtValue FROM LargeExtInf le"
			If Me.IsDirtyRead = True Then Me.mSQL_GetLargeValue &= " WITH (NOLOCK)"
			Me.mSQL_GetLargeValue &= " JOIN " & Me.mExtTabName & " e"
			If Me.IsDirtyRead = True Then Me.mSQL_GetLargeValue &= " WITH (NOLOCK)"
			Me.mSQL_GetLargeValue &= " ON e.LargeExtID=le.LargeExtID"
			Me.mSQL_GetLargeValue &= " WHERE e." & Me.mMainPKeyName & "=@MainTabID AND e.ExtKey=@ExtKey"
			'--------------
			Me.mSQL_SetValue = "MERGE INTO " & Me.mExtTabName & " t USING (SELECT @MainTabID " & Me.mMainPKeyName & ",@ExtKey ExtKey,@ExtValue ExtValue) s ON t." & Me.mMainPKeyName & "=s." & Me.mMainPKeyName & " AND t.ExtKey=s.ExtKey"
			Me.mSQL_SetValue &= " WHEN MATCHED THEN UPDATE SET ExtValue=s.ExtValue,UpdateTime=GETDATE(),LargeExtID=NULL"
			Me.mSQL_SetValue &= " WHEN NOT MATCHED THEN INSERT (" & Me.mMainPKeyName & ",ExtKey,ExtValue)VALUES(s." & Me.mMainPKeyName & ",s.ExtKey,s.ExtValue);"
			'--------------
			Me.mSQL_SetLargeValue = "MERGE INTO _ptLargeExtInf t USING (SELECT @LargeExtID LargeExtID,@LargeExtValue LargeExtValue) s ON t.LargeExtID=s.LargeExtID"
			Me.mSQL_SetLargeValue &= " WHEN MATCHED THEN UPDATE SET LargeExtValue=s.LargeExtValue,UpdateTime=GETDATE()"
			Me.mSQL_SetLargeValue &= " WHEN NOT MATCHED THEN INSERT(LargeExtID,LargeExtValue)VALUES(s.LargeExtID,s.LargeExtValue);"
			Me.mSQL_SetLargeValue &= " MERGE INTO " & Me.mExtTabName & " t USING (SELECT " & Me.mMainPKeyName & ",ExtValue FROM " & Me.mExtTabName & " WHERE " & Me.mMainPKeyName & "=@MainTabID AND ExtKey=@ExtKey) s ON t." & Me.mMainPKeyName & "=s." & Me.mMainPKeyName
			Me.mSQL_SetLargeValue &= " WHEN MATCHED THEN UPDATE ExtValue=NULL,UpdateTime=GETDATE(),LargeExtID=s.LargeExtID"
			Me.mSQL_SetLargeValue &= " WHEN NOT MATCHED THEN  INSERT (" & Me.mMainPKeyName & ",ExtKey,LargeExtID) VALUES (@MainTabID,@ExtKey,@LargeExtID);"
			'--------------
		Catch ex As Exception
			Me.SetSubErrInf("RefSQL", ex)
		End Try
	End Sub

	Public Sub New(WhatTab As EnmWhatTab, ConnSQLSrv As ConnSQLSrv)
		MyBase.New(CLS_VERSION)
		Try
			Select Case WhatTab
				Case EnmWhatTab.HostInf
					Me.mMainTabName = "_ptHostInf"
					Me.mExtTabName = "_ptHostExtInf"
					Me.mMainPKeyName = "HostID"
				Case EnmWhatTab.HFFolderInf
					Me.mMainTabName = "_ptHFFolderInf"
					Me.mExtTabName = "_ptHFFolderExtInf"
					Me.mMainPKeyName = "FolderID"
				Case EnmWhatTab.HFFileInf
					Me.mMainTabName = "_ptHFFileInf"
					Me.mExtTabName = "_ptHFFileExtInf"
					Me.mMainPKeyName = "FileID"
				Case EnmWhatTab.HFDirInf
					Me.mMainTabName = "_ptHFDirInf"
					Me.mExtTabName = "_ptHFDirExtInf"
					Me.mMainPKeyName = "DirID"
				Case Else
					Throw New Exception("Unsupported " & WhatTab.ToString)
			End Select
			Me.RefSQL()
			Me.mConnSQLSrv = ConnSQLSrv
		Catch ex As Exception
			Me.SetSubErrInf("New", ex)
		End Try
	End Sub

	Public Function RefConn(ConnSQLSrv As ConnSQLSrv) As String
		Try
			Me.mConnSQLSrv = ConnSQLSrv
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("RefConn", ex)
		End Try
	End Function


	Public Function GetExtLongValue(MainTabID As String, ExtKey As String, Optional ByRef Res As String = "OK") As Long
		Dim strExtValue As String = ""
		Res = Me.mGetExtValue(MainTabID, ExtKey, strExtValue)
		Return Me.mPigFunc.GECLng(strExtValue)
	End Function

	Public Function GetExtDecValue(MainTabID As String, ExtKey As String, Optional ByRef Res As String = "OK") As Decimal
		Dim strExtValue As String = ""
		Res = Me.mGetExtValue(MainTabID, ExtKey, strExtValue)
		Return Me.mPigFunc.GEDec(strExtValue)
	End Function

	Public Function GetExtDateValue(MainTabID As String, ExtKey As String, Optional ByRef Res As String = "OK") As Date
		Dim strExtValue As String = ""
		Res = Me.mGetExtValue(MainTabID, ExtKey, strExtValue)
		Return Me.mPigFunc.GECDate(strExtValue)
	End Function

	Public Function GetExtBoolValue(MainTabID As String, ExtKey As String, Optional ByRef Res As String = "OK") As Boolean
		Dim strExtValue As String = ""
		Res = Me.mGetExtValue(MainTabID, ExtKey, strExtValue)
		Return Me.mPigFunc.GECBool(strExtValue)
	End Function

	Public Function GetExtBoolEmpTrueValue(MainTabID As String, ExtKey As String, Optional ByRef Res As String = "OK") As Boolean
		Dim strExtValue As String = ""
		Res = Me.mGetExtValue(MainTabID, ExtKey, strExtValue)
		If strExtValue = "" Then
			Return True
		Else
			Return Me.mPigFunc.GECBool(strExtValue)
		End If
	End Function

	Public Function GetExtStrValue(MainTabID As String, ExtKey As String, Optional ByRef Res As String = "OK") As String
		Try
			GetExtStrValue = ""
			Res = Me.mGetExtValue(MainTabID, ExtKey, GetExtStrValue)
		Catch ex As Exception
			Res = Me.GetSubErrInf("GetExtStrValue", ex)
			Return ""
		End Try
	End Function

	Private Function mGetExtValue(MainTabID As String, ExtKey As String, ByRef ExtValue As String) As String
		Dim LOG As New PigStepLog("mGetExtValue")
		Dim oCmdSQLSrvText As CmdSQLSrvText = Nothing
		Try
			LOG.StepName = "New mSQL_GetValue"
			oCmdSQLSrvText = New CmdSQLSrvText(Me.mSQL_GetValue)
			With oCmdSQLSrvText
				Select Case Me.mMainTabName
					Case "_ptHFFileInf"
						.AddPara("@MainTabID", Data.SqlDbType.Char, 36)
					Case Else
						.AddPara("@MainTabID", Data.SqlDbType.Char, 32)
				End Select
				.ParaValue("@MainTabID") = MainTabID
				.AddPara("@ExtKey", Data.SqlDbType.NVarChar, 128)
				.ParaValue("@ExtKey") = ExtKey
			End With
			If Me.mConnSQLSrv.IsDBConnReady = False Then
				LOG.StepName = "OpenOrKeepActive"
				Me.mConnSQLSrv.OpenOrKeepActive()
				If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
			End If
			Dim bolIsLarge As Boolean = False
			LOG.StepName = "ActiveConnection"
			oCmdSQLSrvText.ActiveConnection = Me.mConnSQLSrv.Connection
			LOG.StepName = "Execute"
			Dim rsMain As Recordset = oCmdSQLSrvText.Execute()
			If oCmdSQLSrvText.LastErr <> "" Then Throw New Exception(oCmdSQLSrvText.LastErr)
			If rsMain.LastErr <> "" Then Throw New Exception(rsMain.LastErr)
			If rsMain.EOF = False Then
				ExtValue = rsMain.Fields.Item(0).StrValue
				If ExtValue = "" Then bolIsLarge = True
			End If
			If bolIsLarge = True Then
				LOG.StepName = "New mSQL_GetLargeValue"
				oCmdSQLSrvText = New CmdSQLSrvText(Me.mSQL_GetLargeValue)
				With oCmdSQLSrvText
					Select Case Me.mMainTabName
						Case "_ptHFFileInf"
							.AddPara("@MainTabID", Data.SqlDbType.Char, 36)
						Case Else
							.AddPara("@MainTabID", Data.SqlDbType.Char, 32)
					End Select
					.ParaValue("@MainTabID") = MainTabID
					.AddPara("@ExtKey", Data.SqlDbType.NVarChar, 128)
					.ParaValue("@ExtKey") = ExtKey
				End With
				LOG.StepName = "Execute"
				rsMain = oCmdSQLSrvText.Execute()
				If oCmdSQLSrvText.LastErr <> "" Then Throw New Exception(oCmdSQLSrvText.LastErr)
				If rsMain.LastErr <> "" Then Throw New Exception(rsMain.LastErr)
				If rsMain.EOF = False Then ExtValue = rsMain.Fields.Item(0).StrValue
			End If
			rsMain = Nothing
			oCmdSQLSrvText = Nothing
			Return "OK"
		Catch ex As Exception
			ExtValue = ""
			If oCmdSQLSrvText IsNot Nothing Then
				If oCmdSQLSrvText.SQLText <> "" Then
					LOG.AddStepNameInf(oCmdSQLSrvText.SQLText)
					LOG.AddStepNameInf(oCmdSQLSrvText.DebugStr)
				End If
			End If
			oCmdSQLSrvText = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function SetExtValue(MainTabID As String, ExtKey As String, ExtValue As String) As String
		Dim LOG As New PigStepLog("SetExtValue")
		Dim oCmdSQLSrvText As CmdSQLSrvText = Nothing
		Const EXT_VALUE_LEN As Integer = 768
		Try
			If Me.mConnSQLSrv.IsDBConnReady = False Then
				LOG.StepName = "OpenOrKeepActive"
				Me.mConnSQLSrv.OpenOrKeepActive()
				If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
			End If
			Dim bolIsLarge As Boolean = False
			If Len(ExtValue) > EXT_VALUE_LEN Then bolIsLarge = True
			If bolIsLarge = True Then
				LOG.StepName = "New mSQL_SetLargeValue"
				oCmdSQLSrvText = New CmdSQLSrvText(Me.mSQL_SetLargeValue)
			Else
				LOG.StepName = "New mSQL_SetValue"
				oCmdSQLSrvText = New CmdSQLSrvText(Me.mSQL_SetValue)
			End If
			With oCmdSQLSrvText
				LOG.StepName = "ActiveConnection"
				.ActiveConnection = Me.mConnSQLSrv.Connection
				If .LastErr <> "" Then Throw New Exception(.LastErr)
				Select Case Me.mMainTabName
					Case "_ptHFFileInf"
						.AddPara("@MainTabID", Data.SqlDbType.Char, 36)
					Case Else
						.AddPara("@MainTabID", Data.SqlDbType.Char, 32)
				End Select
				.ParaValue("@MainTabID") = MainTabID
				.AddPara("@ExtKey", Data.SqlDbType.NVarChar, 128)
				.ParaValue("@ExtKey") = ExtKey
				If bolIsLarge = True Then
					Dim strLargeExtID As String = ""
					LOG.StepName = "LargeExtID GetTextPigMD5"
					LOG.Ret = Me.mPigFunc.GetTextPigMD5(ExtValue, PigMD5.enmTextType.Unicode, strLargeExtID)
					.AddPara("@LargeExtID", Data.SqlDbType.Char, 32)
					.ParaValue("@LargeExtID") = strLargeExtID
					.AddPara("@ExtValue", Data.SqlDbType.NVarChar, -1)
					.ParaValue("@ExtValue") = ExtValue
				Else
					.AddPara("@ExtValue", Data.SqlDbType.NVarChar, EXT_VALUE_LEN)
					.ParaValue("@ExtValue") = ExtValue
				End If
			End With
			With oCmdSQLSrvText
				LOG.StepName = "ExecuteNonQuery"
				LOG.Ret = .ExecuteNonQuery()
				If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			End With
			oCmdSQLSrvText = Nothing
			Return "OK"
		Catch ex As Exception
			If oCmdSQLSrvText IsNot Nothing Then
				If oCmdSQLSrvText.SQLText <> "" Then
					LOG.AddStepNameInf(oCmdSQLSrvText.SQLText)
					LOG.AddStepNameInf(oCmdSQLSrvText.DebugStr)
				End If
			End If
			oCmdSQLSrvText = Nothing
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function


End Class
