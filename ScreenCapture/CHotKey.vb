Public Class CHotKey

#Region "宣告API"
    Private Declare Auto Function RegisterHotKey Lib "user32.dll" Alias "RegisterHotKey" (ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Boolean
    Private Declare Auto Function UnRegisterHotKey Lib "user32.dll" Alias "UnregisterHotKey" (ByVal hwnd As IntPtr, ByVal id As Integer) As Boolean
#End Region

    Private Structure RegKey
        Dim ID As Integer
        Dim Modifiers As Integer
        Dim vk As Integer
    End Structure

    Private m_closeKey As Boolean = True
    Private m_HotKeyRecord As New ArrayList
    Private m_frmMain As IntPtr

    ''' <summary>
    ''' 初始化
    ''' </summary>
    ''' <param name="RegForm">熱鍵登錄的表單</param>
    Public Sub New(ByVal RegForm As Form)
        m_frmMain = RegForm.Handle
    End Sub

    ''' <summary>
    ''' 熱鍵計數
    ''' </summary>
    Public ReadOnly Property Count As Integer
        Get
            Return m_HotKeyRecord.Count
        End Get
    End Property

    ''' <summary>
    ''' 登錄熱鍵
    ''' </summary>
    ''' <param name="id">指定要登錄的 ID</param>
    ''' <param name="fsModifiers">組合鍵</param>
    ''' <param name="vk">熱鍵</param>
    Public Sub RegisterKey(ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer)
        Dim myRegKey As RegKey

        myRegKey.ID = id
        myRegKey.Modifiers = fsModifiers
        myRegKey.vk = vk

        If CheckID(myRegKey) Then
            If RegisterHotKey(m_frmMain, id, fsModifiers, vk) Then
                m_HotKeyRecord.Add(myRegKey)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 移除熱鍵
    ''' </summary>
    ''' <param name="id">指定要移除的 ID</param>
    Public Sub UnRegisterKey(ByVal id As Integer)
        If UnRegisterHotKey(m_frmMain, id) Then
            If m_closeKey Then
                UnKeyRecord(id)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 釋放所有熱鍵資源
    ''' </summary>
    Public Sub Close()
        m_closeKey = False

        For Each i In m_HotKeyRecord
            UnRegisterKey(i.ID)
        Next

        m_HotKeyRecord.Clear()
        m_closeKey = True
    End Sub

    ''' <summary>
    ''' 確認 ID 是否登錄過
    ''' </summary>
    ''' <param name="id">i指定要登入的 ID</param>
    ''' <returns>有登錄過傳回 False，無傳回 True</returns>
    Private Function CheckArray(ByVal id As Integer) As Boolean
        For Each i In m_HotKeyRecord
            If i = id Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 確認有為登錄過熱鍵
    ''' </summary>
    ''' <param name="tmpValue">RegKey結構</param>
    ''' <returns>有登錄過傳回 False，無登錄過傳回 True</returns>
    Private Function CheckID(ByVal tmpValue As RegKey) As Boolean
        Dim blnReturn As Boolean = True

        For Each i In m_HotKeyRecord
            If i.id = tmpValue.ID Then
                Console.WriteLine("ID 已經被登錄過")
                blnReturn = False
            End If

            If i.Modifiers = tmpValue.Modifiers And i.vk = tmpValue.vk Then
                Console.WriteLine("熱鍵已經被登錄過")
                blnReturn = False
            End If
        Next

        Return blnReturn
    End Function

    ''' <summary>
    ''' 移除熱鍵記錄
    ''' </summary>
    ''' <param name="id">指定要移除的 ID</param>
    Private Sub UnKeyRecord(ByVal id As Integer)
        Dim myRegKey As RegKey

        For Each i In m_HotKeyRecord
            If i.ID = id Then
                myRegKey.ID = i.ID
                myRegKey.Modifiers = i.Modifiers
                myRegKey.vk = i.vk
                m_HotKeyRecord.Remove(myRegKey)
                Exit For
            End If
        Next
    End Sub
End Class