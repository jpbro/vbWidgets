VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextUndoStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTextUndoAction
   textundoaction_None  ' No operation
   textundoaction_Type  ' A typing/insert operation
   textundoaction_Delete   ' Text deletion operation
   textundoaction_Cut   ' A clipboard cut operation
   textundoaction_Paste ' A clipboard paste operation
End Enum

Private mo_UndoStack As vbRichClient5.cCollection
Private mo_TempUndoItem As cTextUndoStackItem
Private m_UndoIndex As Long
Private m_MaxUndoLevels As Long

Public Property Get MaxUndoLevels() As Long
   ' The size of the undo/redo stack in levels/steps
   MaxUndoLevels = m_MaxUndoLevels
End Property

Public Property Let MaxUndoLevels(ByVal p_Max As Long)
   ' The size of the undo/redo stack in levels/steps
   ' 0 = No undo Stack.
   
   Dim i As Long
   
   If p_Max < 0 Then Err.Raise 5, , "Must be >=0"
   
   m_MaxUndoLevels = p_Max
   
   Do While mo_UndoStack.Count > p_Max
      ' Fit an existing stack to the new size
      mo_UndoStack.RemoveByIndex 0
   Loop
   
   m_UndoIndex = mo_UndoStack.Count - 1
End Property

Public Sub AddUndoItem(ByVal p_Action As eTextUndoAction, ByVal p_Text As String, ByVal p_SelStart As Long, ByVal p_SelLength As Long, ByVal p_ReverseSelection As Boolean)
   ' Push undo information onto the stack
   
   Dim lo_UndoItem As cTextUndoStackItem
   Dim l_ExtendAction As Boolean
   Dim i As Long
   
   ' Short-circuits
   If p_Action = textundoaction_None Then Exit Sub
   If m_MaxUndoLevels = 0 Then Exit Sub
   
   ' Clear Redo side of Stack
   If m_UndoIndex >= 0 Then
      For i = m_UndoIndex + 1 To mo_UndoStack.Count - 1
         mo_UndoStack.RemoveByIndex i
      Next i
      m_UndoIndex = mo_UndoStack.Count - 1
   End If
      
   ' Check for previous extendable action (typing or deleting)
   If mo_TempUndoItem Is Nothing Then
      If mo_UndoStack.Count > 0 Then
         Set lo_UndoItem = mo_UndoStack.ItemByIndex(mo_UndoStack.Count - 1)
      End If
   Else
      Set lo_UndoItem = mo_TempUndoItem
   End If
   
   If Not lo_UndoItem Is Nothing Then
      If lo_UndoItem.Action = p_Action Then
         ' Previous action in Stack matches the new action
         
         Select Case lo_UndoItem.Action
         Case textundoaction_Type, textundoaction_Delete
            ' The previous action is an extendable action
            
            If lo_UndoItem.SelStart = p_SelStart + IIf(lo_UndoItem.Action = textundoaction_Delete, 1, -1) Then
               ' The action is occuring at an extendable caret position
               ' (immediately before the last delete or after the last insert)
               l_ExtendAction = True
            End If
         End Select
      End If
   End If

   ' Build the current undo information object
   Set lo_UndoItem = New cTextUndoStackItem
   With lo_UndoItem
      .Action = p_Action
      .Text = p_Text
      .SelStart = p_SelStart
      .SelLength = p_SelLength
      .ReverseSelection = p_ReverseSelection
   End With
      
   If l_ExtendAction Then
      ' The new action is an extension of the previous action
      ' So keep it in memory, but don't add it to the undo stack yet
      Set mo_TempUndoItem = lo_UndoItem
   
   Else
      ' The current action is independent of the previous action
      Set mo_TempUndoItem = Nothing
      
      ' Push the current action onto the stack
      mo_UndoStack.Add lo_UndoItem
      ' Record the current undo index for future Undo/Redo requests
      m_UndoIndex = mo_UndoStack.Count - 1
            
      Do While mo_UndoStack.Count > m_MaxUndoLevels
         ' Remove old undo actions that no longer fit in the stack
         mo_UndoStack.RemoveByIndex 0
      Loop
   End If
End Sub

Public Function Undo() As cTextUndoStackItem
   ' Undo the next operation in the stack
   ' Returns an undo information object if available
   ' Otherwise returns Nothing
   
   ' Add any lingering extendable undo action to the stack
   FlattenUndoStack
   
   If m_UndoIndex <= 0 Then Exit Function
   If Not Me.UndoAvailable Then Exit Function
   
   Set Undo = mo_UndoStack.ItemByIndex(m_UndoIndex - 1)
   m_UndoIndex = m_UndoIndex - 1
End Function

Public Function Redo() As cTextUndoStackItem
   ' Redo the next operation in the stack
   ' Returns an undo information object if available
   ' Otherwise returns Nothing
   
   If m_UndoIndex >= mo_UndoStack.Count - 1 Then Exit Function
   If Not Me.RedoAvailable Then Exit Function
   If m_UndoIndex < 0 Then Debug.Assert False
   
   Set mo_TempUndoItem = Nothing
      
   Set Redo = mo_UndoStack.ItemByIndex(m_UndoIndex + 1)
   m_UndoIndex = m_UndoIndex + 1
End Function

Public Function UndoAvailable() As Boolean
   ' Returns True if there is at least one undo information object in the stack in the undo direction
   UndoAvailable = (m_UndoIndex >= 0) And (m_UndoIndex < mo_UndoStack.Count + IIf(mo_TempUndoItem Is Nothing, 0, 1))
End Function

Public Function RedoAvailable() As Boolean
   ' Returns True if there is at least one undo information object in the stack in the redo direction
   RedoAvailable = (m_UndoIndex >= 0) And (m_UndoIndex < mo_UndoStack.Count + IIf(mo_TempUndoItem Is Nothing, -1, 0))
End Function

Public Sub Clear()
   ' Clears the undo stack
   m_UndoIndex = 0
   mo_UndoStack.RemoveAll
   Set mo_TempUndoItem = Nothing
End Sub

Private Sub Class_Initialize()
   Set mo_UndoStack = New_c.Collection
   
   Me.MaxUndoLevels = 32  ' TODO: Determine an appropriate default value for the undo stack
End Sub

Private Sub FlattenUndoStack()
   If mo_TempUndoItem Is Nothing Then Exit Sub
   
   mo_UndoStack.Add mo_TempUndoItem
   Set mo_TempUndoItem = Nothing

   m_UndoIndex = mo_UndoStack.Count - 1
End Sub
