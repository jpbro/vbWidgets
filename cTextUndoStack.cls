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

' Multi-level Undo/Redo stack for text operations

Public Enum eTextUndoAction
   textundoaction_None  ' No operation
   textundoaction_Type  ' A typing/insert operation
   textundoaction_Delete   ' Text deletion operation
   textundoaction_Cut   ' A clipboard cut operation
   textundoaction_Paste ' A clipboard paste operation
End Enum

Private mo_UndoStack As vbRichClient5.cCollection
Private mo_ExtendedUndoItem As cTextUndoStackItem
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

Public Function CreateUndoItem(ByVal p_Action As eTextUndoAction, ByVal p_Text As String, ByVal p_SelStart As Long, ByVal p_SelLength As Long, ByVal p_SelBackwards As Boolean) As cTextUndoStackItem
   Set CreateUndoItem = New cTextUndoStackItem
   With CreateUndoItem
      .Action = p_Action
      .Text = p_Text
      .SelStart = p_SelStart
      .SelLength = p_SelLength
      .SelBackwards = p_SelBackwards
   End With
End Function

Public Sub AddUndoItem(po_UndoItem As cTextUndoStackItem)
   ' Push undo information onto the stack
   
   Dim lo_UndoItem As cTextUndoStackItem
   Dim l_ExtendAction As Boolean
   Dim i As Long
   
   ' Short-circuits
   If po_UndoItem Is Nothing Then Exit Sub
   If po_UndoItem.Action = textundoaction_None Then Exit Sub
   If m_MaxUndoLevels = 0 Then Exit Sub
   
   ' Clear Redo side of Stack
   If m_UndoIndex >= 0 Then
      For i = m_UndoIndex + 1 To mo_UndoStack.Count - 1
         mo_UndoStack.RemoveByIndex i
      Next i
      m_UndoIndex = mo_UndoStack.Count - 1
   End If
      
   ' Check for previous extendable action (typing or deleting)
   If mo_ExtendedUndoItem Is Nothing Then
      If mo_UndoStack.Count > 0 Then
         Set lo_UndoItem = mo_UndoStack.ItemByIndex(mo_UndoStack.Count - 1)
      End If
   Else
      Set lo_UndoItem = mo_ExtendedUndoItem
   End If

   If Not lo_UndoItem Is Nothing Then
      If lo_UndoItem.SelBackwards = po_UndoItem.SelBackwards Then
         If lo_UndoItem.SelStart = po_UndoItem.SelStart Then
            If lo_UndoItem.SelLength = po_UndoItem.SelLength Then
               If lo_UndoItem.Text = po_UndoItem.Text Then
                  ' Exact duplicate of previous undo item, skip
                  Exit Sub
               End If
            End If
         End If
      End If
      
      If lo_UndoItem.Action = po_UndoItem.Action Then
         ' Previous action in Stack matches the new action
         Select Case lo_UndoItem.Action
         Case textundoaction_Type, textundoaction_Delete
            ' The previous action is an extendable action
            
            If lo_UndoItem.SelStart = po_UndoItem.SelStart + IIf(lo_UndoItem.Action = textundoaction_Delete, 1, -1) Then
               ' The action is occuring at an extendable caret position
               ' (immediately before the last delete or after the last insert)
               l_ExtendAction = True
            End If
         End Select
      End If
   End If

   If l_ExtendAction Then
      ' The new action is an extension of the previous action
      ' So keep it in memory, but don't add it to the undo stack yet
      Set mo_ExtendedUndoItem = po_UndoItem
   
   Else
      ' The current action is independent of the previous action
      Set mo_ExtendedUndoItem = Nothing
      
      StoreUndoItem po_UndoItem
   End If
End Sub

Private Sub StoreUndoItem(po_UndoItem As cTextUndoStackItem)
   If po_UndoItem Is Nothing Then Exit Sub
   
   ' Push the current action onto the stack
   mo_UndoStack.Add po_UndoItem
         
   Do While mo_UndoStack.Count > m_MaxUndoLevels
      ' Remove old undo actions that no longer fit in the stack
      mo_UndoStack.RemoveByIndex 0
   Loop

   ' Record the current undo index for future Undo/Redo requests
   m_UndoIndex = mo_UndoStack.Count - 1
End Sub

Private Sub ConsolidateStack()
   ' Add any lingering temporary undo item to the stack
   If mo_ExtendedUndoItem Is Nothing Then Exit Sub
   
   StoreUndoItem mo_ExtendedUndoItem
      
   Set mo_ExtendedUndoItem = Nothing
End Sub

Public Function Undo() As cTextUndoStackItem
   ' Undo the next operation in the stack
   ' Returns an undo information object if available
   ' Otherwise returns Nothing
   
   ' Add any lingering extendable undo action to the stack
   ConsolidateStack
   
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
   
   Set mo_ExtendedUndoItem = Nothing
         
   Set Redo = mo_UndoStack.ItemByIndex(m_UndoIndex + 1)
   
   m_UndoIndex = m_UndoIndex + 1
End Function

Public Function UndoAvailable() As Boolean
   ' Returns True if there is at least one undo information object in the stack in the undo direction
   If m_UndoIndex < mo_UndoStack.Count + IIf(mo_ExtendedUndoItem Is Nothing, 0, 1) Then
      If m_UndoIndex > 0 Then
         UndoAvailable = True
      ElseIf m_UndoIndex = 0 Then
         UndoAvailable = Not mo_ExtendedUndoItem Is Nothing
      End If
   End If
End Function

Public Function RedoAvailable() As Boolean
   ' Returns True if there is at least one undo information object in the stack in the redo direction
   If m_UndoIndex < mo_UndoStack.Count + IIf(mo_ExtendedUndoItem Is Nothing, -1, 0) Then
      RedoAvailable = (m_UndoIndex >= 0)
   End If
End Function

Public Sub Clear()
   ' Clears the undo stack
   m_UndoIndex = -1
   mo_UndoStack.RemoveAll
   Set mo_ExtendedUndoItem = Nothing
End Sub

Private Sub Class_Initialize()
   Set mo_UndoStack = New_c.Collection
   m_UndoIndex = -1
   
   Me.MaxUndoLevels = 32  ' TODO: Determine an appropriate default value for the undo stack
End Sub

