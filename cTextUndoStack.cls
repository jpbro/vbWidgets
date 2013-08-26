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
   textundoaction_Undo  ' An undo operation
End Enum

Public Event RequestTextBox(po_TextBox As cwTextBox)

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
            ' The action is an extendable action
            
            If lo_UndoItem.SelStart = p_SelStart + IIf(lo_UndoItem.Action = textundoaction_Delete, 1, -1) Then
               ' The action is occuring at an extendable caret position (right before or after the last delete or insert respectively)
               l_ExtendAction = True
            End If
         End Select
      End If
   End If

   ' Build the undo information object
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
      ' The new action is independent of the previous action
      
      If Not mo_TempUndoItem Is Nothing Then
         ' Push the previous extended action onto the stack
         mo_UndoStack.Add mo_TempUndoItem
      End If
      
      ' Push the new action onto the stack
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
   
   Dim lo_TextBox As cwTextBox
   Dim lo_UndoItem As cTextUndoStackItem
   
   If Not Me.UndoAvailable Then Exit Function
   If m_UndoIndex < 0 Then m_UndoIndex = 0
   
   Debug.Print "Undoing index " & m_UndoIndex
   
'   If m_UndoIndex = mo_UndoStack.Count - 1 Then
'      RaiseEvent RequestTextBox(lo_TextBox)
'
'      Set lo_UndoItem = New cTextUndoStackItem
'      With lo_UndoItem
'         .Action = textundoaction_Undo
'         .ReverseSelection = lo_TextBox.CaretPos < lo_TextBox.SelStart
'         .SelLength = lo_TextBox.SelLength
'         .SelStart = lo_TextBox.SelStart
'         .Text = lo_TextBox.Text
'      End With
'
'      mo_UndoStack.Add lo_UndoItem, , , m_UndoIndex
'   End If

   Me.FlattenUndoStack
   
   m_UndoIndex = m_UndoIndex - 1
   Set Undo = mo_UndoStack.ItemByIndex(m_UndoIndex)
End Function

Public Function Redo() As cTextUndoStackItem
   ' Redo the next operation in the stack
   ' Returns an undo information object if available
   ' Otherwise returns Nothing
   
   If Not Me.RedoAvailable Then Exit Function
   If m_UndoIndex < 0 Then m_UndoIndex = 0
   
   Set mo_TempUndoItem = Nothing
      
   Debug.Print "Redoing index " & m_UndoIndex
   Set Redo = mo_UndoStack.ItemByIndex(m_UndoIndex)
   
   m_UndoIndex = m_UndoIndex + 1
End Function

Public Function UndoAvailable() As Boolean
   ' Returns True if there is at least one undo information object in the stack in the undo direction
   UndoAvailable = (m_UndoIndex < mo_UndoStack.Count)

   Debug.Print "Undo Available: " & UndoAvailable
End Function

Public Function RedoAvailable() As Boolean
   ' Returns True if there is at least one undo information object in the stack in the redo direction
   RedoAvailable = (m_UndoIndex < mo_UndoStack.Count - 1)
   
   Debug.Print "Redo Available: " & RedoAvailable
End Function

Public Sub Clear()
   ' Clears the undo stack
   m_UndoIndex = 0
   mo_UndoStack.RemoveAll
End Sub

Private Sub Class_Initialize()
   Set mo_UndoStack = New_c.Collection
   
   Me.MaxUndoLevels = 32  ' TODO: Determine an appropriate default value for the undo stack
End Sub

Public Sub FlattenUndoStack()
   If Not mo_TempUndoItem Is Nothing Then
      mo_UndoStack.Add mo_TempUndoItem, , , m_UndoIndex
      m_UndoIndex = m_UndoIndex + 1
      Set mo_TempUndoItem = Nothing
   End If
End Sub
