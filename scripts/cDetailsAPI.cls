Option Explicit
Private pTotalCellsAlloc As Long ' Total number of cells to be used
Private pBlockStartIndex As Long 'Where the range of cells first addon-index is
Private pHiddenCells As Long 'How many cells used as hidden storage
Private pDisplayCells As Long 'How many cells used to display values?
Private pInteractCells As Long 'How many cells used for interaction?

Property Get HiddenCells() As String
     HiddenCells = pHiddenCells
End Property
Property Get DisplayCells() As String
     DisplayCells = pDisplayCells
End Property
Property Get InteractCells() As String
     InteractCells = pInteractCells
End Property
Property Get TotalCells() As String
     TotalCells = pStorageCells + pSummeryCells
End Property


Private Sub Class_Initialize()
    pTotalCellsAlloc = 0
    pBlockStartIndex = -1
    pStorageCells = 0
    pDisplayCells = 0
    pInteractCells = 0
End Sub