VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MR_Global"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Container of "global" variables (global singleton)
'   so that to centralize all of them.
' The only true global variable will be "Global global_ As MR_Global"
'   and all global accesses will be done via "Set global_.StartSnapshot = TakeSnapshot()"
' TODO
