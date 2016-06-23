Sub btnReplacedotxcoma_Click()
Range("A1:A5").Replace What:=".", Replacement:=",", LookAt:=xlPart, _
       	SearchOrder:=xlByRows, MatchCase:=False,
End Sub

'Traducción:
'Macro: cuando haga clic en el botón btnReplacedotxcoma reemplaza los “puntos”
'por las “comas” buscando en cualquier parte de la palabra buscando primero por
'filas y no distinguiendo mayúsculas de minúsculas en el rango de celdas que va
'de la A1 a la A5.
