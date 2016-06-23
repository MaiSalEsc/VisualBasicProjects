  'REEMPLAZAR UN STRING POR OTRO EN UN SOLO CLICK
  'Las siguientes líneas de código muestran las posibilidades que tiene VB para
  'automatizar tareas como la substitución de trozos de texto (u otro tipo de
  'variables) por otro.

Sub btnReplacedotxcoma_Click()
Range("A1:A5").Replace What:=".", Replacement:=",", LookAt:=xlPart, _
       	SearchOrder:=xlByRows, MatchCase:=False,
End Sub

'Traducción:
'Macro: cuando haga clic en el botón btnReplacedotxcoma reemplaza los “puntos”
'por las “comas” buscando en cualquier parte de la palabra buscando primero por
'filas y no distinguiendo mayúsculas de minúsculas en el rango de celdas que va
'de la A1 a la A5.
