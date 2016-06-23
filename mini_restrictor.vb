  'RESTRINGIR EL ESPACIO VISIBLE Y SELECCIONABLE EN UNA HOJA DE EXCEL EN 1 CLICK
  'Las siguientes líneas de código muestran las posibilidades que tiene VB para
  'personalizar y automatizar distintos procesos.

'1) Creamos un botón

'2) Abrimos el editor de código y añadimos la siguiente línea de código:

Sub btnBlockRange_Click()

  ActiveSheet.ScrollArea = "A1:G10”

End Sub

'Traducción:

 'Macro: cuando haga clic en el botón btnBlockRange_Click haz que solo se pueda
 'seleccionar y mover hasta el rango "A1:G10” de la hoja activa.

'Es decir, cuando hagamos clic en el botón, solo podremos seleccionar las celdas
'que haya dentro del rango de las celdas A1 hasta G10. Asimismo, solo podremos
'movernos (hacer scroll) dentro de ese rango. Este es un buen método para
'ocultar formulas en la misma hoja y evitar que nadie pueda verlas o acceder
'a ellas. Se puede substituir ActiveSheet por el nombre de la hoja en la que
'deseamos aplicar esta acción.

'3) Vincula esta macro al botón anteriormente creado.
