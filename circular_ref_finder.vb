  'HALLAR REFERENCIAS RICULARES EN 1 CLICK
  'Las siguientes líneas de código muestran las posibilidades que ofrece VB para
  'automatizar y personalizar procesos rutinarios, como hallar referencias
  'circulares en nuestras fórmulas.

'1) Creamos un botón en nuestra hoja de Excel

'2) Abrimos el editor de código y añadimos la siguiente línea de código.

Sub btnFindCR_Click()

  Worksheet("Workseet1").CircularReference.Select

End Sub

'Traducción:
  'Macro: cuando haga clic en el botón btnFindCR_Click selecciona la celda que
  'contenga la primera referencia circular que encuentres en la hoja llamada
  '"Workseet1".
'La traducción resulta bastante autoexplicatoria del funcionamiento de la macro.
'Únicamente hay que tener en cuenta que hay que substituir "Workseet1" por el
'nombre que tenga la pestaña (hoja) en la que queremos buscar la referencia
'circular.

'3) Vincula esta macro al botón creado anteriormente.
