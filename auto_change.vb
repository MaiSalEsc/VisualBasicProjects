 'REALIZAR CAMBIOS AUTOMÁTICOS CADA VEZ QUE EL VALOR DE UNA CELDA CAMBIE
 'Las siguientes líneas de código muestran las posibilidades que ofrece VB para
 'automatizar y personalizar cambios condicionales.

'Imaginemos que necesitamos conocer el momento en que se ha editado por última
'vez una hoja determinada. Para conseguirlo, vamos a seguir los siguientes
'pasos:

'1) Creamos una tabla con el siguiente contenido:

  'Encabezados:
  'Celda A1: Columna 1
  'Celda B1: Última Edición

  'Contenido:
  'Celdas A2 a A10: cualquier tipo de dato.
  'Celda B2: no escribimos nada

'2) Ahora, abrimos el editor de código e insertamos las siguientes líneas de
'código en la worksheet dónde queremos que sucedan los cambios:

  Private Sub Worksheet_Change(ByVal Target As Range)
      If Intersect(Target, Range("A1:A10")) Is Nothing Or Target.Cells.Count > 1
      _Then Exit Sub
      Application.EnableEvents = False
      Range("B2") = Application.WorksheetFunction.Now()
      Application.EnableEvents = True
  End Sub

'Traducción:
  'Macro, sólo en esta hoja: cuando haya un cambio en la hoja, siendo el "target"
  'de dicho cambio un rango de celdas haz lo siguiente:
  'Si el "target" (o celda que cambia) no coincide con el rango de celdas de la
  'celda A1 a la celda A10 o si hay más de 1 celda que cambia sal de la función.
  'Ahora, no permitas que puedan suceder eventos en la hoja.
  'Ahora, inserta la función de excel "Now()" en el rango de celdas B2.
  'Finalmente, vuelve a permitir que sucedan eventos en la hoja.

'1ª Línea: "Target" es el elemento cuyo cambio va a desencadenar la acción que
'vamos a definir.
'ByVal Target As Range viene a decir que el elemento desencadenador queremos que
'sea un rango de celdas.

'2ª Línea: Esta parte tiene dos funciones:
  '1) Para definir que el rango "target" esté entre las celdas A1 y la A10
  '2) Para evitar que nos salga un molesto "runtime error" si el cambio se
  'produce en una celda que está fuera del rango que va de de la celda A1 a la
  'A10 o si cambia más de una celda a la vez. Si eso sucede cualquiera de las
  'dos cosas, le decimos a la macro que termine esta instrucción.

'3ª Línea: Deshabilitamos los eventos para evitar incompatibilidades.

'4ª Línea: Aquí sucede la acción en sí. Cuando la función detecta un cambio en
'el "target" que hemos establecido (entre las celdas A1 y A10) en la celda B2
'se crea la "WorksheetFunction" (o función de Excel) llamada "Now()".
'El resultado es que cuando cambie cualquiera de las celdas mencionadas, en la
'celda B2 se pondrá la fecha con hora y minutos en que ha sucedido el cambio.

'5ª Línea: Vuelve a habilitar los eventos.
