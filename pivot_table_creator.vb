  'CREAR TABLAS PIVOTANTES PERSONALIZADAS CON UN SOLO CLICK
  'Las siguientes líneas de código muestran las posibilidades que tiene VB para
  'personalizar al máximo las tablas dinámicas y para automatizar su creación

'PARTE I
Sub PivotTbl_Click()
  ActiveWorkbook.PivotCaches.Create(SourceType: =xlDatabase,
  SourceData: ="R1C1:R13C5", Version: =xlPivotTableVersion14).CreatePivotTable
  TableDestination: ="R18C1", TableName: ="PivotTable1",
  DefaultVersion: =xlPivotTableVersion14
'FIN PARTE I
'PARTE II
  With ActiveSheet.PivotTables("PivotTable1").PivotFields("MESES")
        .Orientation = xlRowField
        .Position = 1
    End With
'FIN PARTE II
'PARTE III
  ActiveSheet.PivotTables("PivotTable1").AddDataField
  ActiveSheet.PivotTables("PivotTable1").PivotFields("ALQUILER"), "ALQUILER",
  xlSum
  ActiveSheet.PivotTables("PivotTable1").AddDataField
  ActiveSheet.PivotTables("PivotTable1").PivotFields("TELEFONO"), "TELEFONO",
  xlSum
  ActiveSheet.PivotTables("PivotTable1").AddDataField
  ActiveSheet.PivotTables("PivotTable1").PivotFields("LUZ"), "LUZ", xlSum
  ActiveSheet.PivotTables("PivotTable1").AddDataField
  ActiveSheet.PivotTables("PivotTable1").PivotFields("GAS"), "GAS", xlSum
Sub
'FIN PARTE III

'EXPLICACIÓN PARTE I

'Traducción:
  'Macro: cuando haga clic en el botón PivotTbl_Click() crea un cache de tablas
  'pivotantes (PivotCaches) ubicadas en el libro activo de tipo (SourceType:)
  'rango de celdas (=xlDatabase), con los datos ubicados (SourceData:) entre la
  'celda R1C1 y la celda R13C5 (="R1C1:R13C5") y con versión (Version:) de office
  '2010 (=xlPivotTableVersion14). Ahora, crea una tabla pivotante en la celda
  'R18C1 con nombre PivotTable1 y con versión de office 2010.

'En este primer pedazo de código tenemos dos acciones una detrás de otra:
'Create y CreatePivotTable. La primera acción, Create, es una acción genérica
'de creación. En este caso, está creando un caché de tabla pivotante
'(PivotCaches) con unos parámetros determinados ¿Qué es eso del caché de tabla
'pivotante? Los datos de una tabla pivotante se almacenan de forma temporal en
'el caché del ordenador para que sea más fácil manipular los datos contenidos
'en ella. Entonces, lo primero que ha de hacerse al crear una tabla pivotante
'es crear su caché (PivotCaches).

'EXPLICACIÓN PARTE II

'Traducción:
  'Luego exclusivamente con (with) lo siguiente: orienta el campo MESES de la
  'tabla pivotante PivotTable1 de la hoja activa como fila (xlRowField) y
  'posiciónalo en el puesto 1.

'Este pedazo de código, ubicado dentro de la palabra With, se ejecutará solo a
'aquello que pongamos dentro (todo lo que escribamos entre With y End With) y,
'por lo tanto, no afectará al resto de código dentro de la macro
'(dentro del Sub). Podríamos decir que With es como una sub-macro.
'Aquí, estamos escogiendo como se relacionará el campo respeto al resto de
'campos.
  'Primero, definimos que la posición del campo MESES de la tabla
  'pivotante PivotTable1 de la hoja activa sea 1, esto significa que será la
  'primera. Podemos poner el valor de la posición donde queremos que se ubique.

  'Segundo, definimos que la orientación del campo MESES de la tabla pivotante
  'PivotTable1 de la hoja activa sea en filas (xlRowField). Vamos a ver todos los
  'posibles valores de la acción Orientation y su significado:

'EXPLICACIÓN PARTE III

'Traducción:
  'Luego añade un campo en la tabla pivotante PivotTable1 de la hoja activa.
  'Ahora, ponle la cabecera “ALQUILER” al campo ALQUILER de la tabla pivotante
  'PivotTable1 de la hoja activa y presenta los datos como suma (xlSum).
  'Luego añade un campo en la tabla pivotante PivotTable1 de la hoja activa.
  'Ahora, ponle la cabecera “TELEFONO” al campo TELEFONO de la tabla pivotante
  'PivotTable1 de la hoja activa y presenta los datos como suma (xlSum).
  'Luego añade un campo en la tabla pivotante PivotTable1 de la hoja activa.
  'Ahora, ponle la cabecera “LUZ” al campo LUZ de la tabla pivotante PivotTable1
  'de la hoja activa y presenta los datos como suma (xlSum).
  'Luego añade un campo en la tabla pivotante PivotTable1 de la hoja activa.
  'Ahora, ponle la cabecera “GAS” al campo GAS de la tabla pivotante PivotTable1
  'de la hoja activa y presenta los datos como suma (xlSum).

'Aunque parezca que hay mucho código, en realidad estamos haciendo poca cosa.
'Cada par de líneas define la etiqueta con la que encabezaremos cada campo y
'como se van a mostrar los datos en ella.
'Por ejemplo, en la primera fila añadimos un campo a la tabla pivotante
'PivotTable1 de la hoja activa. Luego, definimos que la etiqueta del campo
'ALQUILER  de la tabla pivotante PivotTable1 de la hoja activa sea ALQUILER y
'que los datos que contiene se sumen (xlSum).
