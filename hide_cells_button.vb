  'OCULTAR UN RANGO DE CELDAS EN 1 CLICK
  'Las siguientes líneas de código muestran las posibilidades que ofrece VB para
  'automatizar y personalizar procesos relacionados con la visibilidad y los
  'permisos.

'Imaginemos que tenemos una tabla que está compuesta de varias celdas con
'valores fijos, otras tantas con valores variables (a introducir) y otras de
'autocalculo. Imaginemos que, por una u otra razón, no queremos que otras
'personas puedan ver las celdas que contienen los valores fijos a partir de los
'cuales se va a realizar el cálculo.

'1) Creamos una tabla con el siguiente contenido:

  'Encabezamientos:
    'Celda A1: INPUT 1
    'Celda B1: INPUT 2
    'Celda C1: INPUT 3
    'Celda D1: VALOR FIJO 1
    'Celda E1: VALOR FIJO 2
    'Celda F1: VALOR FIJO 3

  'Contenido rangos:
    'Celda A2: 0
    'Celda B2: 0
    'Celda C2: 0
    'Celda D2: 542
    'Celda E2: 365
    'Celda F2: 52

  'Fórumla autocálculo:
    'Celda A4: CÁLCULO
    'Celda B4: =(A2+B2+C2)-(D2+E2+F2)

'En este caso, la celda de auto cálculo suma las tres primeras celdas y resta
'esto a la suma de las tres siguientes. Es decir, la fórmula es la siguiente:
'=(A2+B2+C2)-(D2+E2+F2).
'Ahora, lo que queremos es ocultar las columnas que contienen los valores fijos
'(en este caso, las columnas D, E y F).

'2A) Abrimos el editor de código e insertamos las siguientes líneas:

Private Sub WorkSheet_Activate()
  ActiveSheet.Columns("D").Hidden = True
  ActiveSheet.Columns("E").Hidden = True
  ActiveSheet.Columns("F").Hidden = True
End Sub

'Traducción:
  'Macro, solo en esta hoja: cuando la hoja se active:
  'Cambia el estado de oculto de la columna D (xDown) de la hoja activa a True.
  'Cambia el estado de oculto de la columna E (xDown) de la hoja activa a True.
  'Cambia el estado de oculto de la columna F (xDown) de la hoja activa a True.

'Es decir, cuando se active la hoja la propiedad hidden de las columnas D, E y F
'se pondrá en true. Dicho de otra forma, cuando se active la hoja, se ocultarán
'las columnas D, E y F.  Para aplicar este mismo cambio a filas en vez de
'columnas, deberemos substituir Columns por Rows (ojo, que las filas se numeran
'con números, no con letras, como las columnas).

'2B) En el ejemplo anterior, vemos como ocultamos cada columna por separado.
'También podemos obtener el mismo resultado con un poco menos de código
'del siguiente modo:

  Private Sub WorkSheet_Activate()
    ActiveSheet.Columns("D:F").Hidden = True
  End Sub

'Es decir, cuando la hoja se active, las columnas de la D a la F de la hoja
'activa van a ocultarse (su propiedad hidden se pondrá en true). Igual que en el
'caso anterior, podemos substituir el argumento Columns por Rows.

'El primer ejemplo nos sirve si tenemos que ocultar columnas (o filas, cambiando
'el Columns por Rows) que no se encuentran seguidas una al lado de la otra.
'Para los casos en que tenemos que ocultar columnas (o filas) que SI se
'encuentran una al lado de la otra, podemos usar el segundo método. 
