
'PARTE I
  'Tras crear el formulario correspondiente mediante las herramientas de la
  'pestaña DESARROLLADOR de la hoja de Excel le asociamos al formulario el
  'siguiente código, que regirá el contenido de los campos de texto:

Private Sub Form1_Initialize()
  TextBox1.Value = ""
  TextBox2.Value = ""
  TextBox3.Value = ""
  TextBox4.Value = ""
  TextBox5.Value = ""
End Sub

'Traducción:
  'Macro: cuando se abra el formulario Form1 pon dentro de la caja de texto
  'TextBox1 lo que ponga después del signo “=”. Luego pon dentro de la caja de
  'texto TextBox3 lo que ponga después del signo “=”. Luego pon dentro de la caja
  'de texto TextBox3 lo que ponga después del signo “=”. Luego pon dentro de la
  'caja de texto TextBox4 lo que ponga después del signo “=”. Luego pon dentro de
  'la caja de texto TextBox5 lo que ponga después del signo “=”.

'Es decir, le estamos pidiendo a la macro que deje en blanco el interior de las
'cajas de texto del formulario, para que así podamos escribir nosotros.

'PARTE II
  'Ahora, tras crear el botón cuya finalidad es la de añadir los registros
  'introducidos en las cajas de texto anteriores, le asociaremos el siguiente
  'código:

Private Sub CommandButton1_Click()
  Range("A2:E2").Select
  Selection.Insert Shift: =xlShiftDown
  Range("A2").Value = TextBox1.Value
  Range("B2").Value = TextBox2.Value
  Range("C2").Value = TextBox3.Value
  Range("D2").Value = TextBox4.Value
  Range("E2").Value = TextBox5.Value
  Range("A2:E2").ClearFormats
End Sub

'Traducción:
  'Macro: cuando haga clic en el botón CommandButton1 selecciona los registros
  'incluidos en el rango que va de la A2 a la E2. Luego inserta una fila el en el
  'rango seleccionado desplazando las filas hacia abajo. Luego haz que el valor
  'del rango (la celda A2) sea igual al valor de la caja de texto TextBox1.
  'Luego haz que el valor del rango (la celda B2) sea igual al valor de la caja
  'de texto TextBox2. Luego haz que el valor del rango (la celda C2) sea igual al
  'valor de la caja de texto TextBox3. Luego haz que el valor del rango
  '(la celda D2) sea igual al valor de la caja de texto TextBox4. Luego haz que el
  'valor del rango (la celda E2) sea igual al valor de la caja de texto TextBox5.
  'Finalmente borra el formato del rango (celdas de la A2 a la E2).

  'Es decir, cuando hagamos clic en el botón CommandButton1 (al que nosotros
  'hemos puesto un texto que dice “AÑADIR”) primero insertará una nueva fila en
  'la fila 2 (bajando las demás filas hacia abajo). Luego pondrá en la casilla
  'A2 lo que hayamos puesto en el TextBox1. Y así sucesivamente con todos los
  'campos.

'PARTE III (opcional)
  'Ahora, en el menú superior izquierdo (aún dentro de Visual Basic) haz clic
  'con el botón derecho sobre el nombre de la hoja donde está la tabla (y el
  'botón para abrir el formulario) y selecciona la opción “View Code”. Ahora,
  'introduce el siguiente código:

Sub btnOpenForm()
  Form1.Show
End Sub

'Traducción:
 'Macro: cuando haga clic en el botón btnOpenForm muestra el formulario Form1.
