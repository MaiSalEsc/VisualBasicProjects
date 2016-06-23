  'DESPLAZARSE A LA 1ª o ÚLTIMA CELDA DE UN RANGO EN 1 CLICK
  'Las siguientes líneas de código muestran las posibilidades que ofrece VB para
  'automatizar y personalizar procesos rutinarios, como ir a un lugar concreto
  'de una hoja.

'Imaginemos que tenemos una columna con una gran cantidad de datos, de modo que
'cada vez que queremos ir a la parte inferior de la misma tenemos que hacer
'scroll durante mucho rato. Imaginemos que con un par de botones pudiésemos
'movernos por este rango de celdas con un solo click.

'Para hacerlo, vamos a seguir los siguientes pasos:

'1) Crea una tabla con los siguientes parámetros:

  'Haz que la primera fila tenga el doble del grosor habitual y formatéala con
  'un color de fondo deseado (por ejemplo, gris o azul claro).

  'Celda A2: COLUMNA 1
  'de Celda A3 a celda A300: Introduce en cada celda datos aleatorios

'2) A continuación, congela la primera fila para que estemos donde estemos de la
'tabla siempre tengamos los botones a la vista.

'3) Ahora, crea dos botones e insértalos en las celdas A1 y B1.

'4) Abrimos el editor de código y añadimos las siguientes líneas:

Sub btnUP_Click()
  Range("A3").End(xlUp).Select
End Sub

'Traducción:

  'Macro: cuando haga click en el botón btnUP_Click(): Selecciona la parte final
  'hacia arriba (xlUp) del rango de celdas que comienza en la celda A3 ("A3")

'Es decir, cuando hagamos click en el botón que hemos llamado ARRIBA se
'seleccionará la celda que esté justo encima de la primera celda del rango
'(en este caso, A3). La elección de esta celda (A3) como la primera del rango
'no es aleatoria. Si tenemos en cuenta que la primera fila la tenemos congelada
'(para que los botones estén siempre visibles) y que la segunda línea contiene
'el nombre de la tabla, para que la función nos lleve a la parte superior de la
'tabla (la que está encima de la primera celda del rango) deberemos indicarle a
'la macro que la primera celda es la A3.

'También hay que fijarse en el hecho de que no hemos establecido cual es la
'última celda del rango. No lo hemos hecho porque resulta redundante.
'Esta función ya se dirige automáticamente a la última celda que tenga algún
'registro. Si está vacía, no la cuenta. Así que la última celda será siempre
'la última que tenga información.

'5) A continuación, añadimos las siguientes líneas de código:

Sub btnDOWN_Click()
  Range("A3").End(xlDown).Select
End Sub

'Traducción:

'Macro: cuando haga click en el botón btnDown_Click(): Selecciona la parte final
'hacia abajo (xDown) del rango de celdas que comienza en la celda A3 ("A3")

'Es decir, cuando hagamos click en el botón que hemos llamado ABAJO se
'seleccionará la celda que esté justo debajo de la última celda del rango
'(en este caso, A300).

'6) Finalmente asociamos cada macro con su botón correspondiente
'(la macro btnUP_Click() con el botón ARRIBA y la macro btnDOWN_Click()
'con el botón ABAJO).
