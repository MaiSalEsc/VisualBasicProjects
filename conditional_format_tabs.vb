  'COLOREAR UNA PESTAÑA DEPENDIENDO DE UN VALOR CONTENIDO EN SU INTERIOR
  'Las siguientes líneas de código muestran las posibilidades que tiene VB para
  'personalizar el formato de distintos elementos y automatizar su aplicación.

'Para este ejemplo, supongamos que tenemos un libro Excel donde hacemos el
'balance de gastos e ingresos del hogar. Introducimos los gastos, luego los
'ingresos y de ahí deducimos un saldo neto, que será dependiendo del día,
'inferior, igual o superior a 0. Imaginemos que cada día creamos una pestaña
'nueva donde ponemos los gastos e ingresos de ese día. Y supongamos que queremos
'saber, de un vistazo rápido, que días hemos entrado en déficit, qué días hemos
'entrado en superávit y qué días hemos gastado tanto como ingresado.

'Podemos conseguir esto coloreando cada pestaña de acuerdo al valor de la celda
'que nos indica el saldo al final del día. Si es positivo, hacemos que la
'pestaña de ese día sea verde, si el saldo es 0 que sea amarillo y si el saldo
'es inferior a 0 hacemos que sea de color rojo.

'1). Creamos tres hojas y las nombramos DIA1, DIA2 y DIA3.

'2). En la pestaña DIA1 ponemos lo siguiente:

  'Encabezamientos:
    'Celda A1: INGRESOS
    'Celda B1: GASTOS
    'Celda C1: BALANCE

  'Valores de cada campo:
    'Celda A2: 500
    'Celda B2: 400
    'Celda C2: =A2-B2 (resultado = 100)

'3). En la pestaña DIA2 ponemos lo siguiente:

      'Encabezamientos:
        'Celda A1: INGRESOS
        'Celda B1: GASTOS
        'Celda C1: BALANCE

      'Valores de cada campo:
        'Celda A2: 500
        'Celda B2: 500
        'Celda C2: =A2-B2 (resultado = 0)

'4). En la pestaña DIA3 ponemos lo siguiente:

      'Encabezamientos:
        'Celda A1: INGRESOS
        'Celda B1: GASTOS
        'Celda C1: BALANCE

      'Valores de cada campo:
        'Celda A2: 500
        'Celda B2: 600
        'Celda C2: =A2-B2 (resultado = -100)

'5) Abrimos el editor de código y añadimos las siguientes líneas
'de código EN CADA HOJA:

Private Sub WorkSheet_Calculate()
  If Range("C2") < 0 Then ActiveSheet.Tab.ColorIndex = 3
  If Range("C2") = 0 Then ActiveSheet.Tab.ColorIndex = 44
  If Range("C2") > 0 Then ActiveSheet.Tab.ColorIndex = 43
End Sub

'Traducción:
  'Macro, solo en esta hoja: cuando la hoja seleccionada haga un cálculo:
  'si el rango de celdas "C2" es inferior a (<) 0 entonces haz que el color de
  'la pestaña de la hoja activa sea (=) del color cuyo índice es 3 (rojo)
  'si el rango de celdas "C2" es igual a (=) 0 entonces haz que el color de
  'la pestaña de la hoja activa sea (=) del color cuyo índice es 44 (ámbar)
  'si el rango de celdas "C2" es superior a (>) 0 entonces haz que el color de
  'la pestaña de la hoja activa sea (=) del color cuyo índice es 43
  '(verde oscuro).

'Es decir, cuando en la hoja que tengamos seleccionada se haga un cálculo
'(WorkSheet_Calculate) comprueba el valor de la celda "C2". Si es inferior a 0,
'entonces colorea la pestaña de la hoja activa de color rojo (cuyo índice es el
'número 3). Si es igual a 0, entonces colorea la pestaña de la hoja activa de
'color ámbar (cuyo índice es el número 44). Y si es superior a 0, entonces
'colorea la pestaña de la hoja activa de color verde oscuro (cuyo índice es
'el número 43).

'Nótese que se ha añadido el argumento Private antes de Sub. Esto es para
'decirle a esta macro en concreto que se ejecute solamente dentro de la hoja
'donde la hemos puesto, en ningún otro sitio.

'Recuerda de copiar el código en cada nueva hoja que crees, de lo contrario no
'se aplicará en la nueva hoja.
