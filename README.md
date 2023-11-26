-Adaptado a Windows-

Este código es complemento del renombrado XML (el verdaderamente importante). Lamentablemente, en muchos lados el papel y el horrible formato PDF siguen siendo un requisito para muchos procesos, sobre todo contables, por eso también cree este código. OJO: VER NOTAS FINALES

Este código en particular, sólo hace lo siguiente:

1) En el foder de "input", busca el UUID (ese valor alfa-numérico de 32 digitos único e indistiguible) de cada factura PDF y la renombra con él en letras MAYÚSCULAS.
2) Si ya existe un archivo con ese nombre en el mismo folder, le añade un "_2" o "_3", "_4"...etc. a los archivos repetidos para que los puedas eliminar antes renombrarlos.
3) Si un achivo ya está renombrado con su UUID, lo omite y por desgracia si estaba en minúsculas no puede convertirlo a mayúsculas (por restricciones de Windows).
4) Crea un reporte en Excel que resume lo que fue renombrado, así como cualquier error que pudiera haber surgido durante su procesamiento. Recomiendo siempre echarle un ojo antes de usar los archivos para cualquier otro fin. Igual es buena idea guardar todos los reportes en una carpeta de tu preferencia durante un tiempo razonable, quizás 1-2 meses, antes de eliminarlos.
(FALTA MEJORAR ESTE REPORTE Y HOMOLOGARLO JUNTO CON EL DE XML).

Instrucciones:
1) Descarga el archivo .py y guárdalo en donde prefieras, recomiendo crear una nueva carpeta en la ruta de tu editor de código Python (ahí suelen guardarse los proyetos de manera predeterminada).
2) Abre tu editor de código Python y carga el archivo.
3) El folder_path es la ruta a tu carpeta con los archivos a procesar. Debes copiarlos del explorador de archivos y pegarlos entre las comillas simples ('').
   Ejemplo: r'C:\Users\aquí va tu ruta al folder con los archivos' (Aquí mismo estarán los archivos renombrados).
5) Define el output_path para el reporte de Excel, que puede ser el mismo de arriba o donde tu quieras.
6) Ejecuta el código, valida en el reporte de Excel que todo haya salido bien, y ponlos en la misma carpeta que tus archivos XML.

NOTAS FINALES:
-Por la naturaleza del formato PDF:

  1) ALGUNOS ACRHIVOS NO LOGRARÁN SER RENOMBRADOS y quedarán con su nombre original. Tómalo muy en cuenta. En mi caso de cada 1000 facturas quizás sólo el 1% a 2% no logran ser renombradas. 
  Si tienes ideas de cómo mejorar esto (nunca se podrá al 100%) sin que se sacrifique el rendimiento del procesamiento, adelante.

  3) EL PROCESAMIENTO ES CONSIDERABLEMENTE MÁS LENTO QUE LA VERSIÓN XML, por lo que si tu carpeta supera -por ejemplo- los 100 archivos esto demorará varios minutos. 

-Siéntete libre de proponer mejoras al código porque no soy un experto en programación, es más: no sabía nada y a través de meses, durante mis tiempos libres en el trabjo, pude desarrollar esto con ayuda de herramientas de IA y la documentación oficial de Python.

-El código tiene notas en inglés, pues así suelo trabajar (es más conciso que el español) y porque es el lenguaje que usa Python.

-Este código tiene una licencia opensource por lo que puedes hacer con ello casi lo que quieras...sólo no olvides darme crédito.


