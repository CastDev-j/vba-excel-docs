"use client";

import { useState } from "react";
import { Book, Code, FileText, Menu, Play, X } from "lucide-react";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { CodeBlock } from "@/components/code-block";
import { SearchBar } from "@/components/search-bar";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Button } from "@/components/ui/button";

const excelObjects = [
  {
    object: "Application",
    description: "Controlar la aplicación Excel completa",
    properties: [
      {
        name: "ScreenUpdating",
        desc: "Controla las actualizaciones de la pantalla",
      },
      { name: "Calculation", desc: "Controla el modo de cálculo" },
      { name: "StatusBar", desc: "Muestra texto en la barra de estado" },
    ],
    methods: [
      { name: "Quit", desc: "Cierra la aplicación Excel" },
      { name: "Run", desc: "Ejecuta una macro específica" },
      { name: "Save", desc: "Guarda el libro actual" },
    ],
    commonCode: `' Ejemplo común con Application
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.StatusBar = "Procesando datos..."
Application.Run "MacroNombre"
Application.Quit`,
  },
  {
    object: "Workbook",
    description: "Trabajar con libros de Excel",
    properties: [
      { name: "Name", desc: "Nombre del libro" },
      { name: "Path", desc: "Ruta del libro" },
      { name: "Saved", desc: "Indica si el libro está guardado" },
    ],
    methods: [
      { name: "Save", desc: "Guarda el libro" },
      { name: "Close", desc: "Cierra el libro" },
      { name: "Activate", desc: "Activa el libro" },
    ],
    commonCode: `' Ejemplo común con Workbook
Dim wb As Workbook
Set wb = Workbooks("NombreDelLibro.xlsx")
wb.Save
wb.Close False`,
  },
  {
    object: "Worksheet",
    description: "Manipular hojas de cálculo individuales",
    properties: [
      { name: "Name", desc: "Nombre de la hoja" },
      { name: "Cells", desc: "Colección de todas las celdas en la hoja" },
      { name: "Range", desc: "Accede a un rango específico de celdas" },
    ],
    methods: [
      { name: "Activate", desc: "Activa la hoja" },
      { name: "Copy", desc: "Copia la hoja" },
      { name: "Delete", desc: "Elimina la hoja" },
    ],
    commonCode: `' Ejemplo común con Worksheet
Dim ws As Worksheet
Set ws = Worksheets("Hoja1")
ws.Activate
ws.Cells.Clear`,
  },
  {
    object: "Range",
    description: "Trabajar con rangos de celdas",
    properties: [
      { name: "Value", desc: "Valor del rango" },
      { name: "Address", desc: "Dirección del rango" },
      { name: "Count", desc: "Número de celdas en el rango" },
    ],
    methods: [
      { name: "Select", desc: "Selecciona el rango" },
      { name: "Clear", desc: "Limpia el contenido del rango" },
      { name: "Copy", desc: "Copia el rango" },
    ],
    commonCode: `' Ejemplo común con Range
Dim rng As Range
Set rng = Range("A1:B10")
rng.Select
rng.Clear`,
  },
  {
    object: "Cells",
    description: "Acceder a celdas individuales por índice",
    properties: [
      { name: "Value", desc: "Valor de la celda" },
      { name: "Font", desc: "Propiedades de fuente de la celda" },
      { name: "Interior", desc: "Propiedades de fondo de la celda" },
    ],
    methods: [
      { name: "Select", desc: "Selecciona la celda" },
      { name: "Clear", desc: "Limpia el contenido de la celda" },
      { name: "Copy", desc: "Copia el contenido de la celda" },
    ],
    commonCode: `' Ejemplo común con Cells
Dim cell As Range
Set cell = Cells(1, 1)
cell.Value = "Hola"
cell.Font.Bold = True
cell.Interior.Color = RGB(255, 255, 0)`,
  },
];

export default function VBADocumentation() {
  const [activeSection, setActiveSection] = useState("inicio");
  const [difficultyFilter, setDifficultyFilter] = useState("Todos");

  const searchData = [
    // Sintaxis básica
    {
      id: "variables",
      title: "Declaración de Variables",
      category: "Variables",
      section: "sintaxis",
      description:
        "Cómo declarar variables en VBA con diferentes tipos de datos",
    },
    {
      id: "constantes",
      title: "Constantes",
      category: "Variables",
      section: "sintaxis",
      description:
        "Definir valores constantes que no cambian durante la ejecución",
    },
    {
      id: "arrays",
      title: "Arrays (Matrices)",
      category: "Variables",
      section: "sintaxis",
      description:
        "Trabajar con arrays estáticos y dinámicos para almacenar múltiples valores",
    },
    {
      id: "if-then",
      title: "Estructuras If-Then",
      category: "Control",
      section: "sintaxis",
      description:
        "Estructuras condicionales para tomar decisiones en el código",
    },
    {
      id: "loops",
      title: "Bucles For y While",
      category: "Control",
      section: "sintaxis",
      description: "Repetir código usando bucles For, While y Do-Loop",
    },
    {
      id: "sub",
      title: "Subrutinas (Sub)",
      category: "Procedimientos",
      section: "sintaxis",
      description: "Crear subrutinas que ejecutan código sin devolver valores",
    },
    {
      id: "function",
      title: "Funciones (Function)",
      category: "Procedimientos",
      section: "sintaxis",
      description:
        "Crear funciones que devuelven valores y pueden ser reutilizadas",
    },

    // Objetos Excel
    {
      id: "application",
      title: "Objeto Application",
      category: "Objetos",
      section: "objetos",
      description: "Controlar la aplicación Excel completa",
    },
    {
      id: "workbook",
      title: "Objeto Workbook",
      category: "Objetos",
      section: "objetos",
      description: "Trabajar con libros de Excel",
    },
    {
      id: "worksheet",
      title: "Objeto Worksheet",
      category: "Objetos",
      section: "objetos",
      description: "Manipular hojas de cálculo individuales",
    },
    {
      id: "range",
      title: "Objeto Range",
      category: "Objetos",
      section: "objetos",
      description: "Trabajar con rangos de celdas",
    },
    {
      id: "cells",
      title: "Objeto Cells",
      category: "Objetos",
      section: "objetos",
      description: "Acceder a celdas individuales por índice",
    },

    // Ejemplos prácticos
    {
      id: "formato-reporte",
      title: "Automatizar Formato de Reportes",
      category: "Formato",
      section: "ejemplos",
      description:
        "Aplicar formato automático a reportes con encabezados, bordes y colores",
    },
    {
      id: "validacion-datos",
      title: "Validación de Datos",
      category: "Validación",
      section: "ejemplos",
      description: "Validar entrada de datos y mostrar mensajes de error",
    },
    {
      id: "dashboard",
      title: "Generar Dashboard Automático",
      category: "Dashboard",
      section: "ejemplos",
      description:
        "Crear dashboards automáticos con gráficos y tablas dinámicas",
    },
    {
      id: "importar-datos",
      title: "Importar Datos Externos",
      category: "Datos",
      section: "ejemplos",
      description: "Importar datos desde archivos CSV y bases de datos",
    },
    {
      id: "email-automatico",
      title: "Envío de Emails Automático",
      category: "Automatización",
      section: "ejemplos",
      description: "Enviar emails automáticamente desde Excel usando Outlook",
    },
  ];

  const handleSearchResult = (section: string, itemId?: string) => {
    setActiveSection(section);
    if (itemId) {
      // Scroll al elemento específico después de un pequeño delay
      setTimeout(() => {
        const element = document.getElementById(itemId);
        if (element) {
          element.scrollIntoView({ behavior: "smooth", block: "start" });
        }
      }, 100);
    }
  };

  const sidebarSections = [
    { id: "inicio", title: "Inicio", icon: Book },
    { id: "sintaxis", title: "Sintaxis Básica", icon: Code },
    { id: "objetos", title: "Objetos Excel", icon: FileText },
    { id: "ejemplos", title: "Ejemplos Prácticos", icon: Play },
    { id: "referencia", title: "Referencia Completa", icon: Book },
  ];

  const vbaSyntax = [
    {
      category: "Variables y Tipos de Datos",
      items: [
        {
          id: "variables",
          name: "Declaración de Variables",
          syntax: `' Declaración básica
Dim nombre As String
Dim edad As Integer
Dim salario As Double
Dim activo As Boolean

' Declaración con valor inicial
Dim mensaje As String
mensaje = "Hola Mundo"
Dim contador As Integer
contador = 0

' Variables de objeto
Dim hoja As Worksheet
Set hoja = ActiveSheet`,
          description:
            "Las variables deben declararse antes de usarse. Use 'Set' para objetos.",
          tips: "Siempre declare variables para evitar errores. Use nombres descriptivos.",
        },
        {
          id: "constantes",
          name: "Constantes",
          syntax: `' Constantes públicas
Public Const IVA As Double = 0.21
Public Const EMPRESA As String = "Mi Empresa"

' Constantes privadas
Private Const MAX_FILAS As Long = 1000000

' Uso de constantes
Dim precio As Double
precio = 100
Dim precioConIVA As Double
precioConIVA = precio * (1 + IVA)`,
          description:
            "Las constantes no cambian durante la ejecución del programa.",
          tips: "Use constantes para valores que no cambiarán, como tasas o límites.",
        },
        {
          id: "arrays",
          name: "Arrays (Matrices)",
          syntax: `' Array estático
Dim numeros(1 To 10) As Integer
Dim nombres(5) As String

' Array dinámico
Dim datos() As Double
ReDim datos(1 To 100)

' Array multidimensional
Dim tabla(1 To 10, 1 To 5) As String

' Llenar array
For i = 1 To 10
    numeros(i) = i * 2
Next i`,
          description: "Los arrays almacenan múltiples valores del mismo tipo.",
          tips: "Use ReDim para cambiar el tamaño de arrays dinámicos. Los índices pueden empezar en 0 o 1.",
        },
      ],
    },
    {
      category: "Estructuras de Control",
      items: [
        {
          id: "if-then",
          name: "Condicionales If-Then-Else",
          syntax: `' If simple
If edad >= 18 Then
    MsgBox "Es mayor de edad"
End If

' If-Else
If nota >= 7 Then
    resultado = "Aprobado"
Else
    resultado = "Reprobado"
End If

' If anidado
If ventas > 10000 Then
    comision = ventas * 0.1
ElseIf ventas > 5000 Then
    comision = ventas * 0.05
Else
    comision = 0
End If`,
          description:
            "Ejecuta código basado en condiciones verdaderas o falsas.",
          tips: "Use ElseIf para múltiples condiciones. Siempre termine con End If.",
        },
        {
          id: "select-case",
          name: "Select Case",
          syntax: `' Select Case básico
Select Case dia
    Case 1
        nombreDia = "Lunes"
    Case 2
        nombreDia = "Martes"
    Case 3 To 5
        nombreDia = "Miércoles a Viernes"
    Case Else
        nombreDia = "Fin de semana"
End Select

' Select Case con rangos
Select Case calificacion
    Case 90 To 100
        letra = "A"
    Case 80 To 89
        letra = "B"
    Case 70 To 79
        letra = "C"
    Case Else
        letra = "F"
End Select`,
          description:
            "Evalúa una expresión contra múltiples valores posibles.",
          tips: "Use Select Case cuando tenga muchas condiciones If-ElseIf.",
        },
      ],
    },
    {
      category: "Bucles",
      items: [
        {
          id: "for-next",
          name: "Bucle For-Next",
          syntax: `' For básico
For i = 1 To 10
    Cells(i, 1).Value = i
Next i

' For con Step
For i = 10 To 1 Step -1
    Debug.Print i
Next i

' For Each para rangos
For Each celda In Range("A1:A10")
    celda.Value = celda.Value * 2
Next celda

' For Each para colecciones
For Each hoja In ThisWorkbook.Worksheets
    hoja.Cells(1, 1).Value = "Encabezado"
Next hoja`,
          description:
            "Repite código un número específico de veces o para cada elemento.",
          tips: "Use For Each para recorrer colecciones. Step permite incrementos personalizados.",
        },
        {
          id: "do-while",
          name: "Bucles Do-While y Do-Until",
          syntax: `' Do While (mientras sea verdadero)
Dim contador As Integer
contador = 1
Do While contador <= 10
    Debug.Print contador
    contador = contador + 1
Loop

' Do Until (hasta que sea verdadero)
Dim fila As Long
fila = 1
Do Until Cells(fila, 1).Value = ""
    Debug.Print Cells(fila, 1).Value
    fila = fila + 1
Loop

' While-Wend (alternativa)
While Not IsEmpty(ActiveCell)
    ActiveCell.Value = UCase(ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
Wend`,
          description:
            "Repite código mientras o hasta que se cumpla una condición.",
          tips: "Cuidado con bucles infinitos. Siempre asegúrese de que la condición cambie.",
        },
      ],
    },
    {
      category: "Procedimientos",
      items: [
        {
          id: "sub",
          name: "Subrutinas (Sub)",
          syntax: `' Sub sin parámetros
Sub SaludarUsuario()
    MsgBox "¡Hola Usuario!"
End Sub

' Sub con parámetros
Sub MostrarMensaje(mensaje As String, titulo As String)
    MsgBox mensaje, vbInformation, titulo
End Sub

' Sub con parámetros opcionales
Sub FormatearCelda(rango As Range, Optional color As Long = RGB(255, 255, 0))
    rango.Interior.Color = color
    rango.Font.Bold = True
End Sub

' Llamar subrutinas
Call SaludarUsuario
MostrarMensaje "Proceso completado", "Información"
FormatearCelda Range("A1")
FormatearCelda Range("B1"), RGB(0, 255, 0)`,
          description:
            "Las subrutinas ejecutan código pero no devuelven valores.",
          tips: "Use Call para llamar subs (opcional). Los parámetros opcionales deben ir al final.",
        },
        {
          id: "function",
          name: "Funciones (Function)",
          syntax: `' Función básica
Function CalcularArea(largo As Double, ancho As Double) As Double
    CalcularArea = largo * ancho
End Function

' Función con validación
Function DividirSeguro(dividendo As Double, divisor As Double) As Variant
    If divisor = 0 Then
        DividirSeguro = "Error: División por cero"
    Else
        DividirSeguro = dividendo / divisor
    End If
End Function

' Función que devuelve array
Function ObtenerEstadisticas(rango As Range) As Variant
    Dim stats(1 To 3) As Double
    stats(1) = Application.WorksheetFunction.Average(rango)
    stats(2) = Application.WorksheetFunction.Max(rango)
    stats(3) = Application.WorksheetFunction.Min(rango)
    ObtenerEstadisticas = stats
End Function

' Usar funciones
Dim area As Double
area = CalcularArea(10, 5)
Dim resultado As Variant
resultado = DividirSeguro(10, 2)`,
          description: "Las funciones ejecutan código y devuelven un valor.",
          tips: "Asigne el resultado a la función usando su nombre. Use Variant para múltiples tipos.",
        },
      ],
    },
  ];

  const practicalExamples = [
    // EJEMPLOS BÁSICOS
    {
      id: "hola-mundo",
      title: "Mi Primer Macro",
      difficulty: "Básico",
      code: `Sub MiPrimerMacro()
    ' Mi primera macro en VBA
    MsgBox "¡Hola Mundo desde VBA!"
End Sub`,
      description: "Tu primera macro para mostrar un mensaje simple",
      category: "Introducción",
    },
    {
      id: "escribir-celda",
      title: "Escribir en Celdas",
      difficulty: "Básico",
      code: `Sub EscribirEnCeldas()
    ' Escribir valores en diferentes celdas
    Range("A1").Value = "Nombre"
    Range("B1").Value = "Edad"
    Range("A2").Value = "Juan"
    Range("B2").Value = 25
    
    ' También puedes usar Cells
    Cells(3, 1).Value = "María"
    Cells(3, 2).Value = 30
End Sub`,
      description: "Aprende a escribir datos en celdas usando Range y Cells",
      category: "Básicos",
    },
    {
      id: "leer-celda",
      title: "Leer Datos de Celdas",
      difficulty: "Básico",
      code: `Sub LeerDatosCeldas()
    Dim nombre As String
    Dim edad As Integer
    
    ' Leer valores de las celdas
    nombre = Range("A2").Value
    edad = Range("B2").Value
    
    ' Mostrar los datos leídos
    MsgBox "El usuario " & nombre & " tiene " & edad & " años"
End Sub`,
      description: "Cómo leer y usar datos almacenados en celdas",
      category: "Básicos",
    },
    {
      id: "bucle-simple",
      title: "Bucle For Básico",
      difficulty: "Básico",
      code: `Sub BucleSimple()
    Dim i As Integer
    
    ' Llenar números del 1 al 10 en la columna A
    For i = 1 To 10
        Cells(i, 1).Value = i
    Next i
    
    MsgBox "Números del 1 al 10 escritos en columna A"
End Sub`,
      description:
        "Uso básico de bucles For para automatizar tareas repetitivas",
      category: "Básicos",
    },
    {
      id: "formato-basico",
      title: "Formato Básico de Celdas",
      difficulty: "Básico",
      code: `Sub FormatoBasico()
    ' Formatear una celda
    With Range("A1")
        .Value = "TÍTULO"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 100, 200)
        .HorizontalAlignment = xlCenter
    End With
End Sub`,
      description:
        "Aplicar formato básico: negrita, color, tamaño y alineación",
      category: "Formato",
    },

    // EJEMPLOS INTERMEDIOS
    {
      id: "formato-reporte",
      title: "Formato Automático de Reportes",
      difficulty: "Intermedio",
      code: `Sub FormatearReporte()
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    
    ' Encontrar última fila y columna con datos
    ultimaFila = Cells(Rows.Count, 1).End(xlUp).Row
    ultimaColumna = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Formatear encabezados
    With Range(Cells(1, 1), Cells(1, ultimaColumna))
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Aplicar bordes a toda la tabla
    With Range(Cells(1, 1), Cells(ultimaFila, ultimaColumna))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Alternar colores de filas
    For i = 2 To ultimaFila
        If i Mod 2 = 0 Then
            Range(Cells(i, 1), Cells(i, ultimaColumna)).Interior.Color = RGB(242, 242, 242)
        End If
    Next i
    
    ' Ajustar ancho de columnas
    Columns.AutoFit
    
    ' Agregar filtros
    Range(Cells(1, 1), Cells(ultimaFila, ultimaColumna)).AutoFilter
    
    MsgBox "Formato aplicado correctamente", vbInformation
End Sub`,
      description:
        "Automatiza el formato de reportes con encabezados, bordes, colores alternos y filtros",
      category: "Formato y Presentación",
    },
    {
      id: "validacion-datos",
      title: "Validación de Datos Avanzada",
      difficulty: "Intermedio",
      code: `Function ValidarEmail(email As String) As Boolean
    ' Validación básica de formato de email
    If InStr(email, "@") > 0 And InStr(email, ".") > 0 Then
        ValidarEmail = True
    Else
        ValidarEmail = False
    End If
End Function

Sub ValidarFormulario()
    Dim nombre As String, email As String, edad As Integer
    Dim errores As String
    
    ' Obtener datos del formulario
    nombre = Range("B2").Value
    email = Range("B3").Value
    edad = Range("B4").Value
    
    ' Validar nombre
    If Len(nombre) < 2 Then
        errores = errores & "- El nombre debe tener al menos 2 caracteres" & vbCrLf
    End If
    
    ' Validar email
    If Not ValidarEmail(email) Then
        errores = errores & "- El formato del email no es válido" & vbCrLf
    End If
    
    ' Validar edad
    If edad < 18 Or edad > 100 Then
        errores = errores & "- La edad debe estar entre 18 y 100 años" & vbCrLf
    End If
    
    ' Mostrar resultado
    If errores = "" Then
        MsgBox "Todos los datos son válidos", vbInformation, "Validación Exitosa"
        Range("B6").Value = "VÁLIDO"
        Range("B6").Interior.Color = RGB(144, 238, 144)
    Else
        MsgBox "Errores encontrados:" & vbCrLf & vbCrLf & errores, vbCritical, "Errores de Validación"
        Range("B6").Value = "INVÁLIDO"
        Range("B6").Interior.Color = RGB(255, 182, 193)
    End If
End Sub`,
      description:
        "Sistema completo de validación de formularios con mensajes de error detallados",
      category: "Validación de Datos",
    },
    {
      id: "buscar-reemplazar",
      title: "Buscar y Reemplazar Avanzado",
      difficulty: "Intermedio",
      code: `Sub BuscarReemplazarAvanzado()
    Dim buscar As String
    Dim reemplazar As String
    Dim contador As Integer
    
    buscar = InputBox("¿Qué texto deseas buscar?")
    reemplazar = InputBox("¿Por qué texto lo quieres reemplazar?")
    
    If buscar <> "" Then
        ' Buscar y reemplazar en toda la hoja
        contador = 0
        For Each celda In ActiveSheet.UsedRange
            If InStr(1, celda.Value, buscar, vbTextCompare) > 0 Then
                celda.Value = Replace(celda.Value, buscar, reemplazar, , , vbTextCompare)
                contador = contador + 1
            End If
        Next celda
        
        MsgBox "Se reemplazaron " & contador & " coincidencias", vbInformation
    End If
End Sub`,
      description:
        "Herramienta para buscar y reemplazar texto en toda la hoja con contador",
      category: "Manipulación de Datos",
    },
    {
      id: "graficos-automaticos",
      title: "Crear Gráficos Automáticamente",
      difficulty: "Intermedio",
      code: `Sub CrearGraficoAutomatico()
    Dim rango As Range
    Dim grafico As Chart
    Dim ultimaFila As Long
    
    ' Encontrar el rango de datos
    ultimaFila = Cells(Rows.Count, 1).End(xlUp).Row
    Set rango = Range("A1:B" & ultimaFila)
    
    ' Crear el gráfico
    Set grafico = ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Chart
    
    ' Configurar el gráfico
    With grafico
        .SetSourceData rango
        .HasTitle = True
        .ChartTitle.Text = "Gráfico Automático"
        .ChartStyle = 26
        .Parent.Left = Range("D2").Left
        .Parent.Top = Range("D2").Top
        .Parent.Width = 400
        .Parent.Height = 300
    End With
    
    MsgBox "Gráfico creado exitosamente", vbInformation
End Sub`,
      description: "Crea gráficos automáticamente basados en datos de la hoja",
      category: "Gráficos y Visualización",
    },
    {
      id: "importar-csv",
      title: "Importar Archivos CSV",
      difficulty: "Intermedio",
      code: `Sub ImportarCSV()
    Dim archivo As String
    Dim hoja As Worksheet
    
    ' Seleccionar archivo CSV
    archivo = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv")
    
    If archivo <> "False" Then
        Set hoja = ActiveSheet
        
        ' Limpiar datos existentes
        hoja.Cells.Clear
        
        ' Importar el archivo CSV
        With hoja.QueryTables.Add(Connection:="TEXT;" & archivo, Destination:=Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFileColumnDataTypes = Array(1)
            .Refresh BackgroundQuery:=False
        End With
        
        ' Formatear como tabla
        hoja.ListObjects.Add(xlSrcRange, hoja.UsedRange, , xlYes).Name = "TablaImportada"
        
        MsgBox "Archivo CSV importado correctamente", vbInformation
    End If
End Sub`,
      description: "Importa archivos CSV y los convierte en tablas formateadas",
      category: "Importación de Datos",
    },

    // EJEMPLOS AVANZADOS
    {
      id: "emails-automaticos",
      title: "Sistema de Emails Automáticos",
      difficulty: "Avanzado",
      code: `Sub EnviarEmailsMasivos()
    Dim i As Long
    Dim ultimaFila As Long
    Dim email As String, nombre As String
    
    ultimaFila = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaFila ' Asume que la fila 1 tiene encabezados
        email = Cells(i, 1).Value ' Columna A: emails
        nombre = Cells(i, 2).Value ' Columna B: nombres
        
        If email <> "" Then
            Call EnviarEmailPersonalizado(email, nombre)
            Cells(i, 3).Value = "Enviado - " & Format(Now, "hh:mm:ss")
        End If
    Next i
    
    MsgBox "Envío masivo completado", vbInformation
End Sub

Sub EnviarEmailPersonalizado(destinatario As String, nombrePersona As String)
    Dim outlookApp As Object
    Dim outlookMail As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    
    With outlookMail
        .To = destinatario
        .Subject = "Mensaje personalizado para " & nombrePersona
        .Body = "Estimado/a " & nombrePersona & "," & vbCrLf & vbCrLf & _
                "Este es un mensaje personalizado generado automáticamente." & vbCrLf & vbCrLf & _
                "Saludos cordiales"
        .Send
    End With
    
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub`,
      description:
        "Sistema completo para envío automático de emails individuales y masivos usando Outlook",
      category: "Automatización de Emails",
    },
    {
      id: "base-datos-completa",
      title: "Sistema de Base de Datos Completo",
      difficulty: "Avanzado",
      code: `' Clase para manejar registros de empleados
Class EmpleadoManager
    Private hoja As Worksheet
    
    Private Sub Class_Initialize()
        Set hoja = ThisWorkbook.Worksheets("Empleados")
    End Sub
    
    Public Sub AgregarEmpleado(nombre As String, puesto As String, salario As Double)
        Dim ultimaFila As Long
        ultimaFila = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row + 1
        
        With hoja
            .Cells(ultimaFila, 1).Value = ultimaFila - 1 ' ID
            .Cells(ultimaFila, 2).Value = nombre
            .Cells(ultimaFila, 3).Value = puesto
            .Cells(ultimaFila, 4).Value = salario
            .Cells(ultimaFila, 5).Value = Date ' Fecha de registro
        End With
        
        Call FormatearFila(ultimaFila)
    End Sub
    
    Public Function BuscarEmpleado(nombre As String) As Long
        Dim i As Long
        For i = 2 To hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
            If UCase(hoja.Cells(i, 2).Value) = UCase(nombre) Then
                BuscarEmpleado = i
                Exit Function
            End If
        Next i
        BuscarEmpleado = 0
    End Function
    
    Private Sub FormatearFila(fila As Long)
        With hoja.Range(hoja.Cells(fila, 1), hoja.Cells(fila, 5))
            .Borders.LineStyle = xlContinuous
            If fila Mod 2 = 0 Then
                .Interior.Color = RGB(240, 240, 240)
            End If
        End With
    End Sub
End Class

Sub InicializarSistemaEmpleados()
    Dim manager As New EmpleadoManager
    
    ' Crear encabezados si no existen
    With ThisWorkbook.Worksheets("Empleados")
        .Cells(1, 1).Value = "ID"
        .Cells(1, 2).Value = "Nombre"
        .Cells(1, 3).Value = "Puesto"
        .Cells(1, 4).Value = "Salario"
        .Cells(1, 5).Value = "Fecha Registro"
        .Range("A1:E1").Font.Bold = True
    End With
    
    ' Agregar empleados de ejemplo
    manager.AgregarEmpleado "Juan Pérez", "Desarrollador", 50000
    manager.AgregarEmpleado "María García", "Analista", 45000
    manager.AgregarEmpleado "Carlos López", "Gerente", 70000
    
    MsgBox "Sistema de empleados inicializado", vbInformation
End Sub`,
      description:
        "Sistema completo de gestión de empleados con clases, búsqueda y formato automático",
      category: "Sistemas Complejos",
    },
    {
      id: "dashboard-interactivo",
      title: "Dashboard Interactivo con Botones",
      difficulty: "Avanzado",
      code: `Sub CrearDashboardInteractivo()
    Dim hoja As Worksheet
    Set hoja = ActiveSheet
    
    ' Limpiar hoja
    hoja.Cells.Clear
    
    ' Crear título
    With hoja.Range("A1:F1")
        .Merge
        .Value = "DASHBOARD INTERACTIVO"
        .Font.Size = 20
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Crear botones funcionales
    Call CrearBoton(hoja, "B3", "Generar Reporte", "GenerarReporte")
    Call CrearBoton(hoja, "D3", "Actualizar Datos", "ActualizarDatos")
    Call CrearBoton(hoja, "B5", "Enviar Emails", "EnviarEmails")
    Call CrearBoton(hoja, "D5", "Crear Gráfico", "CrearGrafico")
    
    ' Crear área de estado
    With hoja.Range("A7:F7")
        .Merge
        .Value = "Estado: Dashboard listo"
        .Interior.Color = RGB(146, 208, 80)
        .HorizontalAlignment = xlCenter
        .Name = "EstadoDashboard"
    End With
    
    MsgBox "Dashboard interactivo creado", vbInformation
End Sub

Sub CrearBoton(hoja As Worksheet, celda As String, texto As String, macro As String)
    Dim boton As Button
    Set boton = hoja.Buttons.Add(hoja.Range(celda).Left, hoja.Range(celda).Top, 100, 30)
    
    With boton
        .Caption = texto
        .OnAction = macro
        .Font.Bold = True
    End With
End Sub

Sub ActualizarEstado(mensaje As String)
    Range("EstadoDashboard").Value = "Estado: " & mensaje
End Sub`,
      description:
        "Dashboard completo con botones interactivos y área de estado dinámico",
      category: "Interfaces de Usuario",
    },
    {
      id: "web-scraping",
      title: "Web Scraping Automático",
      difficulty: "Avanzado",
      code: `Sub ExtraerDatosWeb()
    Dim ie As Object
    Dim doc As Object
    Dim elementos As Object
    Dim i As Long
    
    ' Crear instancia de Internet Explorer
    Set ie = CreateObject("InternetExplorer.Application")
    
    With ie
        .Visible = False
        .Navigate "https://example.com/datos"
        
        ' Esperar a que cargue la página
        Do While .Busy Or .ReadyState <> 4
            DoEvents
        Loop
        
        Set doc = .Document
    End With
    
    ' Extraer datos de la tabla
    Set elementos = doc.getElementsByTagName("tr")
    
    ' Escribir encabezados
    Range("A1").Value = "Dato 1"
    Range("B1").Value = "Dato 2"
    Range("C1").Value = "Fecha Extracción"
    
    ' Procesar cada fila de la tabla web
    For i = 1 To elementos.Length - 1
        If elementos(i).Children.Length >= 2 Then
            Cells(i + 1, 1).Value = elementos(i).Children(0).innerText
            Cells(i + 1, 2).Value = elementos(i).Children(1).innerText
            Cells(i + 1, 3).Value = Now
        End If
    Next i
    
    ' Cerrar Internet Explorer
    ie.Quit
    Set ie = Nothing
    
    ' Formatear datos extraídos
    Range("A1:C" & i).Borders.LineStyle = xlContinuous
    Range("A1:C1").Font.Bold = True
    
    MsgBox "Datos extraídos exitosamente: " & (i - 1) & " registros", vbInformation
End Sub`,
      description:
        "Extrae datos automáticamente de páginas web y los organiza en Excel",
      category: "Web Scraping",
    },
    {
      id: "backup-automatico",
      title: "Sistema de Backup Automático",
      difficulty: "Avanzado",
      code: `Sub SistemaBackupCompleto()
    Dim rutaBackup As String
    Dim nombreArchivo As String
    Dim archivoOriginal As String
    
    ' Configurar rutas
    archivoOriginal = ThisWorkbook.FullName
    rutaBackup = ThisWorkbook.Path & "\Backups\"
    nombreArchivo = "Backup_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & "_" & ThisWorkbook.Name
    
    ' Crear carpeta de backup si no existe
    If Dir(rutaBackup, vbDirectory) = "" Then
        MkDir rutaBackup
    End If
    
    ' Crear backup
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs rutaBackup & nombreArchivo
    Application.DisplayAlerts = True
    
    ' Limpiar backups antiguos (mantener solo los últimos 10)
    Call LimpiarBackupsAntiguos(rutaBackup)
    
    ' Registrar backup en log
    Call RegistrarBackup(rutaBackup & nombreArchivo)
    
    MsgBox "Backup creado exitosamente en:" & vbCrLf & rutaBackup & nombreArchivo, vbInformation
End Sub

Sub LimpiarBackupsAntiguos(ruta As String)
    Dim archivo As String
    Dim archivos() As String
    Dim fechas() As Date
    Dim i As Integer, j As Integer
    
    ' Obtener lista de archivos de backup
    archivo = Dir(ruta & "Backup_*.xlsx")
    i = 0
    
    Do While archivo <> ""
        ReDim Preserve archivos(i)
        ReDim Preserve fechas(i)
        archivos(i) = archivo
        fechas(i) = FileDateTime(ruta & archivo)
        i = i + 1
        archivo = Dir
    Loop
    
    ' Si hay más de 10 backups, eliminar los más antiguos
    If i > 10 Then
        ' Ordenar por fecha (más antiguos primero)
        For j = 0 To i - 2
            For k = j + 1 To i - 1
                If fechas(j) > fechas(k) Then
                    ' Intercambiar fechas
                    temp = fechas(j)
                    fechas(j) = fechas(k)
                    fechas(k) = temp
                    ' Intercambiar nombres
                    tempNombre = archivos(j)
                    archivos(j) = archivos(k)
                    archivos(k) = tempNombre
                End If
            Next k
        Next j
        
        ' Eliminar los más antiguos
        For j = 0 To i - 11
            Kill ruta & archivos(j)
        Next j
    End If
End Sub

Sub RegistrarBackup(rutaCompleta As String)
    Dim archivoLog As String
    Dim numeroArchivo As Integer
    
    archivoLog = ThisWorkbook.Path & "\backup_log.txt"
    numeroArchivo = FreeFile
    
    Open archivoLog For Append As #numeroArchivo
    Print #numeroArchivo, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - Backup creado: " & rutaCompleta
    Close #numeroArchivo
End Sub`,
      description:
        "Sistema completo de backup automático con limpieza de archivos antiguos y registro de actividad",
      category: "Automatización de Archivos",
    },
  ];

  const renderContent = () => {
    switch (activeSection) {
      case "inicio":
        return (
          <div className="space-y-6">
            <div className="text-center py-8">
              <h1 className="text-4xl font-bold text-primary mb-4">
                VBA para Excel
              </h1>
              <p className="text-xl text-muted-foreground mb-8">
                Documentación completa para programar con Visual Basic for
                Applications en Excel
              </p>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 max-w-4xl mx-auto">
                <Card
                  className="hover:shadow-lg transition-shadow cursor-pointer"
                  onClick={() => setActiveSection("sintaxis")}
                >
                  <CardHeader className="text-center">
                    <Code className="h-12 w-12 text-primary mx-auto mb-2" />
                    <CardTitle>Sintaxis Básica</CardTitle>
                    <CardDescription>
                      Aprende los fundamentos de VBA
                    </CardDescription>
                  </CardHeader>
                </Card>
                <Card
                  className="hover:shadow-lg transition-shadow cursor-pointer"
                  onClick={() => setActiveSection("objetos")}
                >
                  <CardHeader className="text-center">
                    <FileText className="h-12 w-12 text-primary mx-auto mb-2" />
                    <CardTitle>Objetos Excel</CardTitle>
                    <CardDescription>
                      Domina los objetos de Excel
                    </CardDescription>
                  </CardHeader>
                </Card>
                <Card
                  className="hover:shadow-lg transition-shadow cursor-pointer"
                  onClick={() => setActiveSection("ejemplos")}
                >
                  <CardHeader className="text-center">
                    <Play className="h-12 w-12 text-primary mx-auto mb-2" />
                    <CardTitle>Ejemplos Prácticos</CardTitle>
                    <CardDescription>Código listo para usar</CardDescription>
                  </CardHeader>
                </Card>
              </div>
            </div>
          </div>
        );

      case "sintaxis":
        return (
          <div className="space-y-6">
            <div>
              <h2 className="text-3xl font-bold text-primary mb-4">
                Sintaxis Básica de VBA
              </h2>
              <p className="text-muted-foreground mb-6">
                Fundamentos esenciales para programar en VBA
              </p>
            </div>

            {vbaSyntax.map((section, sectionIndex) => (
              <div key={sectionIndex} className="space-y-4">
                <h3 className="text-2xl font-semibold text-primary">
                  {section.category}
                </h3>
                {section.items.map((item, itemIndex) => (
                  <Card key={itemIndex} id={item.id} className="flex w-full">
                    <CardHeader>
                      <CardTitle className="text-xl">{item.name}</CardTitle>
                      <CardDescription>{item.description}</CardDescription>
                      {item.tips && (
                        <div className="mt-2 p-3 bg-accent/50 rounded-lg">
                          <p className="text-sm font-medium text-accent-foreground">
                            💡 Consejo: {item.tips}
                          </p>
                        </div>
                      )}
                    </CardHeader>
                    <CardContent className="w-full overflow-x-auto">
                      <CodeBlock
                        code={item.syntax}
                        title={item.name}
                        description={item.description}
                        language="vba"
                      />
                    </CardContent>
                  </Card>
                ))}
              </div>
            ))}
          </div>
        );

      case "ejemplos":
        return (
          <div className="space-y-6">
            <div>
              <h2 className="text-3xl font-bold text-primary mb-4">
                Ejemplos Prácticos
              </h2>
              <p className="text-muted-foreground mb-6">
                Código VBA organizado por nivel de dificultad
              </p>
            </div>

            <div className="flex flex-wrap gap-2 mb-6">
              <Button
                variant={difficultyFilter === "Todos" ? "default" : "outline"}
                size="sm"
                onClick={() => setDifficultyFilter("Todos")}
              >
                Todos ({practicalExamples.length})
              </Button>
              <Button
                variant={difficultyFilter === "Básico" ? "default" : "outline"}
                size="sm"
                onClick={() => setDifficultyFilter("Básico")}
              >
                Básico (
                {
                  practicalExamples.filter((ex) => ex.difficulty === "Básico")
                    .length
                }
                )
              </Button>
              <Button
                variant={
                  difficultyFilter === "Intermedio" ? "default" : "outline"
                }
                size="sm"
                onClick={() => setDifficultyFilter("Intermedio")}
              >
                Intermedio (
                {
                  practicalExamples.filter(
                    (ex) => ex.difficulty === "Intermedio"
                  ).length
                }
                )
              </Button>
              <Button
                variant={
                  difficultyFilter === "Avanzado" ? "default" : "outline"
                }
                size="sm"
                onClick={() => setDifficultyFilter("Avanzado")}
              >
                Avanzado (
                {
                  practicalExamples.filter((ex) => ex.difficulty === "Avanzado")
                    .length
                }
                )
              </Button>
            </div>

            <div className="space-y-6">
              {practicalExamples
                .filter(
                  (example) =>
                    difficultyFilter === "Todos" ||
                    example.difficulty === difficultyFilter
                )
                .map((example, index) => (
                  <Card key={index} id={example.id}>
                    <CardHeader>
                      <div className="flex items-center justify-between">
                        <div>
                          <CardTitle className="text-xl">
                            {example.title}
                          </CardTitle>
                          <div className="flex items-center gap-2 mt-2">
                            <Badge
                              variant={
                                example.difficulty === "Básico"
                                  ? "secondary"
                                  : example.difficulty === "Intermedio"
                                  ? "default"
                                  : "destructive"
                              }
                            >
                              {example.difficulty}
                            </Badge>
                            <Badge variant="outline">{example.category}</Badge>
                          </div>
                        </div>
                      </div>
                      <CardDescription className="mt-2">
                        {example.description}
                      </CardDescription>
                    </CardHeader>
                    <CardContent>
                      <CodeBlock
                        code={example.code}
                        title={example.title}
                        description={`Nivel: ${example.difficulty} | Categoría: ${example.category}`}
                        language="vba"
                      />
                    </CardContent>
                  </Card>
                ))}
            </div>
          </div>
        );

      case "objetos":
        return (
          <div className="space-y-6">
            <div>
              <h2 className="text-3xl font-bold text-primary mb-4">
                Objetos de Excel
              </h2>
              <p className="text-muted-foreground mb-6">
                Comprende los objetos principales que puedes manipular con VBA
              </p>
            </div>

            <div className="grid gap-6">
              {excelObjects.map((obj, index) => (
                <Card key={index}>
                  <CardHeader>
                    <CardTitle className="text-xl text-primary">
                      {obj.object}
                    </CardTitle>
                    <CardDescription>{obj.description}</CardDescription>
                  </CardHeader>
                  <CardContent>
                    <Tabs defaultValue="properties" className="w-full">
                      <TabsList className="grid w-full grid-cols-3">
                        <TabsTrigger value="properties">
                          Propiedades
                        </TabsTrigger>
                        <TabsTrigger value="methods">Métodos</TabsTrigger>
                        <TabsTrigger value="examples">Ejemplos</TabsTrigger>
                      </TabsList>
                      <TabsContent value="properties" className="mt-4">
                        <div className="space-y-3">
                          {obj.properties.map((prop, propIndex) => (
                            <div
                              key={propIndex}
                              className="flex items-start justify-between p-3 bg-muted/50 rounded"
                            >
                              <div className="flex-1">
                                <code className="font-semibold">
                                  {prop.name || prop}
                                </code>
                                {prop.desc && (
                                  <p className="text-sm text-muted-foreground mt-1">
                                    {prop.desc}
                                  </p>
                                )}
                              </div>
                              {prop.example && (
                                <code className="text-xs bg-background px-2 py-1 rounded ml-2">
                                  {prop.example}
                                </code>
                              )}
                            </div>
                          ))}
                        </div>
                      </TabsContent>
                      <TabsContent value="methods" className="mt-4">
                        <div className="space-y-3">
                          {obj.methods.map((method, methodIndex) => (
                            <div
                              key={methodIndex}
                              className="flex items-start justify-between p-3 bg-muted/50 rounded"
                            >
                              <div className="flex-1">
                                <code className="font-semibold">
                                  {method.name || method}
                                </code>
                                {method.desc && (
                                  <p className="text-sm text-muted-foreground mt-1">
                                    {method.desc}
                                  </p>
                                )}
                              </div>
                              {method.example && (
                                <code className="text-xs bg-background px-2 py-1 rounded ml-2">
                                  {method.example}
                                </code>
                              )}
                            </div>
                          ))}
                        </div>
                      </TabsContent>
                      <TabsContent value="examples" className="mt-4">
                        {obj.commonCode && (
                          <CodeBlock
                            code={obj.commonCode}
                            title={`Ejemplo común con ${obj.object}`}
                            language="vba"
                          />
                        )}
                      </TabsContent>
                    </Tabs>
                  </CardContent>
                </Card>
              ))}
            </div>
          </div>
        );

      case "referencia":
        return (
          <div className="space-y-6">
            <div>
              <h2 className="text-3xl font-bold text-primary mb-4">
                Referencia Completa
              </h2>
              <p className="text-muted-foreground mb-6">
                Guía de referencia rápida para VBA en Excel
              </p>
            </div>

            <div className="grid gap-6">
              <Card>
                <CardHeader>
                  <CardTitle>Comandos Esenciales</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <h4 className="font-semibold mb-3">
                        Manipulación de Celdas
                      </h4>
                      <div className="space-y-2 text-sm">
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">
                            Range("A1").Value = "texto"
                          </code>
                          <p className="text-muted-foreground mt-1">
                            Establecer valor en celda
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Cells(1,1).Value</code>
                          <p className="text-muted-foreground mt-1">
                            Acceso por índices (fila, columna)
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">ActiveCell.Address</code>
                          <p className="text-muted-foreground mt-1">
                            Dirección de celda activa
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Selection.Count</code>
                          <p className="text-muted-foreground mt-1">
                            Número de celdas seleccionadas
                          </p>
                        </div>
                      </div>
                    </div>
                    <div>
                      <h4 className="font-semibold mb-3">
                        Navegación y Control
                      </h4>
                      <div className="space-y-2 text-sm">
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">
                            Worksheets("Hoja1").Activate
                          </code>
                          <p className="text-muted-foreground mt-1">
                            Activar hoja específica
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Range("A1").Select</code>
                          <p className="text-muted-foreground mt-1">
                            Seleccionar rango
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">ActiveSheet.Name</code>
                          <p className="text-muted-foreground mt-1">
                            Nombre de hoja activa
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">ThisWorkbook.Path</code>
                          <p className="text-muted-foreground mt-1">
                            Ruta del archivo actual
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardHeader>
                  <CardTitle>Funciones Útiles</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <h4 className="font-semibold mb-3">Funciones de Texto</h4>
                      <div className="space-y-2 text-sm">
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Len(texto)</code>
                          <p className="text-muted-foreground mt-1">
                            Longitud del texto
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Left(texto, n)</code>
                          <p className="text-muted-foreground mt-1">
                            Primeros n caracteres
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Right(texto, n)</code>
                          <p className="text-muted-foreground mt-1">
                            Últimos n caracteres
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Trim(texto)</code>
                          <p className="text-muted-foreground mt-1">
                            Eliminar espacios extra
                          </p>
                        </div>
                      </div>
                    </div>
                    <div>
                      <h4 className="font-semibold mb-3">
                        Funciones Matemáticas
                      </h4>
                      <div className="space-y-2 text-sm">
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Int(numero)</code>
                          <p className="text-muted-foreground mt-1">
                            Parte entera del número
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">
                            Round(numero, decimales)
                          </code>
                          <p className="text-muted-foreground mt-1">
                            Redondear número
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Abs(numero)</code>
                          <p className="text-muted-foreground mt-1">
                            Valor absoluto
                          </p>
                        </div>
                        <div className="p-2 bg-muted/50 rounded">
                          <code className="font-mono">Rnd()</code>
                          <p className="text-muted-foreground mt-1">
                            Número aleatorio 0-1
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardHeader>
                  <CardTitle>Manejo de Errores</CardTitle>
                </CardHeader>
                <CardContent>
                  <CodeBlock
                    code={`' Manejo básico de errores
On Error GoTo ErrorHandler

' Tu código aquí...
Dim resultado As Double
resultado = 10 / 0  ' Esto causará un error

Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Next  ' Continuar con la siguiente línea

' Otros tipos de manejo
On Error Resume Next  ' Ignorar errores
On Error GoTo 0       ' Desactivar manejo de errores`}
                    title="Estructura básica de manejo de errores"
                    language="vba"
                  />
                </CardContent>
              </Card>

              <Card>
                <CardHeader>
                  <CardTitle>Atajos de Teclado en VBA Editor</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <h4 className="font-semibold mb-2">Navegación</h4>
                      <ul className="space-y-1 text-sm">
                        <li>
                          <kbd className="bg-muted px-1 rounded">F5</kbd> -
                          Ejecutar macro
                        </li>
                        <li>
                          <kbd className="bg-muted px-1 rounded">F8</kbd> -
                          Ejecutar paso a paso
                        </li>
                        <li>
                          <kbd className="bg-muted px-1 rounded">F9</kbd> -
                          Punto de interrupción
                        </li>
                        <li>
                          <kbd className="bg-muted px-1 rounded">Ctrl+G</kbd> -
                          Ventana inmediata
                        </li>
                      </ul>
                    </div>
                    <div>
                      <h4 className="font-semibold mb-2">Edición</h4>
                      <ul className="space-y-1 text-sm">
                        <li>
                          <kbd className="bg-muted px-1 rounded">
                            Ctrl+Space
                          </kbd>{" "}
                          - Autocompletar
                        </li>
                        <li>
                          <kbd className="bg-muted px-1 rounded">
                            Ctrl+Shift+F9
                          </kbd>{" "}
                          - Limpiar puntos de interrupción
                        </li>
                        <li>
                          <kbd className="bg-muted px-1 rounded">Ctrl+H</kbd> -
                          Buscar y reemplazar
                        </li>
                        <li>
                          <kbd className="bg-muted px-1 rounded">Ctrl+Z</kbd> -
                          Deshacer
                        </li>
                      </ul>
                    </div>
                  </div>
                </CardContent>
              </Card>
            </div>
          </div>
        );

      default:
        return <div>Sección en desarrollo...</div>;
    }
  };

  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  return (
    <div className="min-h-screen bg-background">
      {/* Header con buscador */}
      <header className="sticky top-0 z-40 w-full border-b bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/60">
        <div className="container flex h-16 items-center justify-between px-4">
          <div className="flex items-center gap-4">
            {/* Botón hamburguesa visible solo en móviles */}
            <button
              className="lg:hidden p-2 rounded-md hover:bg-accent"
              onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            >
              {isSidebarOpen ? (
                <X className="h-6 w-6" />
              ) : (
                <Menu className="h-6 w-6" />
              )}
            </button>
            <h1 className="text-xl font-bold hidden lg:block">
              VBA Excel Docs
            </h1>
          </div>
          <SearchBar
            onResultClick={handleSearchResult}
            searchData={searchData}
          />
        </div>
      </header>

      <div className="flex sticky">
        {/* Sidebar */}
        <aside
          className={`
            fixed inset-y-0 left-0 z-50 w-64 border-r bg-sidebar overflow-y-auto transform transition-transform duration-300
            lg:sticky lg:top-16 lg:h-[calc(100vh-4rem)] lg:translate-x-0
            ${isSidebarOpen ? "translate-x-0" : "-translate-x-full"}
          `}
        >
          <nav className="p-4 space-y-2 sm:mt-0 mt-14">
            {sidebarSections.map((section) => {
              const Icon = section.icon;
              return (
                <button
                  key={section.id}
                  onClick={() => {
                    setActiveSection(section.id);
                    setIsSidebarOpen(false); // cerrar al seleccionar en móvil
                  }}
                  className={`w-full flex items-center gap-3 px-3 py-2 rounded-lg text-left transition-colors ${
                    activeSection === section.id
                      ? "bg-sidebar-accent text-sidebar-accent-foreground"
                      : "text-sidebar-foreground hover:bg-sidebar-accent/50"
                  }`}
                >
                  <Icon className="h-4 w-4" />
                  {section.title}
                </button>
              );
            })}
          </nav>
        </aside>

        {/* Overlay en móviles */}
        {isSidebarOpen && (
          <div
            className="fixed inset-0 bg-black/50 z-40 lg:hidden"
            onClick={() => setIsSidebarOpen(false)}
          />
        )}

        {/* Main content */}
        <main className="flex-1 p-6">
          <div className="max-w-4xl mx-auto">{renderContent()}</div>
        </main>
      </div>
    </div>
  );
}
