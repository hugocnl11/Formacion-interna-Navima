# 📊 Módulo 3: Automatización de hojas de cálculo con IA

Las hojas de cálculo son una herramienta esencial en departamentos como Facturación, Contabilidad, Ecommerce y Marketing.  
Gracias a la IA, ahora es posible **generar fórmulas complejas**, **crear macros** o **analizar grandes volúmenes de datos** sin necesidad de conocimientos técnicos avanzados.

Este módulo te enseña cómo utilizar asistentes como **ChatGPT**, **Copilot** o **Gemini** para trabajar de forma más rápida y eficiente en Excel o Google Sheets.

---

## Generación automática de fórmulas complejas

¿Tienes una idea de lo que quieres calcular pero no sabes cómo escribir la fórmula? Puedes explicárselo a la IA en lenguaje natural.

### ✅ Ejemplo 1: Traducción de lenguaje natural a fórmula

**Prompt:**
```
Quiero una fórmula de Excel que sume la columna B solo si en la columna A pone "Confirmado"
```

**Respuesta esperada:**
```excel
=SUMIF(A:A, "Confirmado", B:B)
```

También puedes pedirle:
- Combinaciones de condiciones (por ejemplo, IF + AND).
- Funciones de búsqueda (VLOOKUP, XLOOKUP).
- Funciones de fecha y hora.
- Fórmulas anidadas.

**Consejo:** Copia la fórmula sugerida y pégala en tu hoja para probarla. Revísala antes de usarla en informes importantes.

---

## Creación y edición de macros con soporte de IA

Las macros permiten automatizar tareas repetitivas (copiar, filtrar, aplicar formato, mover datos, etc.).

Puedes pedirle a ChatGPT o Copilot que:
- Escriba un código básico en VBA (Excel) o Apps Script (Google Sheets).
- Explique paso a paso lo que hace una macro existente.
- Ayude a depurar errores cuando algo no funciona.

### ✅ Ejemplo 2: Crear una macro con IA

**Prompt:**
```
Crea una macro en VBA que copie la hoja activa y la renombre con la fecha actual.
```

**Respuesta esperada:**
```vba
Sub CopiarHojaConFecha()
    Dim nuevaHoja As Worksheet
    ActiveSheet.Copy After:=Sheets(Sheets.Count)
    Set nuevaHoja = ActiveSheet
    nuevaHoja.Name = Format(Date, "dd-mm-yyyy")
End Sub
```

---

## Análisis rápido y visualización de datos con IA

Puedes usar la IA para:

- Sugerir **gráficos apropiados** según el tipo de datos.
- Interpretar una tabla y destacar tendencias o anomalías.
- Automatizar paneles o dashboards con funciones de resumen.
- Proponer filtros y agrupaciones útiles para un informe.

### ✅ Ejemplo 3: Interpretación automática de una tabla

**Prompt:**
```
Tengo una tabla con ventas por categoría. ¿Qué gráfico me recomendarías para mostrar los datos?
```

**Respuesta resumida:**
> Te recomiendo un gráfico de columnas si quieres comparar categorías, o un gráfico circular si quieres mostrar la proporción de cada una respecto al total.

---

## Actividad práctica

> Utiliza ChatGPT o un asistente integrado en tu hoja de cálculo para:  
> - Crear una fórmula que calcule el total de pedidos realizados por un cliente concreto.  
> - Generar una macro que elimine las filas vacías de una hoja.  
> - Pedir una recomendación de gráfico para visualizar un resumen de ventas por mes.

Comparte tu fórmula, macro o gráfico con el grupo y explica cómo lo has generado.

---

## Recursos adicionales

- [Generador de fórmulas en lenguaje natural (Excel AI)](https://excel.microsoft.com/)
- [Editor de Apps Script para Google Sheets](https://script.google.com/)
- [Guía básica de macros en Excel (PDF)](/oficina_basico/stuff/guia_macros_excel.pdf)

---

<p align="center">
  <a href="https://hugocnl11.github.io/Formacion-interna-Navima/oficina_basico/modulo_2.html">⏮️ Módulo anterior</a> &nbsp;&nbsp;&nbsp;
  <a href="https://hugocnl11.github.io/Formacion-interna-Navima/oficina_basico/modulo_4.html">Módulo siguiente ⏭️</a>
</p>
