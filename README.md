Sistema de cálculo de liquidaciones agrícola

> Aplicación de escritorio para calcular y generar liquidaciones de sueldo,
> desarrollada en Python con interfaz gráfica Tkinter.

**** Este es un sistema de gestión de remuneraciones diseñado para empresas prestadoras de servicios agrícolas (packing, selección de fruta, etc.). Permite calcular liquidaciones de sueldo de forma rápida y exportarlas a **PDF** o **Excel**.

Desarrollado como proyecto personal / portafolio para demostrar habilidades en:
- Programación orientada a objetos con Python
- Interfaces gráficas con Tkinter
- Generación de documentos PDF (ReportLab)
- Generación de planillas Excel (openpyxl)
- Persistencia de datos con JSON

Funcionalidades

Funcionalidad | Descripción 
-----------------------------------------------------
| Liquidación completa | Calcula sueldo base, horas extras, bonos y descuentos legales |
| Exportar a Excel | Genera planilla .xlsx con formato profesional y dos hojas |
| Exportar a PDF | Genera liquidación en PDF lista para imprimir o enviar |
| Base de trabajadores | Guarda y carga trabajadores con sus datos y cargo |
| Vista previa | Revisa los totales antes de exportar |
| Buscador | Filtra trabajadores por nombre o RUT |


Labores y Sueldos...


| Labor | Sueldo Base |
|---|---|
| Paletizador/a | $610.000 |
| Seleccionadora de Fruta | $570.000 |
| Armador/a de Caja | $590.000 |
| Embaladora de Caja | $560.000 |

> Los sueldos base son configurables directamente en el campo del formulario.


Lógica...

Sueldo proporcional-- Sueldo base × (días trabajados / 30)
Horas extra---------- Número de horas × $5.000
Total imponible------ Sueldo proporcional + Horas extras

Descuento AFP-------- Total imponible × 10.00%
Descuento Salud------ Total imponible × 7.00%
Seguro Cesantía------ Total imponible × 0.60%

Total haberes-------- Total imponible + Bonos no imponibles
Total descuentos----- Desc. legales + Desc. adicionales
Líquido a pagar------- Total haberes − Total descuentos

__Instalación y su uso.

- Python 3.10 o superior
- pip

__Instala las dependencias

bash
pip install -r requirements.txt


___Ejecuta

bash
python liquidacion.py


Instalar...
reportlab>=4.0
openpyxl>=3.1
Tkinter viene incluido con python así que no requiere instalación adicional.



Estructura
sistema-liquidaciones-agricola
├── liquidacion.py        # Ventana principal e interfaz gráfica
├── calculos.py           # Lógica de negocio: cálculo de liquidaciones
├── generador_pdf.py      # Generación de PDF con ReportLab
├── generador_excel.py    # Generación de Excel con openpyxl
├── trabajadores.json     # Base de datos local (autogenerado)
├── requirements.txt      # Dependencias del proyecto
└── README.md             # Este archivo

Versión 1.0

Martín Hernández Tapia
Programming student
Chile
GitHub: [@mhernandeztapia](https://github.com/mhernandeztapia)  
