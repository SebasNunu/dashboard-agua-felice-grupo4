"""Informe Final | Grupo 4 – Agua Felice.

Dashboard Streamlit basado en el Excel final del grupo.
Mantén datos_agua_felice.xlsx en la misma carpeta que este script.
"""
from pathlib import Path
from html import escape
from io import BytesIO
import textwrap

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Agua Felice | Informe Final – Grupo 4",
    layout="wide",
    initial_sidebar_state="expanded",
)

BG = "#0E1628"
PANEL = "#1C2B47"
CYAN = "#16D1F5"
PINK = "#FF478A"
GREEN = "#4CE0B3"
PURPLE = "#A58CFF"
TEXT = "#EAF3FF"
MUTED = "#9CB3D0"

st.markdown("""
<style>
.stApp {background:#0E1628;color:#EAF3FF;}
[data-testid="stSidebar"] {background:#17233B;}
[data-testid="stSidebar"] * {color:#F2F7FF!important;}
.stApp h1,.stApp h2,.stApp h3,.stApp p,.stApp label {color:#EAF3FF;}
[data-testid="stMetric"] {background:#1C2B47;padding:18px;border:1px solid #314361;border-radius:14px;min-height:113px;}
[data-testid="stMetricValue"] * {color:#FFFFFF!important;}
[data-testid="stMetricLabel"] * {color:#BAD9FA!important;}
[data-testid="stHeader"] {background:#0E1628;}
.block-container {padding-top:1.65rem;max-width:1600px;}
div[data-testid="stDataFrame"] {border:1px solid #314361;border-radius:12px;}
/* Textos legibles y fondo oscuro en desplegables de alternativas */
[data-testid="stExpander"] summary {
    background-color: #20314F !important;
    border-radius: 8px;
}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary * {
    color: #FFFFFF !important;
}
[data-testid="stFileUploader"] {border:1px solid #314361;border-radius:12px;padding:6px;}
/* Área para subir Excel con alto contraste en el menú oscuro */
[data-testid="stFileUploaderDropzone"] {
    background-color: #20314F !important;
    border: 1px dashed #6381AD !important;
    border-radius: 12px !important;
}
[data-testid="stFileUploaderDropzone"] *,
[data-testid="stFileUploaderDropzoneInstructions"] * {
    color: #F1F7FF !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background-color: #314F7B !important;
    border: 1px solid #698AB5 !important;
    color: #FFFFFF !important;
    opacity: 1 !important;
}
[data-testid="stFileUploaderDropzone"] button * {
    color: #FFFFFF !important;
}
[data-testid="stDownloadButton"] button {border-radius:10px;}
.formula-card {
    background: #20314F;
    color: #F1F7FF;
    border: 1px solid #3B5073;
    border-radius: 10px;
    padding: 13px 16px;
    font-size: 1rem;
    line-height: 1.5;
}
</style>
""", unsafe_allow_html=True)

EXCEL = Path(__file__).resolve().parent / "datos_agua_felice.xlsx"

NOMBRES = {
    "P1": "Productividad",
    "P2": "Disponibilidad",
    "P3": "Entregas a tiempo (OTD)",
    "P4": "Rotura de stock",
    "P5": "Eficiencia global (OEE)",
    "P6": "Scrap",
    "P7": "Utilización del almacén",
}
ACTUAL = {"P1":75.36,"P2":88.47,"P3":85.82,"P4":28.13,"P5":65.66,"P6":1.50,"P7":80.66}
META = {"P1":85.0,"P2":95.0,"P3":95.0,"P4":15.0,"P5":80.0,"P6":0.80,"P7":95.0}
TIPO = {"P1":"Aumentar","P2":"Aumentar","P3":"Aumentar","P4":"Reducir","P5":"Aumentar","P6":"Reducir","P7":"Aumentar"}

DATOS_INFORME = {'P1': {'meses': [75.8, 75.6, 75.0, 73.9, 75.9, 75.7, 74.2, 75.0, 76.0, 74.0, 76.7, 76.4], 'prioridades': [{'categoria': 'Maquinaria', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Alta'}, {'categoria': 'Material', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medición', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medio ambiente', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Bajo cumplimiento del tiempo estándar de ciclo por operador → 2) Curva de aprendizaje incompleta de operadores nuevos → 3) (Raíz) Ausencia de un programa formal de certificación interna por puesto de trabajo', 'evidencia': 'Tiempo de ciclo real vs. estándar del DAP (570 s por botellón de 8 L); por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Micro-paradas recurrentes en la selladora de botellones → 2) Desgaste de empaquetaduras y rodillos de sellado → 3) (Raíz) Omisión del mantenimiento preventivo de la selladora en el plan de mantenimiento', 'evidencia': 'Productividad promedio de 75.36% frente al 100% planificado'}, {'categoria': 'Material', 'cadena': '1) Reajustes de máquina por cambio de lote de envases → 2) Variabilidad dimensional de botellones entre proveedores → 3) (Raíz) Ausencia de un criterio de aceptación dimensional en la recepción de envases', 'evidencia': 'Tiempo de ajuste por cambio de lote de insumo, min/lote; por levantar'}, {'categoria': 'Método', 'cadena': '1) Variabilidad de hasta 18% en el tiempo de ciclo entre operadores → 2) Ausencia de una Hoja de Operación Estándar (HOE) vigente → 3) (Raíz) Falta de actualización del instructivo tras el último cambio de formato de botellón', 'evidencia': 'Productividad promedio de 75.36% frente al 100% planificado (variabilidad de 18% entre operadores)'}, {'categoria': 'Medición', 'cadena': '1) Discrepancia entre la Productividad reportada y la verificada por cronometraje → 2) Registro manual del tiempo de ciclo, sujeto a error del operador → 3) (Raíz) Ausencia de un procedimiento de verificación cruzada del tiempo de ciclo registrado', 'evidencia': 'Diferencia entre Productividad reportada y verificada por cronometraje; por levantar'}, {'categoria': 'Medio ambiente', 'cadena': '1) Variación de temperatura ambiental en el área de envasado → 2) Cambios de humedad que alteran la adherencia de etiquetas → 3) (Raíz) Ausencia de control de condiciones ambientales (temperatura/humedad) en planta', 'evidencia': 'Fallas por mes vs. condiciones ambientales; por levantar'}], 'alternativas': [{'nombre': 'Ingeniería de métodos (estudio de tiempos y movimientos + balanceo de línea)', 'descripcion': 'Rediseña la asignación de tareas entre estaciones de envasado para eliminar cuellos de botella y equilibrar los tiempos de ciclo entre operadores.'}, {'nombre': '5S en el puesto de trabajo', 'descripcion': 'Ordena, estandariza y disciplina el puesto de trabajo del operador mediante instructivos y controles visuales, reduciendo la variabilidad entre turnos.'}, {'nombre': 'Kaizen (eventos de mejora continua)', 'descripcion': 'Organiza equipos multidisciplinarios que atacan de forma puntual las pérdidas de rendimiento mediante ciclos cortos de planificar, hacer, verificar y actuar.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['Ingeniería de métodos (balanceo de línea)', '2', '2', '2', '3', '2', '2.20'], ['5S en el puesto de trabajo', '5', '5', '4', '4', '4', '4.40'], ['Kaizen (eventos de mejora)', '3', '3', '3', '3', '3', '3.00']]}, 'formulas': [{'indicador': 'Productividad', 'formula': '(Producción real ÷ Producción ideal) × 100, con ideal de 65,000 unidades por mes', 'calculo': '587,772 ÷ (65,000 × 12) = 587,772 ÷ 780,000; 75.36% (razón de totales)', 'valor': '75.36%'}, {'indicador': 'Eficiencia de ciclo', 'formula': '(Tiempo de valor agregado ÷ Tiempo total del ciclo) × 100', 'calculo': '390 s ÷ 570 s (DAP, sección 1); Tiempo sin valor: 180 s (31.58%)', 'valor': '68.42%'}, {'indicador': 'Utilización de la capacidad', 'formula': '(Producción real ÷ Capacidad teórica) × 100, con capacidad de 57,600 botellas por mes', 'calculo': '48,981 ÷ 57,600 (capacidad instalada de la sección 1); Producción mensual media: 48,981 botellas', 'valor': '85.04%'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Cumplimiento de auditorías 5S (P1)', '(Estaciones que aprueban la auditoría ÷ Estaciones auditadas) × 100', '0% (no existe auditoría 5S)', '100% de las 3 estaciones; 10 auditorías quincenales entre las semanas 6 y 25 (20 semanas ÷ 2)']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['Productividad (%)', '75.36%', '85.00%', '82.75%', '80.18%'], ['Brecha respecto de la meta (puntos porcentuales)', '9.64', '0.00', '2.25', '4.82'], ['Beneficio mensual (S/)', '—', '56,421.00', '—', '—'], ['VAN del proyecto (S/)', '—', '238,652', '134,464', '20,367'], ['Recuperación de la inversión (meses)', '—', '0.02', '0.04', '—'], ['Probabilidad de VAN positivo', '—', '—', '100.0%', '—']]}, 'tabla_mensual': 14, 'tabla_prioridades': 15, 'tabla_causas': 16, 'tabla_restricciones': 18}, 'P2': {'meses': [89.2, 88.7, 90.3, 89.4, 88.8, 86.4, 88.4, 86.5, 84.8, 89.7, 91.9, 87.6], 'prioridades': [{'categoria': 'Maquinaria', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Material', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Medición', 'prioridad': 20, 'nivel': 'Media'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medio ambiente', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Intervenciones correctivas ejecutadas sin diagnóstico técnico previo → 2) Falta de capacitación en mantenimiento autónomo de primer nivel → 3) (Raíz) Ausencia de un plan de formación técnica certificada para operadores de mantenimiento', 'evidencia': '% de intervenciones correctivas sin diagnóstico previo; por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Paro del motorreductor del sistema de llenado → 2) Sobrecalentamiento por falta de lubricación periódica → 3) (Raíz) Omisión del mantenimiento preventivo programado en el plan de mantenimiento', 'evidencia': 'MTTR promedio de 4.62 h por evento de falla'}, {'categoria': 'Material', 'cadena': '1) Uso de repuestos no certificados por rotura de stock de los originales → 2) Vida útil reducida de piezas de reemplazo no homologadas → 3) (Raíz) Ausencia de una especificación técnica homologada para repuestos críticos', 'evidencia': 'Costo de gestión de repuestos críticos: S/ 911.33 a S/ 3,933.70 al año por repuesto'}, {'categoria': 'Método', 'cadena': '1) MTTR promedio de 4.62 horas por evento de falla → 2) Ausencia de procedimientos estandarizados de diagnóstico de fallas → 3) (Raíz) No existe un protocolo documentado de respuesta ante fallas críticas', 'evidencia': 'MTTR promedio de 4.62 h por evento de falla'}, {'categoria': 'Medición', 'cadena': '1) Sin historial documentado de fallas por componente → 2) Failure Rate calculado solo de forma mensual, sin trazabilidad por equipo → 3) (Raíz) Ausencia de un registro (bitácora técnica) de fallas por componente', 'evidencia': 'Failure Rate promedio de 2.84%, con pico de 3.70%'}, {'categoria': 'Medio ambiente', 'cadena': '1) Mayor frecuencia de fallas en meses de alta humedad → 2) Corrosión acelerada de piezas móviles expuestas → 3) (Raíz) Ausencia de protección anticorrosiva en componentes expuestos a humedad', 'evidencia': 'Fallas por mes vs. humedad relativa; por levantar'}], 'alternativas': [{'nombre': 'Mantenimiento Autónomo (pilar del TPM)', 'descripcion': 'Empodera a los operadores para ejecutar limpieza, inspección y lubricación diaria de sus propios equipos, siguiendo los 7 pasos del pilar.'}, {'nombre': 'Mantenimiento Centrado en Confiabilidad (RCM)', 'descripcion': 'Diseña el plan de mantenimiento a partir del análisis de modos y efectos de falla (AMEF) de los equipos de mayor criticidad.'}, {'nombre': 'Mantenimiento Preventivo Planificado (pilar del TPM)', 'descripcion': 'Programa intervenciones periódicas por horas de uso o calendario para reemplazar piezas antes de que fallen.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['Mantenimiento Autónomo (TPM)', '5', '4', '4', '4', '5', '4.40'], ['Mantenimiento Centrado en Confiabilidad (RCM)', '4', '3', '3', '4', '3', '3.40'], ['Mantenimiento Preventivo Planificado', '2', '3', '3', '4', '3', '3.00']]}, 'formulas': [{'indicador': 'MTBF y MTTR', 'formula': 'MTBF = Tiempo total de operación ÷ N.° de fallas; MTTR = Tiempo total de reparación ÷ N.° de fallas', 'calculo': 'Promedio de los 12 valores mensuales; La ficha de la empresa reporta 36.16 h y 4.62 h', 'valor': '36.17 h y 4.63 h'}, {'indicador': 'Disponibilidad', 'formula': '[MTBF ÷ (MTBF + MTTR)] × 100', 'calculo': 'Promedio de las 12 razones MTBF ÷ (MTBF + MTTR); 88.47% en la ficha de la empresa; 88.64% con los promedios anuales', 'valor': '88.45%'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Cumplimiento de la inspección autónoma (P2)', '(Inspecciones ejecutadas ÷ Inspecciones programadas) × 100', '0% (hoy no existe)', '95% o más: 612 inspecciones programadas (17 semanas × 6 días × 3 líneas × 2 turnos), es decir, al menos 582 ejecutadas']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['Disponibilidad (%)', '88.47%', '95.00%', '93.48%', '91.73%'], ['Brecha respecto de la meta (puntos porcentuales)', '6.53', '0.00', '1.52', '3.27'], ['Beneficio mensual (S/)', '—', '22,338.00', '—', '—'], ['VAN del proyecto (S/)', '—', '143,375', '104,433', '59,015'], ['Recuperación de la inversión (meses)', '—', '0.13', '0.18', '—'], ['Probabilidad de VAN positivo', '—', '—', '100.0%', '—']]}, 'tabla_mensual': 41, 'tabla_prioridades': 42, 'tabla_causas': 43, 'tabla_restricciones': 45}, 'P3': {'meses': [87.9, 88.4, 94.7, 86.0, 80.1, 88.0, 84.2, 82.0, 85.8, 86.1, 83.4, 83.2], 'prioridades': [{'categoria': 'Material', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Alta'}, {'categoria': 'Maquinaria', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medición', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medio ambiente', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Reprogramación frecuente de despachos por descoordinación entre áreas → 2) Falta de un responsable único del cumplimiento de despacho → 3) (Raíz) Ausencia de un rol formal de coordinador de despachos', 'evidencia': 'N.° de pedidos reprogramados por mes; por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Unidades de reparto insuficientes en picos de demanda → 2) Indisponibilidad de vehículos por mantenimiento no programado → 3) (Raíz) Ausencia de un plan de mantenimiento preventivo de la flota de distribución', 'evidencia': '% de pedidos retrasados por falta de unidad disponible; por levantar'}, {'categoria': 'Material', 'cadena': '1) Retrasos en el picking por variabilidad en el tiempo de preparación → 2) Ubicación no estandarizada de productos en el almacén → 3) (Raíz) Ausencia de un layout optimizado de picking en el almacén de producto terminado', 'evidencia': 'Tiempo real de picking vs. tiempo estándar de 3.3 h'}, {'categoria': 'Método', 'cadena': '1) OTD promedio de 85.82%, con mínimo mensual de 80.12% → 2) Ausencia de un sistema formal de planificación de ventas y operaciones → 3) (Raíz) Falta de un ciclo periódico de coordinación entre Producción, Logística y Ventas', 'evidencia': 'OTD promedio de 85.82%, mínimo mensual de 80.12%'}, {'categoria': 'Medición', 'cadena': '1) Sin medición sistemática del lead time de despacho por pedido → 2) Falta de un indicador de cumplimiento por ruta de reparto → 3) (Raíz) Ausencia de un procedimiento de registro de tiempos de despacho por pedido', 'evidencia': 'Lead time de despacho en horas; por levantar'}, {'categoria': 'Medio ambiente', 'cadena': '1) Congestión vehicular en rutas de reparto urbano → 2) Condiciones climáticas que retrasan la distribución en temporada de lluvias → 3) (Raíz) Ausencia de rutas alternativas planificadas para condiciones adversas', 'evidencia': 'Retrasos por clima o tránsito; por levantar'}], 'alternativas': [{'nombre': 'Planificación de Ventas y Operaciones (S&OP)', 'descripcion': 'Alinea el plan de producción con la programación de despachos mediante reuniones periódicas entre Planificación, Producción y Logística.'}, {'nombre': 'Value Stream Mapping (VSM) de la cadena logística', 'descripcion': 'Mapea el flujo completo desde el pedido hasta la entrega para identificar y eliminar esperas y reprocesos en el despacho.'}, {'nombre': 'Kanban de despacho (gestión visual)', 'descripcion': 'Ordena la preparación y salida de pedidos mediante tarjetas y tableros visuales que muestran el estado de cada despacho.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['Planificación de Ventas y Operaciones (S&OP)', '4', '3', '3', '4', '4', '3.60'], ['Value Stream Mapping (VSM) logístico', '4', '3', '3', '2', '3', '3.00'], ['Kanban de despacho', '4', '3', '4', '3', '2', '3.20']]}, 'formulas': [{'indicador': 'OTD', 'formula': '(Pedidos entregados a tiempo ÷ Pedidos totales) × 100', 'calculo': '11,281 ÷ 13,170 pedidos; 85.66% (razón de totales); 1,889 pedidos tardíos', 'valor': '85.82%'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Cumplimiento del ciclo S&OP (P3)', '(Reuniones realizadas ÷ Reuniones programadas) × 100', '0% (no existe el ciclo)', '100%: 9 reuniones mensuales (36 semanas ÷ 4)']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['OTD (entregas a tiempo) (%)', '85.82%', '95.00%', '92.86%', '90.41%'], ['Brecha respecto de la meta (puntos porcentuales)', '9.18', '0.00', '2.14', '4.59'], ['Beneficio mensual (S/)', '—', '531.67', '—', '—'], ['VAN del proyecto (S/)', '—', '2,183', '1,848', '-271'], ['Recuperación de la inversión (meses)', '—', '1.69', '2.10', '—'], ['Probabilidad de VAN positivo', '—', '—', '100.0%', '—']]}, 'tabla_mensual': 68, 'tabla_prioridades': 69, 'tabla_causas': 70, 'tabla_restricciones': 72}, 'P4': {'meses': [31.3, 20.0, 27.3, 23.5, 18.8, 29.4, 31.3, 21.1, 30.8, 26.7, 23.1, 54.5], 'prioridades': [{'categoria': 'Maquinaria', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Material', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Medición', 'prioridad': 20, 'nivel': 'Media'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medio ambiente', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Solicitudes de repuestos gestionadas de forma reactiva → 2) Sin responsable único de la gestión de repuestos críticos → 3) (Raíz) Ausencia de un rol formal de planificador de repuestos críticos', 'evidencia': 'Tiempo de respuesta ante solicitud de repuesto (días); por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Consumo variable de repuestos por fallas no programadas → 2) Correlación entre el Failure Rate y el consumo mensual de repuestos → 3) (Raíz) Ausencia de un plan de mantenimiento preventivo que estabilice el consumo', 'evidencia': 'Failure Rate promedio de 2.84%'}, {'categoria': 'Material', 'cadena': '1) 7 repuestos críticos sin clasificación previa por criticidad → 2) Costos de gestión de inventario dispersos entre S/ 911 y S/ 3,934 al año → 3) (Raíz) Ausencia de una clasificación ABC formal de los repuestos críticos', 'evidencia': '7 repuestos críticos, costo de gestión de S/ 911.33 a S/ 3,933.70 al año'}, {'categoria': 'Método', 'cadena': '1) Tasa de rotura de stock promedio de 28.13%, máximo 54.55% → 2) Ausencia de una política formal de reposición de repuestos críticos → 3) (Raíz) Falta de un procedimiento documentado de gestión de inventario de repuestos', 'evidencia': 'Tasa de rotura de stock promedio de 28.13%, máximo mensual 54.55%'}, {'categoria': 'Medición', 'cadena': '1) Entre 3 y 6 pedidos de repuestos no satisfechos por mes → 2) Sin cálculo del punto de reorden por repuesto → 3) (Raíz) Ausencia de un procedimiento de cálculo del punto de reorden y stock de seguridad', 'evidencia': '3 a 6 pedidos no satisfechos por mes'}, {'categoria': 'Medio ambiente', 'cadena': '1) Repuestos almacenados sin condiciones controladas de humedad → 2) Deterioro prematuro de componentes electrónicos de repuesto → 3) (Raíz) Ausencia de condiciones de almacenamiento controladas para repuestos sensibles', 'evidencia': 'Deterioro de repuestos en almacén; por levantar'}], 'alternativas': [{'nombre': 'Clasificación ABC y cálculo de EOQ', 'descripcion': 'Clasifica los 7 repuestos críticos por nivel de criticidad y calcula el punto de reorden y el lote óptimo (EOQ) de cada uno.'}, {'nombre': 'Kanban de reposición de repuestos críticos', 'descripcion': 'Establece un sistema visual de tarjetas que dispara automáticamente el pedido de reposición al llegar a un nivel mínimo predefinido por criticidad.'}, {'nombre': '5S en el almacén de repuestos', 'descripcion': 'Ordena, identifica y estandariza la ubicación de los repuestos para reducir tiempos de búsqueda y errores de registro.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['Clasificación ABC y cálculo de EOQ', '4', '3', '3', '4', '3', '3.40'], ['Kanban de reposición de repuestos críticos', '5', '5', '5', '3', '4', '4.40'], ['5S en el almacén de repuestos', '4', '4', '4', '3', '3', '3.60']]}, 'formulas': [{'indicador': 'Rotura de stock', 'formula': '(Pedidos de repuestos no satisfechos ÷ Pedidos totales) × 100', 'calculo': '49 ÷ 179 pedidos de repuestos; 27.37% (razón de totales)', 'valor': '28.13%'}, {'indicador': 'Nivel de servicio de repuestos', 'formula': '100 − Rotura de stock', 'calculo': '100 − Rotura de stock; Pedidos atendidos: 130 de 179', 'valor': '71.87%'}, {'indicador': 'Punto de reorden', 'formula': 'ROP = d × L + SS (d: consumo diario; L: plazo de entrega; SS: stock de seguridad)', 'calculo': 'd × L + SS; Los registros de la subsección 4.1 no incluyen el consumo diario (d) ni el plazo de entrega (L) de cada repuesto; el cálculo se hace en la tarea 4.1', 'valor': 'No calculable con los datos entregados'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Repuestos críticos bajo Kanban (P4)', '(Repuestos con tarjeta Kanban ÷ 7 repuestos críticos) × 100', '0% (0 de 7)', '42.9% (3 de 7) con el piloto de la semana 3 y 100% (7 de 7) en la semana 6']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['Rotura de stock (%)', '28.13%', '15.00%', '18.06%', '21.56%'], ['Brecha respecto de la meta (puntos porcentuales)', '13.13', '0.00', '3.06', '6.56'], ['Beneficio mensual (S/)', '—', '659.60', '—', '—'], ['VAN del proyecto (S/)', '—', '2,109', '1,420', '-375'], ['Recuperación de la inversión (meses)', '—', '1.18', '1.69', '—'], ['Probabilidad de VAN positivo', '—', '—', '99.9%', '—']]}, 'tabla_mensual': 95, 'tabla_prioridades': 96, 'tabla_causas': 97, 'tabla_restricciones': 99}, 'P5': {'meses': [66.6, 66.0, 66.6, 64.8, 66.1, 64.7, 64.9, 64.0, 63.7, 65.4, 69.2, 65.8], 'prioridades': [{'categoria': 'Material', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Medición', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Maquinaria', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medio ambiente', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Ningún responsable único hace seguimiento del indicador consolidado → 2) El cálculo del OEE se realiza de forma manual y solo una vez al mes → 3) (Raíz) Ausencia de un rol formal de dueño del indicador OEE', 'evidencia': '0 reuniones periódicas de seguimiento del OEE consolidado; por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Los equipos no cuentan con registro automático de datos de producción → 2) Los tres factores (Disponibilidad, Productividad, Calidad) se calculan por separado y en fechas distintas → 3) (Raíz) Ausencia de un mecanismo único de captura de datos de planta', 'evidencia': 'Cálculo manual y mensual del OEE, no diario; por levantar'}, {'categoria': 'Material', 'cadena': '1) Las pérdidas de material (Scrap, paradas) no se clasifican por causa específica → 2) No se distingue qué parte de la pérdida viene de Disponibilidad, Productividad o Calidad → 3) (Raíz) Ausencia de un procedimiento de clasificación de pérdidas por factor del OEE', 'evidencia': 'Scrap de 1.50% sin desagregar por causa específica'}, {'categoria': 'Método', 'cadena': '1) OEE de 65.66% frente al referente de clase mundial de 85% → 2) Ausencia de una metodología formal de cálculo integrado del OEE → 3) (Raíz) Falta de un procedimiento único que combine Disponibilidad, Productividad y Calidad', 'evidencia': 'OEE promedio de 65.66% frente al referente de clase mundial de 85%'}, {'categoria': 'Medición', 'cadena': '1) El OEE no se mide por línea ni por turno, solo como promedio mensual de planta → 2) Sin segmentación del factor que más lo afecta en cada periodo → 3) (Raíz) Ausencia de un tablero de seguimiento periódico del OEE y sus 3 factores', 'evidencia': 'Productividad (factor del OEE) de 75.36%, sin segmentar por línea o turno'}, {'categoria': 'Medio ambiente', 'cadena': '1) Condiciones de planta que afectan de forma simultánea a varios equipos → 2) Eventos ambientales no diferenciados del resto de causas de pérdida → 3) (Raíz) Ausencia de un registro de condiciones ambientales asociado al cálculo del OEE', 'evidencia': 'Fallas por condición ambiental sin registrar; por levantar'}], 'alternativas': [{'nombre': 'TPM integral (8 pilares)', 'descripcion': 'Coordina bajo un mismo programa las mejoras de Productividad, Disponibilidad y Calidad ya diseñadas para los Problemas 1, 2 y 6, con un responsable único del OEE.'}, {'nombre': 'Gestión visual de indicadores (tablero integrado)', 'descripcion': 'Consolida en un solo tablero físico, actualizado por turno, los tres factores del OEE para su seguimiento conjunto.'}, {'nombre': 'Gestión por indicadores integrados (Balanced Scorecard operativo)', 'descripcion': 'Define un cuadro de mando con metas y responsables por factor del OEE, revisado en una reunión periódica única.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['TPM integral (8 pilares)', '3', '3', '3', '5', '4', '3.60'], ['Gestión visual de indicadores (tablero integrado)', '5', '5', '5', '3', '3', '4.20'], ['Balanced Scorecard operativo', '3', '3', '2', '4', '3', '3.00']]}, 'formulas': [{'indicador': 'OEE', 'formula': 'Disponibilidad × Productividad × Calidad', 'calculo': '88.45% × 75.36% × 98.50% (Disponibilidad × Productividad × Calidad); 65.66% en la ficha de la empresa; promedio de los 12 OEE mensuales: 65.64%', 'valor': '65.65%'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Revisión semanal del OEE (P5)', '(Revisiones realizadas ÷ Semanas transcurridas) × 100', '0% (no existe)', '100%: 44 revisiones entre las semanas 9 y 52']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['OEE (%)', '65.66%', '80.00%', '76.65%', '72.83%'], ['Brecha respecto de la meta (puntos porcentuales)', '14.34', '0.00', '3.35', '7.17'], ['Beneficio mensual (S/)', '—', '8,185.94', '—', '—'], ['VAN del proyecto (S/)', '—', '52,115', '31,272', '3,059'], ['Recuperación de la inversión (meses)', '—', '0.19', '0.33', '—'], ['Probabilidad de VAN positivo', '—', '—', '100.0%', '—']]}, 'tabla_mensual': 122, 'tabla_prioridades': 123, 'tabla_causas': 124, 'tabla_restricciones': 126}, 'P6': {'meses': [1.42, 1.52, 1.59, 1.94, 1.88, 1.16, 1.18, 1.34, 1.11, 1.54, 1.8, 1.56], 'prioridades': [{'categoria': 'Maquinaria', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Medición', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Material', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Medio ambiente', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Defectos detectados solo en la inspección final, no en proceso → 2) Falta de capacitación en detección temprana de defectos → 3) (Raíz) Ausencia de un programa de capacitación en calidad en el puesto de trabajo', 'evidencia': '% de defectos detectados en la inspección final; por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Desajuste mecánico en las estaciones de sellado y etiquetado → 2) Scrap promedio de 1.50%, con máximo mensual de 1.94% → 3) (Raíz) Omisión del mantenimiento preventivo de sellado y etiquetado', 'evidencia': 'Scrap promedio de 1.50%, máximo mensual de 1.94%'}, {'categoria': 'Material', 'cadena': '1) Variabilidad de calidad en envases y tapas de proveedores → 2) Sin criterio de aceptación de insumos en la recepción → 3) (Raíz) Ausencia de una especificación técnica de aceptación de envases y tapas', 'evidencia': '% del Scrap atribuible a insumo; por levantar'}, {'categoria': 'Método', 'cadena': '1) 3 inspecciones manuales (90 s, 15.8% del ciclo) sin criterio estadístico → 2) Ausencia de Control Estadístico de Proceso en las inspecciones del DAP → 3) (Raíz) Falta de un procedimiento documentado de control estadístico de variables críticas', 'evidencia': '3 inspecciones manuales = 90 s (15.8% del ciclo de 570 s)'}, {'categoria': 'Medición', 'cadena': '1) Instrumentos de verificación de sellado sin calibración documentada → 2) Sin historial de calibración disponible → 3) (Raíz) Ausencia de un programa de calibración periódica de instrumentos de inspección', 'evidencia': 'Historial de calibración: no disponible'}, {'categoria': 'Medio ambiente', 'cadena': '1) Fluctuaciones de temperatura en el área de envasado → 2) Alteración de la viscosidad del producto antes del llenado → 3) (Raíz) Ausencia de control de temperatura en el área de llenado y sellado', 'evidencia': 'Fluctuaciones de temperatura vs. Scrap; por levantar'}], 'alternativas': [{'nombre': 'Control Estadístico de Proceso (SPC)', 'descripcion': 'Registra variables críticas de llenado, sellado y etiquetado en cartas de control para detectar desviaciones antes de que generen desecho.'}, {'nombre': 'Poka-Yoke en estaciones críticas', 'descripcion': 'Instala dispositivos a prueba de error en las estaciones de llenado, sellado y etiquetado que impiden físicamente la ocurrencia del defecto.'}, {'nombre': 'Six Sigma (ciclo DMAIC)', 'descripcion': 'Reduce la variabilidad del proceso mediante las fases de definir, medir, analizar, mejorar y controlar, con enfoque estadístico.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['Control Estadístico de Proceso (SPC)', '4', '4', '3', '4', '3', '3.60'], ['Poka-Yoke en estaciones críticas', '4', '4', '4', '3', '4', '3.80'], ['Six Sigma (DMAIC)', '2', '2', '2', '4', '3', '2.60']]}, 'formulas': [{'indicador': 'Scrap', 'formula': '(Unidades defectuosas ÷ Unidades producidas) × 100', 'calculo': '8,836 ÷ 587,772 unidades; 1.50% (razón de totales)', 'valor': '1.50%'}, {'indicador': 'Calidad', 'formula': '100 − Scrap', 'calculo': '100 − Scrap; Factor del OEE', 'valor': '98.50%'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Detecciones Poka-Yoke (P6)', 'N.° de envases retenidos por el dispositivo por semana', '0 (sin dispositivos)', 'Sin meta numérica: la línea base se fija con las 4 primeras semanas de registro tras la instalación']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['Scrap (%)', '1.50%', '0.80%', '0.96%', '1.15%'], ['Brecha respecto de la meta (puntos porcentuales)', '0.70', '0.00', '0.16', '0.35'], ['Beneficio mensual (S/)', '—', '3,100.37', '—', '—'], ['VAN del proyecto (S/)', '—', '10,785', '4,906', '-1,908'], ['Recuperación de la inversión (meses)', '—', '0.85', '1.65', '—'], ['Probabilidad de VAN positivo', '—', '—', '99.8%', '—']]}, 'tabla_mensual': 149, 'tabla_prioridades': 150, 'tabla_causas': 151, 'tabla_restricciones': 153}, 'P7': {'meses': [82.9, 79.2, 82.8, 75.6, 83.7, 80.5, 82.7, 75.5, 84.0, 83.0, 77.6, 80.3], 'prioridades': [{'categoria': 'Método', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Medio ambiente', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Material', 'prioridad': 20, 'nivel': 'Alta'}, {'categoria': 'Medición', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Mano de obra', 'prioridad': 9, 'nivel': 'Media'}, {'categoria': 'Maquinaria', 'prioridad': 9, 'nivel': 'Media'}], 'causas': [{'categoria': 'Mano de obra', 'cadena': '1) Personal coloca los ítems donde hay espacio libre en el momento → 2) Sin reglas de almacenamiento ni capacitación específica → 3) (Raíz) Ausencia de un procedimiento de almacenamiento y de capacitación del personal de almacén', 'evidencia': 'Ítems fuera de su ubicación asignada, %; por levantar'}, {'categoria': 'Maquinaria', 'cadena': '1) Carga apilada a nivel de piso, sin ubicaciones fijas → 2) Estantería y equipos de manipulación no dimensionados al volumen → 3) (Raíz) Ausencia de un dimensionamiento de estanterías según el volumen almacenado', 'evidencia': 'Volumen almacenado a nivel de piso vs. en estantería; por levantar'}, {'categoria': 'Material', 'cadena': '1) Falta de insumos en momentos críticos → 2) Insumos y repuestos sin ubicación fija, difíciles de localizar → 3) (Raíz) Ausencia de codificación e identificación de las ubicaciones', 'evidencia': 'Tiempo de localización de un ítem, min; por levantar'}, {'categoria': 'Método', 'cadena': '1) Utilización promedio de 80.66%, con mínimo mensual de 75.51% → 2) Ítems ubicados sin criterio de rotación → 3) (Raíz) Ausencia de clasificación ABC y de una política de asignación de ubicaciones', 'evidencia': 'Utilización promedio de 80.66% frente al 95% de referencia'}, {'categoria': 'Medición', 'cadena': '1) Ocupación mensual fluctúa entre 75.5% y 84.0% sin explicación registrada → 2) Volumen registrado sin unidad ni capacidad útil definidas → 3) (Raíz) Ausencia de una definición formal de la capacidad útil del almacén', 'evidencia': 'Capacidad total fija de 3,500 unidades todos los meses'}, {'categoria': 'Medio ambiente', 'cadena': '1) Espacio percibido como insuficiente aunque 19.34% del volumen está libre → 2) Zonas del almacén sin identificar ni delimitar → 3) (Raíz) Ausencia de un plano de distribución (layout) con zonas definidas', 'evidencia': '8 de 13 encuestados (61.5%) perciben espacio insuficiente; 677 unidades libres por mes'}], 'alternativas': [{'nombre': 'Clasificación ABC y ubicaciones codificadas (redistribución del almacén)', 'descripcion': 'Clasifica los ítems por rotación, define zonas en un plano de distribución y asigna a cada ítem una ubicación codificada, de modo que el espacio disponible se aproveche con un criterio único.'}, {'nombre': '5S en el almacén', 'descripcion': 'Ordena, identifica y estandariza el almacén mediante las cinco fases de las 5S, reduciendo desorden y tiempos de búsqueda, pero sin redefinir la asignación de ubicaciones.'}, {'nombre': 'Ampliación o alquiler de almacén adicional', 'descripcion': 'Incrementa la capacidad física con un espacio nuevo, sin modificar la forma en que se utiliza el espacio existente.'}], 'restricciones': {'cabeceras': ['Alternativa', 'Económica (20%)', 'Tiempo (20%)', 'Complejidad (20%)', 'Control (20%)', 'Adaptabilidad (20%)', 'Puntaje total'], 'filas': [['Clasificación ABC y ubicaciones codificadas', '4', '4', '3', '4', '4', '3.80'], ['5S en el almacén', '4', '4', '4', '3', '3', '3.60'], ['Ampliación o alquiler de almacén', '1', '2', '4', '3', '3', '2.60']]}, 'formulas': [{'indicador': 'Utilización de almacén', 'formula': '(Volumen ocupado ÷ Volumen total disponible) × 100', 'calculo': '33,876 ÷ 42,000 (volumen ocupado ÷ volumen total de los 12 meses); 80.66% (razón de totales); mínimo 75.51% (mes 8) y máximo 83.97% (mes 9)', 'valor': '80.66%'}, {'indicador': 'Volumen libre', 'formula': 'Volumen total − Volumen ocupado', 'calculo': '3,500 − 2,823; Ningún mes alcanzó la referencia de 95%', 'valor': '677 unidades (19.34%)'}], 'procesos': {'cabeceras': ['Indicador', 'Fórmula', 'Valor de partida', 'Meta y cálculo'], 'filas': [['Ítems con ubicación codificada (P7)', '(Ítems con ubicación codificada ÷ Ítems almacenados) × 100', '0% (no existe codificación)', '100% al cierre de la tarea 7.6 (semana 7)'], ['Cumplimiento de auditorías de ubicaciones (P7)', '(Ubicaciones correctas ÷ Ubicaciones auditadas) × 100', '0% (no existe la auditoría)', '95% o más: 9 auditorías quincenales entre las semanas 8 y 26 (19 semanas ÷ 2, redondeado hacia abajo)']]}, 'antes_despues': {'cabeceras': ['Concepto', 'Situación actual', 'Meta o estándar', 'Esperado (media simulada)', 'Escenario pesimista'], 'filas': [['Utilización del espacio (%)', '80.66%', '95.00%', '91.65%', '87.83%'], ['Brecha respecto de la meta (puntos porcentuales)', '14.34', '0.00', '3.35', '7.17'], ['Beneficio mensual (S/)', '—', '112.86', '—', '—'], ['VAN del proyecto (S/)', '—', '-860', '-995', '-1,621'], ['Recuperación de la inversión (meses)', '—', 'No recupera', 'No recupera', '—'], ['Probabilidad de VAN positivo', '—', '—', '0.0%', '—']]}, 'tabla_mensual': 176, 'tabla_prioridades': 177, 'tabla_causas': 178, 'tabla_restricciones': 180}}

CONCLUSIONES_ORIGINALES = ['1. Se diseñó el plan de implementación y se evaluaron siete alternativas de mejora para Agua Felicé, una por problema, con una inversión total de S/ 10,357.93 y metas medibles para cada indicador: productividad de 75.36% a 85%, disponibilidad de 88.47% a 95%, OTD de 85.82% a 95%, rotura de stock de 28.13% a 15%, OEE de 65.66% a 80%, scrap de 1.50% a 0.80% y utilización del almacén de 80.66% a 95%.', '2. En el Problema 1, las 5S en el puesto de envasado elevan la productividad en 9.64 puntos porcentuales. Con el precio y el costo del libro de la empresa, el VAN base es de S/ 238,652 y se recupera la inversión en menos de un mes. Esa magnitud depende de que cada unidad adicional se venda a S/ 10, y debe leerse con esa cautela.', '3. En el Problema 2, el mantenimiento autónomo reduce las paradas de máquina y lleva la disponibilidad a 95%. El VAN base es de S/ 143,375, con 100% de iteraciones positivas y un VAN pesimista de S/ 59,015.', '4. En el Problema 3, el S&OP mejora el cumplimiento de pedidos de 85.82% a 95% y ahorra S/ 532 al mes en costo del retraso. El VAN base es de S/ 2,183, pero depende de un costo del retraso que la empresa aún no ha levantado.', '5. En el Problema 4, el Kanban de repuestos reduce la rotura de stock de 28.13% a 15% con un VAN base de S/ 2,109, y recupera la inversión en 1.2 meses. En el escenario pesimista el VAN es negativo (–S/ 375).', '6. En el Problema 5, el tablero de OEE con enfoque TPM lleva el OEE a 80% con un VAN base de S/ 52,115. Sus beneficios se apoyan en los de los Problemas 1, 2 y 6, por lo que no deben sumarse con ellos.', '7. En el Problema 6, el Poka-Yoke baja el scrap de 1.50% a 0.80% con un VAN base de S/ 10,785 y B/C de 7.8. Es sensible al logro de la meta: en el escenario pesimista el VAN es de –S/ 1,908.', '8. En el Problema 7, la redistribución del almacén con clasificación ABC no es rentable con los datos actuales: el VAN base es de –S/ 860 y la inversión no se recupera en 6 meses. Su beneficio principal (faltantes evitados, espacio liberado y cumplimiento sanitario) aún no está monetizado, y debe levantarse antes de decidir.', '9. La simulación Monte Carlo de 2,000 iteraciones, conectada por fórmulas en el Excel, muestra que seis de las siete alternativas tienen VAN positivo en al menos 99.8% de las iteraciones. P3, P4, P6 y P7 pierden valor en el escenario pesimista, así que su éxito depende de cumplir la meta y de controlar el sobrecosto.', '10. La evaluación económica sigue la plantilla EEMP del curso: cuadro antes y después, ingreso y costo marginal, flujo de caja, VAN, TIR, PRC y B/C. Los resultados de cada alternativa se contrastan con los de otros autores en la subsección de discusión de cada problema, y las cifras de esos estudios deben verificarse en los documentos originales.', '11. Las principales limitaciones son que los beneficios de los Problemas 3, 4, 5 y 7 usan supuestos del grupo, que el COK se aplica por periodo (mes) como en la plantilla, y que los beneficios de los Problemas 1, 2, 5 y 6 se superponen en parte. Se recomienda validar estos datos con la empresa y medir cada indicador cada mes durante la implementación, para ajustar el modelo con valores reales.']


# CRONOGRAMA GANTT — Fuente: tablas 19, 46, 73, 100, 127, 154 y 181
# del Informe Final Agua Felice, Grupo 4. Sin fechas de calendario inferidas.
GANTT_TAREAS = [{'problema': 'P1', 'codigo': '1.1', 'tarea': 'Clasificar (Seiri)', 'descripcion': 'Identificar y retirar de las tres estaciones de envasado las herramientas, envases y documentos que no se usan.', 'responsable': 'Jefe de Producción', 'inicio': 1, 'fin': 1, 'entregable': 'Lista de ítems retirados', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P1', 'codigo': '1.2', 'tarea': 'Ordenar (Seiton)', 'descripcion': 'Fijar y demarcar la ubicación de herramientas, envases y etiquetas en cada estación.', 'responsable': 'Supervisor de línea', 'inicio': 2, 'fin': 2, 'entregable': 'Estaciones demarcadas', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P1', 'codigo': '1.3', 'tarea': 'Limpiar (Seiso) y elaborar la HOE', 'descripcion': 'Limpieza inicial profunda y redacción de la Hoja de Operación Estándar (HOE) con los tiempos del DAP.', 'responsable': 'Equipo de mejora continua', 'inicio': 3, 'fin': 3, 'entregable': 'HOE por estación', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P1', 'codigo': '1.4', 'tarea': '(A) Ficha del indicador y registro de limpieza', 'descripcion': 'Elaborar la ficha del indicador Productividad y el registro de limpieza y desinfección que se verifica en la auditoría.', 'responsable': 'Jefe de Calidad', 'inicio': 3, 'fin': 4, 'entregable': 'Ficha y registro aprobados', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P1', 'codigo': '1.5', 'tarea': 'Estandarizar (Seiketsu)', 'descripcion': 'Capacitar a los 8 operadores en la HOE y fijar los controles visuales.', 'responsable': 'Supervisor de línea', 'inicio': 4, 'fin': 5, 'entregable': 'Operadores capacitados', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P1', 'codigo': '1.6', 'tarea': 'Sostener (Shitsuke)', 'descripcion': 'Auditorías 5S quincenales y seguimiento mensual de la Productividad.', 'responsable': 'Jefe de Producción', 'inicio': 6, 'fin': 25, 'entregable': 'Informe de auditorías', 'tipo': 'Seguimiento', 'ajuste': False}, {'problema': 'P2', 'codigo': '2.1', 'tarea': 'Limpieza inicial profunda (Día 0)', 'descripcion': 'Limpieza a fondo de los equipos críticos y detección de anomalías con tarjetas.', 'responsable': 'Técnico de Mantenimiento y operadores', 'inicio': 1, 'fin': 1, 'entregable': 'Tarjetas de anomalías', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P2', 'codigo': '2.2', 'tarea': '(A) Bitácora de fallas y ficha de indicadores', 'descripcion': 'Registrar inicio, fin y causa de cada falla y elaborar la ficha de Disponibilidad, MTBF y MTTR.', 'responsable': 'Jefe de Mantenimiento', 'inicio': 1, 'fin': 2, 'entregable': 'Bitácora y ficha', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P2', 'codigo': '2.3', 'tarea': 'Eliminar fuentes de suciedad', 'descripcion': 'Corregir fugas y puntos de difícil acceso que generan suciedad.', 'responsable': 'Operadores y Mantenimiento', 'inicio': 2, 'fin': 3, 'entregable': 'Fuentes eliminadas', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P2', 'codigo': '2.4', 'tarea': '(A) Estándares de limpieza, lubricación e inspección', 'descripcion': 'Redactar los estándares provisionales; la lubricación se especifica con lubricantes de grado alimentario.', 'responsable': 'Jefe de Mantenimiento', 'inicio': 4, 'fin': 5, 'entregable': 'Estándares aprobados', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P2', 'codigo': '2.5', 'tarea': 'Capacitar en inspección general', 'descripcion': 'Formar a los operadores en la inspección de sus equipos.', 'responsable': 'Técnico de Mantenimiento', 'inicio': 6, 'fin': 8, 'entregable': 'Operadores capacitados', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P2', 'codigo': '2.6', 'tarea': 'Inspección autónoma plena', 'descripcion': 'Inspección diaria con lista de verificación por línea y turno; control autónomo.', 'responsable': 'Operadores', 'inicio': 9, 'fin': 25, 'entregable': 'Listas de verificación', 'tipo': 'Seguimiento', 'ajuste': False}, {'problema': 'P2', 'codigo': '2.7', 'tarea': '(A) Plan preventivo de componentes críticos', 'descripcion': 'Diseñar el plan preventivo de los componentes de mayor criticidad (segunda fase).', 'responsable': 'Jefe de Mantenimiento', 'inicio': 26, 'fin': 29, 'entregable': 'Plan preventivo', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P3', 'codigo': '3.1', 'tarea': 'Diseñar el proceso S&OP', 'descripcion': 'Definir roles, calendario y formato del plan de demanda frente a capacidad.', 'responsable': 'Jefe de Logística', 'inicio': 1, 'fin': 2, 'entregable': 'Procedimiento S&OP', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P3', 'codigo': '3.2', 'tarea': 'Reunión piloto', 'descripcion': 'Primera reunión conjunta de Planificación, Producción y Logística.', 'responsable': 'Comité S&OP', 'inicio': 3, 'fin': 3, 'entregable': 'Acta de la reunión piloto', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P3', 'codigo': '3.3', 'tarea': '(A) Registro de reclamos y verificación del vehículo', 'descripcion': 'Agregar al formato de despacho el registro de reclamos por entrega tardía y la verificación sanitaria del vehículo.', 'responsable': 'Jefe de Logística', 'inicio': 3, 'fin': 4, 'entregable': 'Formato de despacho', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P3', 'codigo': '3.4', 'tarea': 'Ciclo mensual S&OP', 'descripcion': 'Nueve reuniones mensuales con seguimiento de desviaciones.', 'responsable': 'Comité S&OP', 'inicio': 4, 'fin': 39, 'entregable': 'Actas mensuales', 'tipo': 'Seguimiento', 'ajuste': False}, {'problema': 'P4', 'codigo': '4.1', 'tarea': 'Validar ABC y calcular punto de reorden', 'descripcion': 'Validar la clasificación de los 7 repuestos y calcular su punto de reorden con consumos y plazos reales.', 'responsable': 'Jefe de Mantenimiento', 'inicio': 1, 'fin': 1, 'entregable': 'Tabla de puntos de reorden', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P4', 'codigo': '4.2', 'tarea': '(A) Ficha técnica por repuesto', 'descripcion': 'Registrar especificación técnica, proveedor y plazo de entrega real de cada repuesto crítico, y elaborar la ficha del indicador de rotura de stock.', 'responsable': 'Encargado de Almacén', 'inicio': 2, 'fin': 3, 'entregable': '7 fichas técnicas', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P4', 'codigo': '4.3', 'tarea': 'Diseñar tarjetas y tablero Kanban', 'descripcion': 'Preparar las tarjetas por repuesto y el tablero de control del almacén.', 'responsable': 'Encargado de Almacén', 'inicio': 2, 'fin': 2, 'entregable': 'Tarjetas y tablero', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P4', 'codigo': '4.4', 'tarea': 'Piloto con 3 repuestos', 'descripcion': 'Probar el sistema con los 3 repuestos de mayor criticidad.', 'responsable': 'Encargado de Almacén', 'inicio': 3, 'fin': 3, 'entregable': 'Informe del piloto', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P4', 'codigo': '4.5', 'tarea': 'Extender a los 7 repuestos', 'descripcion': 'Ampliar el Kanban y ajustar los niveles según el consumo real.', 'responsable': 'Jefe de Mantenimiento', 'inicio': 4, 'fin': 6, 'entregable': '7 repuestos bajo Kanban', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P4', 'codigo': '4.6', 'tarea': '(A) Responsable del tablero y revisión mensual', 'descripcion': 'Designar al responsable único del tablero y revisar cada mes los niveles, la rotura de stock y el nivel de servicio.', 'responsable': 'Jefe de Mantenimiento', 'inicio': 7, 'fin': 26, 'entregable': 'Acta de revisión mensual', 'tipo': 'Seguimiento', 'ajuste': True}, {'problema': 'P5', 'codigo': '5.1', 'tarea': 'Designar al responsable del OEE', 'descripcion': 'Nombrar un responsable único del seguimiento del OEE.', 'responsable': 'Jefe de Producción', 'inicio': 1, 'fin': 1, 'entregable': 'Designación formal', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P5', 'codigo': '5.2', 'tarea': 'Diseñar el tablero integrado', 'descripcion': 'Definir el formato semanal con Productividad, Disponibilidad, Calidad y OEE.', 'responsable': 'Jefe de Producción', 'inicio': 2, 'fin': 3, 'entregable': 'Tablero diseñado', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P5', 'codigo': '5.3', 'tarea': '(A) Registro de tiempos por turno', 'descripcion': 'Registrar el tiempo planificado, de parada y de producción real por turno con las definiciones de EN 13306.', 'responsable': 'Jefe de Producción', 'inicio': 2, 'fin': 3, 'entregable': 'Formato de registro', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P5', 'codigo': '5.4', 'tarea': 'Puesta en marcha', 'descripcion': 'Operar el tablero en paralelo con las mejoras de los Problemas 1, 2 y 6.', 'responsable': 'Responsable del OEE', 'inicio': 4, 'fin': 8, 'entregable': 'Tablero en uso', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P5', 'codigo': '5.5', 'tarea': 'Revisión semanal conjunta', 'descripcion': 'Reunión semanal de 1 hora con 4 participantes para revisar los tres factores.', 'responsable': 'Responsable del OEE y jefes de área', 'inicio': 9, 'fin': 52, 'entregable': 'Actas semanales', 'tipo': 'Seguimiento', 'ajuste': False}, {'problema': 'P6', 'codigo': '6.1', 'tarea': 'Identificar modos de falla por estación', 'descripcion': 'Levantar los defectos más frecuentes de llenado, sellado y etiquetado.', 'responsable': 'Jefe de Calidad', 'inicio': 1, 'fin': 2, 'entregable': 'Lista de modos de falla', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P6', 'codigo': '6.2', 'tarea': '(A) Programa de calibración', 'descripcion': 'Calibrar y registrar los instrumentos de inspección del sellado.', 'responsable': 'Jefe de Calidad', 'inicio': 1, 'fin': 3, 'entregable': 'Certificados de calibración', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P6', 'codigo': '6.3', 'tarea': '(A) Procedimiento de producto no conforme', 'descripcion': 'Definir segregación, registro y disposición del envase defectuoso, y aclarar que el Scrap cubre defectos de envase, sellado y etiquetado.', 'responsable': 'Jefe de Calidad', 'inicio': 2, 'fin': 3, 'entregable': 'Procedimiento aprobado', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P6', 'codigo': '6.4', 'tarea': 'Diseñar y probar los Poka-Yoke', 'descripcion': 'Diseñar dispositivos a prueba de error y probarlos en una estación.', 'responsable': 'Técnico de Mantenimiento', 'inicio': 3, 'fin': 4, 'entregable': 'Prototipo probado', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P6', 'codigo': '6.5', 'tarea': '(A) Instalar los dispositivos', 'descripcion': 'Instalar los dispositivos en las 3 estaciones con materiales aptos para contacto con alimentos.', 'responsable': 'Técnico de Mantenimiento', 'inicio': 5, 'fin': 6, 'entregable': '3 estaciones instaladas', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P6', 'codigo': '6.6', 'tarea': '(A) Cartas de control', 'descripcion': 'Llevar la carta p del Scrap por turno y la carta de medias y rangos de variables como el volumen de llenado.', 'responsable': 'Jefe de Calidad', 'inicio': 5, 'fin': 6, 'entregable': 'Cartas en uso', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P6', 'codigo': '6.7', 'tarea': 'Seguimiento del Scrap', 'descripcion': 'Seguimiento semanal del Scrap y ajuste de los dispositivos.', 'responsable': 'Jefe de Calidad', 'inicio': 7, 'fin': 25, 'entregable': 'Informe de Scrap', 'tipo': 'Seguimiento', 'ajuste': False}, {'problema': 'P7', 'codigo': '7.1', 'tarea': 'Medir el almacén y definir la unidad', 'descripcion': 'Medir las dimensiones del almacén y fijar la unidad de volumen y la capacidad útil, sin contar pasillos.', 'responsable': 'Jefe de Almacén', 'inicio': 1, 'fin': 2, 'entregable': 'Plano y capacidad útil', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P7', 'codigo': '7.2', 'tarea': 'Clasificar ABC por rotación', 'descripcion': 'Clasificar los ítems en A, B y C según su rotación con los datos del WMS.', 'responsable': 'Jefe de Almacén', 'inicio': 2, 'fin': 3, 'entregable': 'Lista ABC', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P7', 'codigo': '7.3', 'tarea': '(A) Ficha del indicador', 'descripcion': 'Elaborar la ficha de la Tasa de utilización de almacén con fórmula, unidad y capacidad útil.', 'responsable': 'Jefe de Almacén', 'inicio': 2, 'fin': 3, 'entregable': 'Ficha aprobada', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P7', 'codigo': '7.4', 'tarea': 'Diseñar el plano y codificar ubicaciones', 'descripcion': 'Definir zonas por clase ABC y codificar las ubicaciones (pasillo, estante, nivel) con su señalización.', 'responsable': 'Jefe de Almacén', 'inicio': 3, 'fin': 4, 'entregable': 'Plano de distribución', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P7', 'codigo': '7.5', 'tarea': '(A) Zonificación sanitaria y pasillos', 'descripcion': 'Reservar pasillos y separaciones que permitan limpieza e inspección, y validar el plano con Calidad.', 'responsable': 'Jefe de Calidad', 'inicio': 3, 'fin': 4, 'entregable': 'Plano validado', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P7', 'codigo': '7.6', 'tarea': 'Reubicar los ítems por fases', 'descripcion': 'Mover los ítems a su ubicación asignada, empezando por la clase A.', 'responsable': 'Almaceneros', 'inicio': 5, 'fin': 7, 'entregable': 'Ítems reubicados', 'tipo': 'Implementación', 'ajuste': False}, {'problema': 'P7', 'codigo': '7.7', 'tarea': '(A) Procedimiento y capacitación', 'descripcion': 'Redactar el procedimiento de almacenamiento y capacitar al personal en las reglas de ubicación.', 'responsable': 'Jefe de Almacén', 'inicio': 6, 'fin': 7, 'entregable': 'Procedimiento y registro de capacitación', 'tipo': 'Ajuste (A)', 'ajuste': True}, {'problema': 'P7', 'codigo': '7.8', 'tarea': 'Seguimiento de la utilización', 'descripcion': 'Medición mensual de la utilización y auditoría quincenal de ubicaciones.', 'responsable': 'Jefe de Almacén', 'inicio': 8, 'fin': 26, 'entregable': 'Informe mensual', 'tipo': 'Seguimiento', 'ajuste': False}]
GANTT_TABLAS = {"P1":19,"P2":46,"P3":73,"P4":100,"P5":127,"P6":154,"P7":181}
GANTT_COLORES = {"Implementación":CYAN,"Ajuste (A)":PURPLE,"Seguimiento":GREEN}


def _df_gantt(problema=None):
    base = pd.DataFrame(GANTT_TAREAS)
    if problema is not None:
        base=base.loc[base["problema"]==problema]
    return base.copy()


def _fig_gantt_detalle(df, seleccion_semana, titulo):
    fig = go.Figure()
    if df.empty:
        return fig
    orden = df.sort_values(["inicio","fin","codigo"],ascending=[True,True,True]).copy()
    orden["nombre_corto"] = orden.apply(lambda r: f"{r['codigo']} · {r['tarea'].replace('(A) ', '')[:48]}",axis=1)
    for clase in ("Implementación","Ajuste (A)","Seguimiento"):
        parte=orden.loc[orden["tipo"]==clase]
        if parte.empty:continue
        custom=parte[["tarea","responsable","entregable","descripcion","inicio","fin"]].values.tolist()
        fig.add_trace(go.Bar(
            y=parte["nombre_corto"],
            x=(parte["fin"]-parte["inicio"]+1),
            base=(parte["inicio"]-1),
            orientation="h",name=clase,
            marker=dict(color=GANTT_COLORES[clase],line=dict(width=.6,color=BG)),
            opacity=.94,customdata=custom,
            hovertemplate=("<b>%{customdata[0]}</b><br>Semanas %{customdata[4]} a %{customdata[5]}"
                           "<br>Responsable: %{customdata[1]}<br>Entregable: %{customdata[2]}"
                           "<br>Detalle: %{customdata[3]}<extra></extra>"),
        ))
    ymax=int(df["fin"].max())
    fig.add_vline(x=seleccion_semana-0.5,line=dict(color=PINK,width=2,dash="dash"),
                  annotation_text=f"Semana {seleccion_semana}",annotation_position="top right",
                  annotation_font_color=PINK)
    fig.update_layout(
        title=titulo, barmode="overlay",xaxis_title="Semana del proyecto (planificada)",
        yaxis_title="",xaxis=dict(range=[0,max(ymax+2,10)],
                                    dtick=1 if ymax<=29 else 4,gridcolor="#34435A"),
        yaxis=dict(categoryorder="array",categoryarray=orden["nombre_corto"].tolist()[::-1]),
        hovermode="closest",
    )
    fig = formato_fig(fig,max(350,len(df)*46+105))
    fig.update_layout(legend=dict(orientation="h",y=-0.15,font=dict(color=TEXT,size=12)),
                      margin=dict(l=10,r=15,t=58,b=70))
    return fig


def _fig_gantt_resumen(df,semana):
    rows=[]
    for p in NOMBRES:
        part=df[df["problema"]==p]
        if len(part):
            rows.append(dict(problema=p,nombre=NOMBRES[p],fin=int(part["fin"].max()),tareas=len(part)))
    fig=go.Figure()
    paleta=[CYAN,PINK,GREEN,PURPLE,"#F6C75C","#A0E7EA","#E06CB5"]
    for j,r in enumerate(rows):
        fig.add_trace(go.Bar(x=[r['fin']],y=[f"{r['problema']} · {r['nombre']}"],orientation="h",
                      marker_color=paleta[j],name=r['problema'],showlegend=False,
                      customdata=[[r['tareas'],r['nombre']]],
                      text=[f"Semana {r['fin']}"],textposition="outside",
                      hovertemplate=("<b>%{customdata[1]}</b><br>Semana 1 a %{x}"
                                     "<br>%{customdata[0]} actividades<extra></extra>")))
    fig.add_vline(x=semana-.5,line_color=PINK,line_dash="dash",line_width=2,
                  annotation_text=f"Semana {semana}",annotation_font_color=PINK)
    fig.update_layout(title="Horizonte de las siete propuestas · 52 semanas",
                      xaxis_title="Semana del proyecto",yaxis_title="",
                      xaxis=dict(range=[0,56],dtick=4),
                      yaxis=dict(categoryorder="array",categoryarray=[f"{r['problema']} · {r['nombre']}" for r in rows][::-1]))
    return formato_fig(fig,410)


def mostrar_gantt(problema=None,vista_completa=False):
    gantt=_df_gantt(problema)
    if gantt.empty:
        st.info("No se encontraron actividades para mostrar.")
        return
    source_tables=", ".join(str(GANTT_TABLAS[p]) for p in (NOMBRES if problema is None else [problema]))
    fin=int(gantt["fin"].max())
    cp= "general" if problema is None else problema.lower()
    if problema is None:
        tot=st.columns(4)
        tot[0].metric("Propuestas",7)
        tot[1].metric("Actividades planificadas",len(gantt))
        tot[2].metric("Horizonte del plan","52 semanas")
        tot[3].metric("Inversión conjunta",moneda(float(resumen["Inversión (S/)"].sum())))
        semana=st.slider("Consultar actividades programadas en la semana",min_value=1,max_value=52,value=7,key=f"sem_gantt_{cp}")
        st.plotly_chart(_fig_gantt_resumen(gantt,semana),width="stretch",key=f"plot_gantt_resumen_{cp}")
        st.markdown("#### Cronograma detallado por alternativa")
        elegido=st.selectbox("Selecciona la alternativa para ver sus actividades",
                             list(NOMBRES),format_func=lambda p:f"{p} · {NOMBRES[p]}",key="gantt_selector_consolidado")
        detallado=_df_gantt(elegido)
        st.plotly_chart(_fig_gantt_detalle(detallado,semana,f"Detalle de implementación · {elegido}"),
                        width="stretch",key=f"plot_gantt_consol_{elegido}")
        st.caption("Plan a 52 semanas según el Anexo 2 y las tablas de tareas del informe; "
                   "las barras no representan avances de ejecución verificados.")
    else:
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Actividades",len(gantt))
        c2.metric("Semanas del plan",f"1–{fin}")
        c3.metric("Responsables",gantt['responsable'].nunique())
        row=resumen.loc[resumen["Problema"]==problema]
        c4.metric("Inversión estimada",moneda(float(row["Inversión (S/)"].iloc[0])) if len(row) else "—")
        semana=st.slider("Semana del proyecto",min_value=1,max_value=fin,value=min(7,fin),key=f"sem_gantt_{cp}")
        tipos=st.multiselect("Mostrar tipos de actividad",list(GANTT_COLORES),default=list(GANTT_COLORES),key=f"filtro_gantt_{cp}")
        gantt=gantt[gantt["tipo"].isin(tipos)]
        if not len(gantt):
            st.info("Selecciona por lo menos un tipo de actividad.")
            return
        st.plotly_chart(_fig_gantt_detalle(gantt,semana,f"Cronograma por tarea · {problema}"),
                        width="stretch",key=f"plot_gantt_{cp}")
    programadas = gantt.loc[(gantt["inicio"]<=semana)&(gantt["fin"]>=semana)]
    st.markdown(f"#### Actividades programadas para la semana {semana}")
    if programadas.empty:
        st.caption("No hay actividades programadas en esa semana según el cronograma del informe.")
    else:
        st.dataframe(programadas[["problema","codigo","tarea","responsable","entregable","inicio","fin"]]
                     .rename(columns={"problema":"Problema","codigo":"Código","tarea":"Actividad",
                                     "responsable":"Responsable","entregable":"Entregable",
                                     "inicio":"Inicio (sem.)","fin":"Fin (sem.)"}),
                     hide_index=True,width="stretch")
    with st.expander("Consultar todas las actividades, responsables y entregables"):
        st.dataframe(gantt[["codigo","tarea","descripcion","responsable","inicio","fin","entregable"]]
                     .rename(columns={"codigo":"Código","tarea":"Tarea","descripcion":"Descripción",
                                     "responsable":"Responsable","inicio":"Inicio (sem.)",
                                     "fin":"Fin (sem.)","entregable":"Entregable"}),
                     hide_index=True,width="stretch",height=340)
        st.download_button("Descargar Gantt como CSV",
                           gantt.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"Gantt_AguaFelice_{problema or 'Grupo4'}.csv",mime="text/csv",
                           key=f"descarga_gantt_{cp}")
    st.caption(f"Fuente: Informe Final Agua Felice Grupo 4, Tabla(s) {source_tables}; "
               "Anexo 2 para el consolidado. Inicio/fin expresados en semanas relativas, "
               "sin fechas calendario ni avance real de ejecución reportados.")


# ISHIKAWA INTERACTIVO — Se basa en las seis categorías y las causas raíz
# de las tablas de priorización y 5 porqués del informe final del grupo.
ISHIKAWA_FIGURAS = {"P1":4,"P2":8,"P3":12,"P4":16,"P5":21,"P6":25,"P7":29}
ISHIKAWA_CATEGORIAS = ["Mano de obra","Maquinaria","Material","Método","Medición","Medio ambiente"]
ISHIKAWA_AMARILLO = "#FFD166"


def causa_raiz_de_cadena(cadena: str) -> str:
    """Extrae la expresión marcada como (Raíz) sin interpretarla ni reescribirla."""
    return cadena.rsplit("(Raíz)",1)[-1].strip() if "(Raíz)" in cadena else cadena


def datos_causas_ishikawa(problema: str):
    info = DATOS_INFORME[problema]
    por_categoria = {r["categoria"]:r for r in info["causas"]}
    scores = {r["categoria"]:r for r in info["prioridades"]}
    rows=[]
    for cat in ISHIKAWA_CATEGORIAS:
        c=por_categoria.get(cat, {})
        p=scores.get(cat,{})
        rows.append({
            "categoria":cat,
            "raiz":causa_raiz_de_cadena(c.get("cadena","No reportado")),
            "cadena":c.get("cadena",""),
            "evidencia":c.get("evidencia","No reportada"),
            "prioridad":p.get("prioridad",None),
            "nivel":p.get("nivel","No determinado"),
        })
    return rows


def diagrama_ishikawa(problema: str):
    rows=datos_causas_ishikawa(problema)
    fig=go.Figure()
    # Eje principal, cabeza y seis ramas de causa-efecto.
    fig.add_trace(go.Scatter(x=[.65,10.75],y=[0,0],mode="lines",
                             line=dict(color=CYAN,width=5),hoverinfo="skip",showlegend=False))
    fig.add_annotation(x=10.9,y=0,ax=10.05,ay=0,showarrow=True,
                       arrowhead=3,arrowsize=1.7,arrowwidth=3,arrowcolor=CYAN,text=" ")
    fig.add_shape(type="rect",x0=10.85,x1=12.65,y0=-.72,y1=.72,
                  fillcolor="#293754",line=dict(color=PINK,width=2))
    fig.add_annotation(x=11.75,y=0,text=f"<b>PROBLEMA<br>{problema}</b>",
                       showarrow=False,font=dict(size=17,color="#FFFFFF"))
    xs=[2.6,5.6,8.6,2.6,5.6,8.6]
    for idx,item in enumerate(rows):
        up=idx<3
        sy=1 if up else -1
        x=xs[idx]
        endpoint_x=x-.95
        endpoint_y=2.15*sy
        critical=item["nivel"]=="Alta"
        color=ISHIKAWA_AMARILLO if critical else CYAN
        fig.add_trace(go.Scatter(x=[x,endpoint_x],y=[0,endpoint_y],mode="lines",
                                 line=dict(color=color,width=3),hoverinfo="skip",showlegend=False))
        detail=(f"<b>{escape(item['categoria'])}</b><br>"
                f"Causa raíz: {escape(item['raiz'])}<br>"
                f"Puntaje: {item['prioridad']} · Nivel: {escape(item['nivel'])}<br>"
                f"Evidencia: {escape(item['evidencia'])}<extra></extra>")
        fig.add_trace(go.Scatter(x=[endpoint_x],y=[endpoint_y],mode="markers",
                                 marker=dict(size=17,color=color,line=dict(color=BG,width=2)),
                                 hovertemplate=detail,showlegend=False))
        short=textwrap.shorten(item["raiz"],width=43,placeholder="…")
        label=f"<b>{escape(item['categoria'])}</b><br><span style='font-size:11px'>{escape(short)}</span>"
        fig.add_annotation(x=endpoint_x,y=endpoint_y+(0.30*sy),showarrow=False,
                           text=label,xanchor="center",yanchor="bottom" if up else "top",
                           font=dict(color=color,size=14),align="center")
    fig.update_layout(
        title="Diagrama causa-efecto · 6M",
        xaxis=dict(visible=False,range=[0,12.9]),
        yaxis=dict(visible=False,range=[-3.6,3.6],fixedrange=True),
        showlegend=False,hovermode="closest",
        margin=dict(l=5,r=12,t=62,b=25),
        height=535,
    )
    return formato_fig(fig,535)


def mostrar_ishikawa(problema: str, origen="menu"):
    info=DATOS_INFORME[problema]
    rows=datos_causas_ishikawa(problema)
    ncrit=sum(r["nivel"]=="Alta" for r in rows)
    st.caption(f"Diagrama del informe (Figura {ISHIKAWA_FIGURAS[problema]}) · "
               "Las causas de nivel ALTA se resaltan en amarillo según la clasificación del grupo.")
    a,b,c=st.columns(3)
    a.metric("Categorías de causa",len(rows))
    b.metric("Causas clasificadas como altas",ncrit)
    c.metric("Método de diagnóstico","6M + 5 porqués")
    st.plotly_chart(diagrama_ishikawa(problema),width="stretch",key=f"ishikawa_fig_{origen}_{problema}")
    st.caption("Pasa el cursor sobre cada punto para ver la causa raíz, prioridad y evidencia. "
               "La valoración y las categorías se reproducen como figuran en el informe; "
               "las causas no implican que todas estén confirmadas con mediciones de campo.")
    seleccionado=st.selectbox(
        "Explora una causa y su análisis de 5 porqués",
        ISHIKAWA_CATEGORIAS,
        format_func=lambda cat:("Alta · " if next(r for r in rows if r['categoria']==cat)['nivel']=="Alta" else "Media · ")+cat,
        key=f"ishikawa_cat_{origen}_{problema}",
    )
    causa=next(r for r in rows if r["categoria"]==seleccionado)
    left,right=st.columns([1.6,1])
    with left:
        st.markdown(f"**Causa raíz · {seleccionado}**")
        st.markdown(f"<div class='formula-card'>{escape(causa['raiz'])}</div>",unsafe_allow_html=True)
        with st.expander("Ver la cadena causal y los 5 porqués"):
            for i,parte in enumerate(causa["cadena"].split("→"),start=1):
                st.markdown(f"**Paso {i}.** {parte.strip()}")
    with right:
        st.metric("Puntaje de priorización",causa["prioridad"])
        st.write(f"Nivel reportado: **{causa['nivel']}**")
        st.caption(f"Evidencia o KPI: {causa['evidencia']}")
    st.caption(f"Fuente: Informe Final Agua Felice Grupo 4, figura {ISHIKAWA_FIGURAS[problema]}, "
               f"tablas {info['tabla_prioridades']} y {info['tabla_causas']}.")

@st.cache_data(show_spinner=False)
def leer_resumen(ruta) -> pd.DataFrame:
    tabla = pd.read_excel(BytesIO(ruta) if isinstance(ruta, bytes) else ruta, sheet_name="Resumen", header=2, nrows=7, engine="openpyxl")
    tabla = tabla.loc[tabla["Problema"].isin(NOMBRES)].copy()
    if len(tabla) != 7:
        raise ValueError("No se encontraron los siete problemas en la hoja Resumen.")
    return tabla

@st.cache_data(show_spinner=False)
def leer_montecarlo(ruta, problema: str) -> pd.DataFrame:
    # En las hojas Sim P1–Sim P7, los títulos están en la fila 30.
    tabla = pd.read_excel(BytesIO(ruta) if isinstance(ruta, bytes) else ruta, sheet_name=f"Sim {problema}", header=29,
                          nrows=2000, usecols="A:N", engine="openpyxl")
    for col in ("VAN (S/)", "Inversión (S/)", "Beneficio mensual (S/)"):
        tabla[col] = pd.to_numeric(tabla[col], errors="coerce")
    tabla["Iter."] = pd.to_numeric(tabla["Iter."], errors="coerce")
    return tabla.dropna(subset=["Iter.", "VAN (S/)", "Inversión (S/)", "Beneficio mensual (S/)"]).copy()

@st.cache_data(show_spinner=False)
def leer_sensibilidad(ruta, problema: str) -> pd.DataFrame:
    tabla = pd.read_excel(BytesIO(ruta) if isinstance(ruta, bytes) else ruta, sheet_name=f"Sim {problema}",
                          header=3, nrows=8, usecols="D:E", engine="openpyxl")
    tabla.columns = ["Variable", "Correlación"]
    tabla["Correlación"] = pd.to_numeric(tabla["Correlación"], errors="coerce")
    return tabla.dropna(subset=["Correlación"])

def formato_fig(fig, altura=350):
    if not fig.layout.title.text:
        fig.update_layout(title=dict(text=" "))
    fig.update_layout(template="plotly_dark", height=altura,
                      paper_bgcolor=BG, plot_bgcolor=PANEL,
                      font=dict(color=TEXT), title_font=dict(color=TEXT,size=16), margin=dict(l=12,r=12,t=48,b=20),
                      legend=dict(orientation="h", y=-0.23, font=dict(color=TEXT,size=12)), 
                      hoverlabel=dict(bgcolor=PANEL,font_color=TEXT))
    fig.update_xaxes(gridcolor="#34435A",zerolinecolor="#7A91A9", tickfont=dict(color=MUTED), title_font=dict(color=TEXT))
    fig.update_yaxes(gridcolor="#34435A",zerolinecolor="#7A91A9", tickfont=dict(color=MUTED), title_font=dict(color=TEXT))
    fig.update_annotations(font_color=TEXT)
    return fig

def moneda(valor):
    prefijo = "−S/ " if valor < 0 else "S/ "
    return prefijo + f"{abs(valor):,.2f}"


def factor_presente(r: float, n: int) -> float:
    """Factor de una anualidad mensual de n meses (sin valor residual)."""
    return float(n) if abs(r) < 1e-12 else -np.expm1(-n*np.log1p(r))/r


def tir_beneficio_constante(inversion: float, beneficio: float, n: int):
    """TIR por periodo para inversión en t=0 y beneficios constantes mensuales."""
    if inversion <= 0 or beneficio <= 0 or n <= 0:
        return None
    def van(r):
        return -inversion + beneficio * factor_presente(r, n)
    minimo = -0.999
    if van(minimo) <= 0:
        return None
    maximo = 1.0
    while van(maximo) > 0 and maximo < 1e7:
        maximo = maximo*2 + 1
    if van(maximo) > 0:
        return None
    for _ in range(75):
        mitad = (minimo + maximo)/2
        if van(mitad) > 0:
            minimo = mitad
        else:
            maximo = mitad
    return (minimo+maximo)/2


def mostrar_simulador(problema: str, fila: pd.Series, mc: pd.DataFrame, prefijo: str):
    """Escenario hipotético: recalcula el VAN utilizando las filas del Excel, sin escribir al archivo."""
    st.markdown(f"#### Simulador de escenarios · {problema}: {NOMBRES[problema]}")
    st.caption("Los controles recalculan escenarios hipotéticos a partir del Excel. El archivo y las 2,000 filas de origen permanecen intactos.")
    c1,c2,c3=st.columns(3)
    with c1:
        if problema == "P7":
            minutos=st.slider("Minutos ahorrados por operario y día", 0, 120, 30, 5,
                              key=f"{prefijo}_{problema}_minutos")
            beneficio_pct=100.0*minutos/30.0
        else:
            beneficio_pct=st.slider("Beneficio / ahorro mensual (%)", 0, 350, 100, 5,
                                    key=f"{prefijo}_{problema}_benef")
    with c2:
        inversion_pct=st.slider("Inversión inicial (%)", 50, 250, 100, 5, key=f"{prefijo}_{problema}_inv")
    with c3:
        cok_pct=st.number_input("COK por periodo (%)", min_value=0.0, max_value=40.0,
                                value=10.8513, step=0.25, format="%.4f", key=f"{prefijo}_{problema}_cok")
    a,b=st.columns([1.25,1])
    with a:
        modo=st.radio("Iteraciones", ["Usar las 2,000 iteraciones de origen", "Nuevo remuestreo de escenarios"],
                      horizontal=True, key=f"{prefijo}_{problema}_modo")
    with b:
        if modo == "Nuevo remuestreo de escenarios":
            cantidad=st.select_slider("Cantidad", options=[500,1000,2000,3000,5000], value=2000,
                                      key=f"{prefijo}_{problema}_n")
            semilla=st.number_input("Semilla reproducible",value=2026,min_value=0,max_value=999999,
                                   key=f"{prefijo}_{problema}_seed")
        else:
            cantidad=len(mc)
            semilla=None
    if modo == "Nuevo remuestreo de escenarios":
        # Remuestreo empírico con reemplazo: no equivale a reestimar las distribuciones originales.
        rng=np.random.default_rng(int(semilla))
        indices=rng.integers(0,len(mc),size=int(cantidad))
        simulaciones=mc.iloc[indices].copy().reset_index(drop=True)
    else:
        simulaciones=mc.copy().reset_index(drop=True)
    n=int(fila["Horizonte N (meses)"])
    tasa=cok_pct/100.0
    factor=factor_presente(tasa,n)
    inv_original=simulaciones["Inversión (S/)"].to_numpy(dtype=float)
    bene_original=simulaciones["Beneficio mensual (S/)"].to_numpy(dtype=float)
    inv_nueva=inv_original*inversion_pct/100
    bene_nueva=bene_original*beneficio_pct/100
    van_nuevo=-inv_nueva + bene_nueva*factor
    base_van=mc["VAN (S/)"].to_numpy(dtype=float)
    inversion_base=float(fila["Inversión (S/)"])
    beneficio_base=float(fila["Beneficio mensual (S/)"])
    inv_caso=inversion_base*inversion_pct/100
    bene_caso=beneficio_base*beneficio_pct/100
    van_caso=-inv_caso + bene_caso*factor
    tir_caso=tir_beneficio_constante(inv_caso,bene_caso,n)
    b_c=(bene_caso*n/inv_caso) if inv_caso>0 else None
    c1,c2,c3,c4=st.columns(4)
    c1.metric("VAN del caso simulado", moneda(van_caso),
              delta=f"{van_caso-float(fila['VAN base (S/)']):+,.0f} S/")
    c2.metric("VAN medio de iteraciones", moneda(float(van_nuevo.mean())))
    c3.metric("Prob. VAN positivo",f"{100*float(np.mean(van_nuevo>0)):.1f} %")
    c4.metric("TIR del caso simulado", f"{100*tir_caso:,.1f} %" if tir_caso is not None else "No calculable")
    c5,c6,c7=st.columns(3)
    c5.metric("Inversión simulada",moneda(inv_caso))
    c6.metric("Beneficio mensual simulado",moneda(bene_caso))
    c7.metric("Relación B/C del caso simulado",f"{b_c:.2f}" if b_c is not None else "—")
    izq,der=st.columns(2)
    with izq:
        fig=go.Figure()
        fig.add_trace(go.Histogram(x=base_van,nbinsx=42,name="Monte Carlo original",marker_color=PURPLE,opacity=.58))
        fig.add_trace(go.Histogram(x=van_nuevo,nbinsx=42,name="Escenario interactivo",marker_color=CYAN,opacity=.70))
        fig.add_vline(x=0,line_dash="dash",line_color=PINK)
        fig.update_layout(title="Distribución de VAN · comparación",barmode="overlay",
                          xaxis_title="VAN (S/)",yaxis_title="Frecuencia")
        st.plotly_chart(formato_fig(fig,355),width="stretch")
    with der:
        tabla_bar=pd.DataFrame({"Cálculo":["Caso base original","Caso base simulado","Media original","Media simulada"],
                               "VAN (S/)":[float(fila["VAN base (S/)"]),van_caso,
                                           float(np.mean(base_van)),float(np.mean(van_nuevo))]})
        fig=px.bar(tabla_bar,x="Cálculo",y="VAN (S/)",color="Cálculo",
                   color_discrete_sequence=[PURPLE,CYAN,PINK,GREEN],text="VAN (S/)")
        fig.update_traces(texttemplate="%{y:,.0f}",textposition="outside")
        fig.add_hline(y=0,line_color=PINK,line_dash="dash")
        fig.update_layout(title="Comparación de resultados",showlegend=False,xaxis_title="")
        st.plotly_chart(formato_fig(fig,355),width="stretch")
    # Sensibilidad del VAN determinista al cambio en el beneficio mensual
    niveles=np.linspace(0,5 if problema == "P7" else 3.5,101)
    vandesc=-inv_caso + beneficio_base*niveles*factor
    linea=go.Figure(go.Scatter(x=100*niveles,y=vandesc,mode="lines",line=dict(color=CYAN,width=3),
                              name="VAN según beneficio"))
    linea.add_hline(y=0,line_dash="dash",line_color=PINK)
    linea.add_vline(x=beneficio_pct,line_dash="dot",line_color=GREEN)
    linea.update_layout(title="¿Cuánto debe variar el ahorro para mejorar el VAN?",
                        xaxis_title="Beneficio mensual respecto al original (%)",yaxis_title="VAN (S/)")
    st.plotly_chart(formato_fig(linea,290),width="stretch")
    salida=simulaciones[["Iter.","Inversión (S/)","Beneficio mensual (S/)","VAN (S/)"]].copy()
    salida["Iteración del escenario"]=np.arange(1,len(salida)+1)
    salida["Inversión escenario (S/)"]=inv_nueva
    salida["Beneficio escenario (S/)"]=bene_nueva
    salida["VAN escenario (S/)"]=van_nuevo
    with st.expander("Ver y descargar los resultados de este escenario"):
        st.dataframe(salida,hide_index=True,width="stretch",height=260)
        st.download_button("Descargar escenario CSV",data=salida.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"Escenario_{problema}.csv",mime="text/csv",key=f"{prefijo}_{problema}_csv")
    if problema == "P7":
        st.caption("En P7, los minutos ahorrados modifican proporcionalmente el beneficio monetario del modelo; "
                   "no se está simulando el recorrido físico de almacén.")
    st.caption("Método: las 2,000 filas se utilizan como insumo, pero al modificar un control sus VAN se "
               "recalculan para el escenario; no son ya los VAN originales. "
               "VAN = −inversión + valor presente de los beneficios mensuales constantes. "
               "El remuestreo toma filas con reemplazo. Es un escenario de análisis, no una nueva validación "
               "operativa ni una modificación de los resultados oficiales.")



# Modelo de simulación documentado (secciones 1.10 a 7.10 del informe final).
# Los datos empresariales son los consignados por el grupo, no están validados independientemente.
MODELO_SIMULACION = {
    "P1": {
        "proceso": "Envasado de botellones y productividad del puesto",
        "accion": "Orden y estandarización del puesto mediante 5S para disminuir pérdidas productivas",
        "base": "Producción media de 48,981 unidades/mes; productividad 75.36 % (registro 2025).",
        "hipotesis": "Precio S/ 10 y costo S/ 1 por unidad, grado de logro de la meta, sobrecostos de implementación y otros parámetros del Excel.",
        "beneficio": "Margen atribuido a unidades adicionales vendibles según el logro supuesto de la meta.",
        "referencia": "Secciones 1.9–1.10; Tablas 30–32",
    },
    "P2": {
        "proceso": "Paradas del equipo de llenado y disponibilidad",
        "accion": "Mantenimiento Autónomo para disminuir fallas y mejorar disponibilidad",
        "base": "Disponibilidad 88.47 % (2025); el modelo parte de 10 paradas semanales.",
        "hipotesis": "Reducción proyectada de paradas hasta 1 por semana, S/ 511.25 por parada, recuperación supuesta de 437 unidades y grado de logro.",
        "beneficio": "Ahorro de paradas más margen asociado a unidades recuperadas, según los supuestos.",
        "referencia": "Secciones 2.9–2.10; Tablas 57–59",
    },
    "P3": {
        "proceso": "Atención de pedidos y entregas dentro del plazo (OTD)",
        "accion": "S&OP para mejorar la coordinación entre producción, logística y ventas",
        "base": "OTD 85.82 % (2025); referencia de 1,097.5 pedidos/mes en el informe.",
        "hipotesis": "Valor medio supuesto de S/ 264 por pedido y penalidad de retraso estimada como porcentaje del pedido (no cuantificada por la empresa).",
        "beneficio": "Costo evitado de pedidos tardíos al alcanzar parcialmente la mejora propuesta.",
        "referencia": "Secciones 3.9–3.10; Tablas 84–86",
    },
    "P4": {
        "proceso": "Disponibilidad de repuestos críticos y rotura de stock",
        "accion": "Kanban de reposición para controlar inventarios de repuestos",
        "base": "Rotura de stock 28.13 % (2025); selección de 7 repuestos para Kanban.",
        "hipotesis": "Costo anual por repuesto entre S/ 911.33 y S/ 3,933.70, reducción proporcional de costos según la disminución proyectada de rotura.",
        "beneficio": "Ahorro supuesto de gestión de inventario de los siete repuestos.",
        "referencia": "Secciones 4.9–4.10; Tablas 111–113",
    },
    "P5": {
        "proceso": "Eficiencia global del equipo (OEE)",
        "accion": "Tablero OEE con enfoque TPM y seguimiento de pérdidas",
        "base": "OEE 65.66 % (2025); meta de 80 %.",
        "hipotesis": "Parte atribuible al tablero del 10 % (rango 5–15 %) de los beneficios supuestos de P1, P2 y P6.",
        "beneficio": "Fracción estimada de beneficios atribuible al tablero; no debe duplicarse al consolidar con P1, P2 y P6.",
        "referencia": "Secciones 5.9–5.10; Tablas 138–140",
    },
    "P6": {
        "proceso": "Defectos y merma en la producción",
        "accion": "Poka-Yoke y control del proceso para reducir el scrap",
        "base": "Scrap 1.50 % (2025); producción media de 48,981 unidades/mes.",
        "hipotesis": "Meta de merma 0.80 %; precio S/ 10 y costo S/ 1 por unidad adicional apta, grado de logro y sobrecosto.",
        "beneficio": "Margen proyectado por unidades adicionales aptas para venta.",
        "referencia": "Secciones 6.9–6.10; Tablas 165–167",
    },
    "P7": {
        "proceso": "Localizaciones y utilización del almacén",
        "accion": "Clasificación ABC y redistribución de ubicaciones",
        "base": "Utilización 80.66 % en 2025: 2,823 unidades ocupadas de 3,500 de capacidad media mensual.",
        "hipotesis": "2 operarios × 30 minutos/día × 22 días × S/ 5.13/h de tiempo de búsqueda valorizado; logro de la meta y sobrecosto.",
        "beneficio": "Tiempo de búsqueda valorizado. No incluye faltantes evitados ni capacidad efectivamente liberada.",
        "referencia": "Secciones 7.9–7.10; Tablas 192–194",
    },
}


def _etapa_modelo(titulo, subtitulo, texto, color):
    # Recuadros compactos tipo dashboard; no dependen de imágenes externas.
    st.markdown(
        f'<div style="min-height:160px;background:{PANEL};border:1px solid #345072;'
        f'border-top:4px solid {color};border-radius:14px;padding:16px 15px;">'
        f'<div style="font-size:13px;color:#B9CDE9;">{escape(subtitulo)}</div>'
        f'<div style="font-size:19px;font-weight:700;color:#F4FAFF;margin:6px 0;">{escape(titulo)}</div>'
        f'<div style="font-size:13px;line-height:1.5;color:#DCE9FF;">{escape(texto)}</div></div>',
        unsafe_allow_html=True,
    )


def mostrar_modelo_simulacion(problema, fila, inf, fuente, prefijo="modelo"):
    """Presentación visual del modelo reportado; no modifica ni crea datos operativos."""
    det = MODELO_SIMULACION[problema]
    st.subheader(f"Modelo de simulación · {problema}: {NOMBRES[problema]}")
    st.caption(f"Fuente: Informe Final Grupo 4, {det['referencia']}; Excel activo, hoja «Sim {problema}». ")
    f1, f2, f3, f4 = st.columns(4, gap="small")
    with f1:
        _etapa_modelo("AS-IS · Situación inicial", "01 · Datos históricos", det["base"], CYAN)
    with f2:
        _etapa_modelo("Supuestos", "02 · Condiciones del modelo", det["hipotesis"], PURPLE)
    with f3:
        _etapa_modelo("Monte Carlo", "03 · Modelo financiero", "2,000 iteraciones con distribuciones triangulares independientes y transformada inversa.", PINK)
    with f4:
        _etapa_modelo("Resultados del modelo", "04 · Indicadores de salida", "VAN, B/C, PRC, percentiles, probabilidad de VAN positivo y sensibilidad.", GREEN)
    st.caption("Secuencia conceptual del cálculo: datos del informe → supuestos → iteraciones → resultados financieros. "
               "Los bloques no representan flujos físicos de material ni tiempos reales de operación.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("AS-IS · Indicador 2025", f"{ACTUAL[problema]:.2f} %")
    c2.metric("TO-BE · Meta propuesta", f"{META[problema]:.2f} %")
    c3.metric("Horizonte financiero", f"{int(fila['Horizonte N (meses)'])} meses")
    c4.metric("Iteraciones", "2,000")

    t1,t2,t3=st.tabs(["Comparación AS-IS / TO-BE", "Variables y simulación", "Resultados y sustento"])
    with t1:
        primeras=inf["antes_despues"]["filas"]
        valor_esperado=None
        if primeras and len(primeras[0]) > 3:
            try:
                valor_esperado=float(str(primeras[0][3]).replace("%", "").replace(",", "").strip())
            except (ValueError,TypeError):
                valor_esperado=None
        categorias=["AS-IS · Histórico 2025"]
        valores=[ACTUAL[problema]]
        colores=[CYAN]
        etiquetas=["Dato informado"]
        if valor_esperado is not None and np.isfinite(valor_esperado):
            categorias.append("Resultado proyectado")
            valores.append(valor_esperado)
            colores.append(PURPLE)
            etiquetas.append("Proyección")
        categorias.append("TO-BE · Meta")
        valores.append(META[problema])
        colores.append(PINK)
        etiquetas.append("Objetivo")
        izquierda,derecha=st.columns([1.55,1], gap="large")
        with izquierda:
            fig=go.Figure(go.Bar(x=categorias,y=valores,marker_color=colores,
                                 text=[f"{v:.2f}%" for v in valores],textposition="outside",
                                 customdata=etiquetas,
                                 hovertemplate="%{x}<br>%{y:.2f}% · %{customdata}<extra></extra>"))
            fig.update_layout(title="Indicador inicial, esperado y meta",showlegend=False,
                              yaxis_title="Porcentaje (%)",yaxis_range=[0,max(110,max(valores)*1.14)])
            st.plotly_chart(formato_fig(fig,350),width="stretch")
        with derecha:
            st.markdown("#### ¿Qué explica la mejora?")
            st.write(det["accion"])
            var_meta=META[problema]-ACTUAL[problema]
            st.metric("Cambio objetivo (puntos porcentuales)",f"{var_meta:+.2f} pp")
            if valor_esperado is not None:
                st.metric("Variación esperada del modelo",f"{valor_esperado-ACTUAL[problema]:+.2f} pp")
            st.caption("La variación se calcula en puntos porcentuales, no como porcentaje relativo.")
        st.markdown("##### Matriz de trazabilidad del indicador")
        matriz=pd.DataFrame([
            {"Situación":"AS-IS", "Valor":f"{ACTUAL[problema]:.2f} %", "Naturaleza":"Registro histórico 2025", "Relación con la mejora":"Punto de partida"},
            {"Situación":"Resultado esperado", "Valor":f"{valor_esperado:.2f} %" if valor_esperado is not None else "No cuantificado", "Naturaleza":"Proyección sujeta a supuestos", "Relación con la mejora":det["accion"]},
            {"Situación":"TO-BE", "Valor":f"{META[problema]:.2f} %", "Naturaleza":"Meta definida por el grupo", "Relación con la mejora":"Objetivo a verificar después de implementar"},
        ])
        st.dataframe(matriz,hide_index=True,width="stretch")
        st.caption("AS-IS: registro consignado por el grupo; TO-BE: propuesta. Monte Carlo analiza incertidumbre económica, "
                   "no es una reproducción física del proceso. No se presentan proyecciones como mejoras ya ejecutadas.")

    with t2:
        st.markdown("#### Cómo se calcula cada iteración")
        st.markdown("**1.** Se toman los supuestos financieros y las distribuciones triangulares del Excel. "
                    "**2.** Se sortean parámetros para cada iteración. **3.** Se calcula inversión, beneficio mensual, "
                    "VAN, B/C y PRC. **4.** Se repite y se resumen los resultados.")
        try:
            mc=leer_montecarlo(fuente,problema)
            sens=leer_sensibilidad(fuente,problema)
        except Exception as exc:
            st.warning(f"No se pudieron leer las hojas de simulación del Excel: {exc}")
        else:
            indicadores=[col for col in mc.columns if col not in ("Iter.", "VAN (S/)", "B/C", "PRC (meses)")]
            candidatos=[]
            for nombre in indicadores:
                vals=pd.to_numeric(mc[nombre],errors="coerce")
                if vals.notna().sum()>100 and vals.nunique(dropna=True)>2:
                    candidatos.append(nombre)
            if candidatos:
                sel=st.selectbox("Explorar una variable de entrada o intermedia",candidatos,
                                 key=f"{prefijo}_{problema}_variable")
                serie=pd.to_numeric(mc[sel],errors="coerce")
                tabla=pd.DataFrame({"Variable":serie,"VAN (S/)":mc["VAN (S/)"]}).dropna()
                left,right=st.columns(2)
                with left:
                    scatter=px.scatter(tabla,x="Variable",y="VAN (S/)",opacity=.35,
                                       color_discrete_sequence=[CYAN])
                    scatter.add_hline(y=0,line_color=PINK,line_dash="dash")
                    scatter.update_layout(title=f"Relación entre {sel} y VAN",xaxis_title=sel)
                    st.plotly_chart(formato_fig(scatter,355),width="stretch")
                with right:
                    hist=px.histogram(tabla,x="Variable",nbins=36,color_discrete_sequence=[PURPLE])
                    hist.update_layout(title=f"Distribución de {sel}",xaxis_title=sel,yaxis_title="Frecuencia")
                    st.plotly_chart(formato_fig(hist,355),width="stretch")
                st.caption("Las curvas y puntos proceden de las filas guardadas en Excel; no se han generado nuevas "
                           "observaciones del proceso físico.")
            st.markdown("#### Sensibilidad reportada en el Excel")
            sf=px.bar(sens,x="Correlación",y="Variable",orientation="h",color="Correlación",
                      color_continuous_scale=[PINK,"#BBC3D2",CYAN],text="Correlación")
            sf.update_traces(texttemplate="%{x:.2f}",textposition="outside")
            sf.update_layout(coloraxis_showscale=False,yaxis_title="",xaxis_title="Correlación con el VAN")
            st.plotly_chart(formato_fig(sf,330),width="stretch")
            st.caption("Correlación = asociación en las iteraciones; no demuestra causalidad ni la mejora operativa.")

    with t3:
        st.markdown("#### De la propuesta al resultado económico")
        col1,col2=st.columns([1.1,1])
        with col1:
            st.markdown("**Beneficio previsto**")
            st.write(det["beneficio"])
            st.markdown("**Modelo financiero del informe**")
            st.markdown("`VAN = − inversión inicial + Σ beneficio neto del mes / (1 + COK)^mes`")
            st.caption("El COK se aplica por periodo según la plantilla original. Debe confirmarse su periodicidad.")
        with col2:
            k1,k2=st.columns(2)
            k1.metric("Inversión original",moneda(float(fila["Inversión (S/)"])))
            k2.metric("VAN base original",moneda(float(fila["VAN base (S/)"])))
            k3,k4=st.columns(2)
            k3.metric("VAN medio MC",moneda(float(fila["VAN medio MC (S/)"])))
            k4.metric("P(VAN > 0)",f"{100*float(fila['P(VAN>0)']):.1f} %")
        st.markdown("#### Distribución del VAN de las iteraciones originales")
        try:
            datos=leer_montecarlo(fuente,problema)
            fig=px.histogram(datos,x="VAN (S/)",nbins=42,color_discrete_sequence=[CYAN])
            fig.add_vline(x=0,line_color=PINK,line_dash="dash")
            fig.update_layout(xaxis_title="VAN (S/)",yaxis_title="Frecuencia")
            st.plotly_chart(formato_fig(fig,340),width="stretch")
        except Exception as exc:
            st.info(f"No se pudieron cargar los resultados guardados: {exc}")
        st.caption("La inversión, el VAN y las iteraciones provienen del Excel activo. Los beneficios "
                   "siguen siendo proyecciones sujetas a los supuestos descritos arriba.")


def mostrar_flujo_economico(problema, fila, prefijo="flujo"):
    """Reconstrucción transparente del flujo marginal constante reportado en el resumen."""
    st.markdown(f"#### Flujo de caja económico · {problema}")
    st.caption("Reconstituido con inversión inicial, beneficio mensual y horizonte del Excel activo. "
               "No incluye el préstamo ilustrativo de los flujos financieros del informe.")
    inversion=float(fila["Inversión (S/)"])
    beneficio=float(fila["Beneficio mensual (S/)"])
    meses=int(fila["Horizonte N (meses)"])
    tasa_usuario=st.number_input("COK aplicado por período (%)",min_value=0.0,max_value=40.0,
                     value=10.8513,step=0.1,format="%.4f",key=f"{prefijo}_{problema}_tasa")
    tasa=tasa_usuario/100
    periodos=list(range(meses+1))
    flujos=[-inversion]+[beneficio]*meses
    vp=[flujos[m]/(1+tasa)**m for m in periodos]
    acumulado=np.cumsum(vp)
    tabla=pd.DataFrame({"Mes":periodos,
        "Inversión inicial (S/)": [inversion]+[0]*meses,
        "Beneficio marginal (S/)": [0]+[beneficio]*meses,
        "Flujo neto (S/)":flujos,
        "Valor presente (S/)":vp,
        "VP acumulado (S/)":acumulado})
    a,b,c,d=st.columns(4)
    a.metric("Inversión inicial",moneda(inversion))
    b.metric("Beneficio marginal mensual",moneda(beneficio))
    c.metric("VAN recalculado",moneda(float(acumulado[-1])))
    d.metric("VAN reportado",moneda(float(fila["VAN base (S/)"])))
    izq,der=st.columns(2)
    with izq:
        fig=go.Figure(go.Bar(x=[f"Mes {m}" for m in periodos],y=flujos,
                             marker_color=[PINK if x<0 else CYAN for x in flujos],
                             text=[f"{x:,.0f}" for x in flujos],textposition="outside"))
        fig.add_hline(y=0,line_dash="dash",line_color=GREEN)
        fig.update_layout(title="Flujo marginal por mes",xaxis_title="Período",yaxis_title="Soles (S/)")
        st.plotly_chart(formato_fig(fig,360),width="stretch")
    with der:
        fig=go.Figure(go.Scatter(x=periodos,y=acumulado,mode="lines+markers",
                                 line=dict(color=PURPLE,width=3),marker=dict(size=10)))
        fig.add_hline(y=0,line_color=PINK,line_dash="dash")
        fig.update_layout(title="Valor presente acumulado",xaxis_title="Mes",yaxis_title="Soles (S/)")
        st.plotly_chart(formato_fig(fig,360),width="stretch")
    with st.expander("Ver el cuadro económico mes a mes",expanded=False):
        st.dataframe(tabla.style.format({c:"{:,.2f}" for c in tabla.columns if c!="Mes"}),
                     hide_index=True,width="stretch")
        st.download_button("Descargar flujo CSV",tabla.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"Flujo_economico_{problema}.csv",mime="text/csv",
                           key=f"{prefijo}_{problema}_descarga_flujo")
    diferencia=float(acumulado[-1])-float(fila["VAN base (S/)"])
    if abs(diferencia)>1:
        st.warning(f"El VAN reconstruido difiere del valor reportado en {moneda(abs(diferencia))}. "
                   "Revisa tasa, redondeos y otros conceptos del flujo original.")
    st.caption("Se aplica el COK por período igual que en la plantilla del grupo. Si el COK indicado fuese anual, "
               "debería convertirse antes de utilizar flujos mensuales. Los beneficios son supuestos económicos, "
               "no ahorros de caja ya comprobados.")

st.sidebar.title("AGUA FELICE")
st.sidebar.caption("INFORME FINAL · GRUPO 4")
pagina = st.sidebar.radio("MENÚ PRINCIPAL", [
    "Dashboard general", "Problemas", "Ishikawa interactivo", "Cronograma Gantt", "Modelo de simulación", "Simulación Monte Carlo", "Simulador interactivo",
    "Evaluación económica", "Conclusiones", "📂 Datos y Excel"
])
with st.sidebar.expander("Actualizar datos con Excel", expanded=False):
    st.caption("Carga el mismo formato de Excel del grupo. El cambio solo se aplica a tu sesión actual.")
    excel_subido=st.file_uploader("Selecciona el archivo .xlsx",type=["xlsx"],key="nuevo_excel")
    if excel_subido is not None:
        st.success(f"Archivo activo: {excel_subido.name}")
    else:
        st.caption("Archivo activo: datos_agua_felice.xlsx (original)")
    st.caption("El informe histórico y las conclusiones del Word siguen siendo los originales.")

st.title("INFORME FINAL | GRUPO 4")
st.caption("Dashboard interactivo de Ingeniería Industrial · Agua Felice")

if excel_subido is None and not EXCEL.exists():
    st.error("No encuentro datos_agua_felice.xlsx junto a app.py.")
    st.stop()
FUENTE=excel_subido.getvalue() if excel_subido is not None else str(EXCEL)
try:
    resumen=leer_resumen(FUENTE)
    for columna in ("Inversión (S/)", "Beneficio mensual (S/)", "VAN base (S/)", "VAN medio MC (S/)", "P(VAN>0)"):
        if columna not in resumen.columns:
            raise ValueError(f"Falta la columna: {columna}")
        resumen[columna]=pd.to_numeric(resumen[columna],errors="raise")
except Exception as error:
    st.error(f"El Excel no tiene el formato esperado de la hoja Resumen: {error}")
    st.info("Debe conservar las hojas Resumen y Sim P1...Sim P7 con la misma estructura del Excel original.")
    st.stop()

if pagina == "Dashboard general":
    a,b,c,d = st.columns(4)
    a.metric("Problemas analizados", "7")
    b.metric("Inversión total", moneda(resumen["Inversión (S/)"].sum()))
    c.metric("VAN base positivo", f"{int((resumen['VAN base (S/)'] > 0).sum())} de 7")
    d.metric("Simulaciones Monte Carlo", "14,000")

    st.markdown("### Indicadores de gestión")
    izq, der = st.columns([1.25, 1])
    with izq:
        # Los cinco indicadores cuya mejora consiste en aumentar el porcentaje.
        claves = ["P1", "P2", "P3", "P5", "P7"]
        ind = pd.DataFrame({
            "Indicador": [NOMBRES[p] for p in claves],
            "Situación histórica": [ACTUAL[p] for p in claves],
            "Meta propuesta": [META[p] for p in claves],
        }).melt(id_vars="Indicador", var_name="Serie", value_name="Porcentaje")
        fig = px.bar(ind, y="Indicador", x="Porcentaje", color="Serie", barmode="group",
                     orientation="h", text="Porcentaje",
                     color_discrete_map={"Situación histórica":CYAN,"Meta propuesta":PINK})
        fig.update_traces(texttemplate="%{x:.1f}%",textposition="outside",textfont=dict(color=TEXT,size=11))
        fig.update_layout(title="Indicadores con objetivo de incremento",
                          xaxis_title="Porcentaje (%)",yaxis_title="",xaxis_range=[0,112],
                          legend_title_text="")
        fig.update_yaxes(categoryorder="array",categoryarray=list(reversed([NOMBRES[p] for p in claves])))
        st.plotly_chart(formato_fig(fig,370),width="stretch")
    with der:
        st.markdown("##### Indicadores que buscamos reducir")
        c4,c6=st.columns(2)
        for col, clave, maximo, titulo in [(c4,"P4",40,"Rotura de stock"),
                                            (c6,"P6",2.5,"Scrap")]:
            with col:
                fig=go.Figure(go.Indicator(
                    mode="gauge+number",value=ACTUAL[clave],
                    number={"suffix":" %","font":{"size":25,"color":TEXT},"valueformat":".2f"},
                    title={"text":titulo,"font":{"size":14,"color":TEXT}},
                    gauge={"axis":{"range":[0,maximo],"tickcolor":MUTED,"tickfont":{"color":MUTED}},
                           "bar":{"color":CYAN,"thickness":0.28},
                           "bgcolor":PANEL,"borderwidth":0,
                           "threshold":{"line":{"color":PINK,"width":6},"thickness":0.85,"value":META[clave]}}
                ))
                fig.update_layout(margin=dict(l=15,r=15,t=46,b=10),height=260)
                st.plotly_chart(formato_fig(fig,260),width="stretch")
                st.caption(f"Meta propuesta: {META[clave]:.2f} %")

    izq, der = st.columns([1,1])
    with izq:
        st.markdown("### Distribución de la inversión")
        fig=px.pie(resumen,names="Problema",values="Inversión (S/)",hole=.68,
                   color_discrete_sequence=[CYAN,PINK,GREEN,PURPLE,"#FDBA74","#59A7F8","#F7D96B"])
        fig.update_traces(textinfo="percent+label",textposition="outside",textfont=dict(color=TEXT,size=12))
        fig.update_layout(showlegend=False,annotations=[dict(text=moneda(resumen["Inversión (S/)"].sum()),x=.5,y=.5,
                            font=dict(color=TEXT,size=16),showarrow=False)])
        st.plotly_chart(formato_fig(fig,345),width="stretch")
    with der:
        st.markdown("### Riesgo de rentabilidad")
        r = resumen[["Problema","P(VAN>0)"]].copy()
        r["Probabilidad (%)"] = r["P(VAN>0)"]*100
        r["Color"] = np.where(r["Probabilidad (%)"]>50,"Positiva","Sin VAN positivo")
        fig=px.bar(r,x="Problema",y="Probabilidad (%)",color="Color",
                   text="Probabilidad (%)",color_discrete_map={"Positiva":GREEN,"Sin VAN positivo":PINK})
        fig.update_traces(texttemplate="%{y:.1f}%",textposition="outside",
                          textfont=dict(color=TEXT,size=12))
        fig.update_layout(title="Probabilidad de VAN positivo · Monte Carlo",
                          yaxis_range=[0,110],yaxis_title="Porcentaje (%)",xaxis_title="",showlegend=False)
        st.plotly_chart(formato_fig(fig,345),width="stretch")

    st.markdown("### Evaluación económica comparativa")
    izq,der=st.columns([1.25,1])
    with izq:
        ven = resumen[["Problema","VAN base (S/)","VAN medio MC (S/)"]].melt(
            id_vars="Problema",var_name="Cálculo",value_name="VAN")
        fig = px.bar(ven,x="Problema",y="VAN",color="Cálculo",barmode="group",
                     color_discrete_map={"VAN base (S/)":CYAN,"VAN medio MC (S/)":PURPLE})
        fig.add_hline(y=0,line_dash="dash",line_color=PINK)
        fig.update_layout(title="VAN del caso base frente al VAN promedio simulado",
                          yaxis_title="Soles (S/)",xaxis_title="",legend_title_text="")
        st.plotly_chart(formato_fig(fig,370),width="stretch")
    with der:
        fig=px.bar(resumen,x="Problema",y="B/C",color="B/C",
                   color_continuous_scale=[PINK,PURPLE,CYAN])
        fig.add_hline(y=1,line_color="#F7D96B",line_dash="dash",
                      annotation_text="Referencia B/C = 1",annotation_font_color=TEXT)
        fig.update_layout(title="Relación beneficio/costo por alternativa",
                          coloraxis_showscale=False,xaxis_title="",yaxis_title="B/C (escala logarítmica)")
        fig.update_yaxes(type="log", range=[-0.5, 2.7])
        st.plotly_chart(formato_fig(fig,370),width="stretch")
    st.markdown("### Detalle de propuestas con VAN de menor escala")
    detalle = resumen.loc[resumen["Problema"].isin(["P3", "P4", "P6", "P7"]),
                          ["Problema", "VAN base (S/)", "VAN medio MC (S/)"]].melt(
                          id_vars="Problema",var_name="Escenario",value_name="VAN")
    fig=px.bar(detalle,x="Problema",y="VAN",color="Escenario",barmode="group",
               color_discrete_map={"VAN base (S/)":CYAN,"VAN medio MC (S/)":PURPLE},
               text="VAN")
    fig.add_hline(y=0,line_color=PINK,line_dash="dash")
    fig.update_traces(texttemplate="%{y:,.0f}",textposition="outside",textfont=dict(color=TEXT))
    fig.update_layout(title="VAN P3, P4, P6 y P7 · ampliación visual",xaxis_title="",yaxis_title="Soles (S/)")
    st.plotly_chart(formato_fig(fig,325),width="stretch")
    st.caption("Fuente: informe final y hoja Resumen del Excel del grupo. Las cifras de Monte Carlo corresponden a 2,000 iteraciones por problema; los valores históricos y las metas no son resultados simulados.")

elif pagina == "Problemas":
    problema = st.selectbox("Selecciona el problema que deseas analizar", list(NOMBRES),
                           format_func=lambda p: f"{p} · {NOMBRES[p]}")
    fila = resumen.loc[resumen["Problema"] == problema].iloc[0]
    inf = DATOS_INFORME[problema]
    st.subheader(f"{problema} · {NOMBRES[problema]}")
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Indicador histórico · 2025",f"{ACTUAL[problema]:.2f} %")
    k2.metric("Meta propuesta",f"{META[problema]:.2f} %")
    k3.metric("Inversión estimada",moneda(fila["Inversión (S/)"]))
    k4.metric("VAN del caso base",moneda(fila["VAN base (S/)"]))

    opciones = st.tabs(["Diagnóstico", "Análisis de causas", "Metodología y restricciones",
                        "Indicadores y AS-IS / TO-BE", "Plan de implementación", "Evaluación económica", "Monte Carlo", "Simulador interactivo", "Modelo de simulación", "Flujo de caja"])

    with opciones[0]:
        izq,der=st.columns([1.6,1])
        with izq:
            hist = pd.DataFrame({"Mes":list(range(1,13)),"Registro histórico (%)":inf["meses"]})
            fig=px.line(hist,x="Mes",y="Registro histórico (%)",markers=True,
                        color_discrete_sequence=[CYAN])
            fig.add_hline(y=META[problema],line_color=PINK,line_dash="dash",
                          annotation_text=f"Meta: {META[problema]:.2f}%",annotation_font_color=TEXT)
            fig.update_traces(line_width=3,marker_size=8)
            fig.update_layout(title=f"Evolución mensual del indicador · 2025 (Tabla {inf['tabla_mensual']})",
                              xaxis=dict(tickmode="array",tickvals=list(range(1,13))),
                              yaxis_title="Porcentaje (%)")
            st.plotly_chart(formato_fig(fig,385),width="stretch")
        with der:
            labels=["Situación histórica","Meta propuesta"]
            fig=go.Figure(go.Bar(x=labels,y=[ACTUAL[problema],META[problema]],
                                 marker_color=[CYAN,PINK],
                                 text=[f"{ACTUAL[problema]:.2f}%",f"{META[problema]:.2f}%"],
                                 textposition="outside"))
            fig.update_layout(title="Comparación del indicador",yaxis_title="Porcentaje (%)",
                              yaxis_range=[0,max(110,1.16*max(ACTUAL[problema],META[problema]))])
            st.plotly_chart(formato_fig(fig,385),width="stretch")
        st.caption(f"Fuente: registro mensual 2025, Tabla {inf['tabla_mensual']} del informe final. "
                   "La meta corresponde al diseño de mejora, no a un resultado histórico.")

    with opciones[1]:
        mostrar_ishikawa(problema,"pestana")
        st.divider()
        izq,der=st.columns([1.2,1])
        with izq:
            causas=pd.DataFrame(inf["prioridades"])
            fig=px.bar(causas.sort_values("prioridad"), x="prioridad",y="categoria",orientation="h",
                       text="prioridad",color="nivel",
                       color_discrete_map={"Alta":PINK,"Media":CYAN,"Baja":PURPLE})
            fig.update_layout(title=f"Priorización de causas (Tabla {inf['tabla_prioridades']})",
                              xaxis_title="Frecuencia × impacto",yaxis_title="",legend_title_text="Nivel")
            fig.update_traces(textposition="outside",textfont=dict(color=TEXT))
            st.plotly_chart(formato_fig(fig,370),width="stretch")
        with der:
            st.markdown("#### Principales causas del diagnóstico")
            for item in sorted(inf["prioridades"],key=lambda c:-c["prioridad"])[:3]:
                st.markdown(f"**{item['categoria']}** · prioridad {item['prioridad']} · {item['nivel']}")
            st.caption("Priorización construida por el grupo; no es un ranking estadístico independiente.")
        with st.expander("Consultar los 5 porqués y la evidencia de cada categoría"):
            for c in inf["causas"]:
                st.markdown(f"**{c['categoria']}**")
                st.write(c["cadena"])
                st.caption(f"Evidencia: {c['evidencia']}")
                st.divider()
        st.caption(f"Fuente: tablas {inf['tabla_prioridades']} y {inf['tabla_causas']} del informe final.")

    with opciones[2]:
        st.markdown(f"#### Metodología seleccionada: {fila['Alternativa']}")
        c1,c2=st.columns([1.1,1])
        with c1:
            alternatives=inf["alternativas"]
            st.markdown("##### Alternativas evaluadas")
            for i,a in enumerate(alternatives,1):
                with st.expander(f"{i}. {a['nombre']}",expanded=(i==1)):
                    st.write(a["descripcion"])
        with c2:
            score=inf["restricciones"]
            head=score["cabeceras"]
            rows=score["filas"]
            df=pd.DataFrame(rows,columns=head)
            numeric=[c for c in df.columns if "20%" in c]
            for c in numeric:df[c]=pd.to_numeric(df[c],errors="coerce")
            score_col=[c for c in df.columns if "Puntaje" in c or "puntaje" in c]
            if score_col:df[score_col[0]]=pd.to_numeric(df[score_col[0]],errors="coerce")
            fig=px.bar(df.sort_values(score_col[0] if score_col else numeric[0]),
                       x=score_col[0] if score_col else numeric[0],y=df.columns[0],orientation="h",
                       text=score_col[0] if score_col else numeric[0],
                       color_discrete_sequence=[GREEN])
            fig.update_traces(textposition="outside",textfont=dict(color=TEXT))
            fig.update_layout(title="Puntaje ponderado de las alternativas",xaxis_title="Puntaje sobre 5",yaxis_title="",
                              xaxis_range=[0,5.4])
            st.plotly_chart(formato_fig(fig,370),width="stretch")
        st.markdown("##### Matriz de restricciones")
        st.dataframe(df,hide_index=True,width="stretch")
        st.caption(f"Fuente: Tabla {inf['tabla_restricciones']} del informe. "
                   "Las cinco restricciones se ponderan con 20% cada una.")

    with opciones[3]:
        first=inf["antes_despues"]["filas"][0]
        try:
            esperado=float(str(first[3]).replace('%','').strip())
        except (ValueError,IndexError):
            esperado=None
        shown=["Histórico 2025","Meta del proyecto"]
        vals=[ACTUAL[problema],META[problema]]
        colors=[CYAN,PINK]
        if esperado is not None:
            shown.insert(1,"Esperado en el modelo")
            vals.insert(1,esperado)
            colors.insert(1,PURPLE)
        fig=go.Figure(go.Bar(x=shown,y=vals,marker_color=colors,
                             text=[f"{v:.2f}%" for v in vals],textposition="outside"))
        fig.update_layout(title="AS-IS · escenario del modelo · TO-BE propuesto",
                          yaxis_title="Porcentaje (%)",yaxis_range=[0,max(110,1.17*max(vals))])
        st.plotly_chart(formato_fig(fig,370),width="stretch")
        st.caption("Histórico: registros de 2025 · Esperado: proyección del modelo · Meta: propuesta del grupo")
        st.markdown("##### Fórmulas e indicadores")
        for f in inf["formulas"]:
            with st.expander(f["indicador"]+" · "+f["valor"],expanded=(f==inf["formulas"][0])):
                st.markdown("**Fórmula**")
                st.markdown(f'<div class="formula-card">{escape(f["formula"])}</div>', unsafe_allow_html=True)
                st.caption("Cálculo reportado: "+f["calculo"])
        with st.expander("Ver indicadores de proceso del informe"):
            proc=inf["procesos"]
            st.dataframe(pd.DataFrame(proc["filas"],columns=proc["cabeceras"]),
                         hide_index=True,width="stretch")
        with st.expander("Comparación completa del escenario propuesto"):
            comp=inf["antes_despues"]
            st.dataframe(pd.DataFrame(comp["filas"],columns=comp["cabeceras"]),
                         hide_index=True,width="stretch")

    with opciones[4]:
        mostrar_gantt(problema)

    with opciones[5]:
        ec1,ec2,ec3,ec4=st.columns(4)
        ec1.metric("VAN base",moneda(fila["VAN base (S/)"]))
        ec2.metric("VAN medio Monte Carlo",moneda(fila["VAN medio MC (S/)"]))
        ec3.metric("Relación B/C",f"{float(fila['B/C']):.2f}")
        ec4.metric("Prob. VAN positivo",f"{float(fila['P(VAN>0)'])*100:.1f}%")
        e1, e2, e3 = st.columns(3)
        tir = float(fila["TIR por periodo"])
        e1.metric("TIR por periodo", ">1,000 %" if tir > 10 else f"{tir*100:,.1f} %")
        e2.metric("Periodo de recuperación", (
            f"{float(fila['PRC (meses)']):.2f} meses" if isinstance(fila["PRC (meses)"], (int,float,np.integer,np.floating))
            else str(fila["PRC (meses)"])))
        e3.metric("Horizonte de evaluación", f"{int(fila['Horizonte N (meses)'])} meses")
        if fila["VAN base (S/)"]<0:
            st.caption("Decisión del caso base: no viable económicamente en el horizonte evaluado.")
        izq,der=st.columns(2)
        with izq:
            fig=go.Figure(go.Bar(x=["VAN base","VAN medio MC"],
                                 y=[fila["VAN base (S/)"],fila["VAN medio MC (S/)"]],
                                 marker_color=[CYAN,PURPLE],
                                 text=[moneda(fila["VAN base (S/)"]),moneda(fila["VAN medio MC (S/)"])],
                                 textposition="outside"))
            fig.add_hline(y=0,line_color=PINK,line_dash="dash")
            fig.update_layout(title="Comparación de VAN",yaxis_title="Soles (S/)")
            st.plotly_chart(formato_fig(fig,330),width="stretch")
        with der:
            fig=go.Figure(go.Indicator(mode="gauge+number",value=float(fila["P(VAN>0)"])*100,
                                      number=dict(suffix=" %",font=dict(color=TEXT)),
                                      gauge=dict(axis=dict(range=[0,100]),bar=dict(color=GREEN),
                                                 bgcolor=PANEL)))
            fig.update_layout(title=dict(text="Probabilidad simulada de VAN positivo"))
            st.plotly_chart(formato_fig(fig,330),width="stretch")
        st.caption("Fuente: hoja Resumen del Excel final; Monte Carlo se analiza en detalle desde "
                   "el menú lateral «Simulación Monte Carlo».")

    with opciones[6]:
        st.markdown(f"#### Simulación Monte Carlo · {problema}")
        try:
            mc_p = leer_montecarlo(FUENTE, problema)
            sens_p = leer_sensibilidad(FUENTE, problema)
        except Exception as error:
            st.error(f"No se pudo leer la simulación de {problema}: {error}")
        else:
            van_p = mc_p["VAN (S/)"]
            x1,x2,x3,x4 = st.columns(4)
            x1.metric("Iteraciones", f"{len(mc_p):,}")
            x2.metric("VAN medio", moneda(van_p.mean()))
            x3.metric("Prob. VAN positivo", f"{van_p.gt(0).mean()*100:.1f} %")
            x4.metric("Intervalo P5–P95", f"{van_p.quantile(.05):,.0f} / {van_p.quantile(.95):,.0f}")
            izq,der = st.columns(2)
            with izq:
                fig=px.histogram(mc_p,x="VAN (S/)",nbins=42,color_discrete_sequence=[CYAN])
                fig.add_vline(x=0,line_color=PINK,line_dash="dash")
                fig.update_layout(title="Distribución de VAN · 2,000 iteraciones",yaxis_title="Frecuencia")
                st.plotly_chart(formato_fig(fig,335),width="stretch")
            with der:
                fig=px.bar(sens_p,x="Correlación",y="Variable",orientation="h",
                           color="Correlación",color_continuous_scale=[PINK,"#B3B8D0",CYAN])
                fig.update_layout(title="Sensibilidad del VAN",yaxis_title="",coloraxis_showscale=False)
                st.plotly_chart(formato_fig(fig,335),width="stretch")
            with st.expander("Consultar y descargar las 2,000 iteraciones"):
                st.dataframe(mc_p,hide_index=True,width="stretch",height=300)
                st.download_button("Descargar CSV",mc_p.to_csv(index=False).encode("utf-8-sig"),
                                   file_name=f"MonteCarlo_{problema}_2000.csv",mime="text/csv",
                                   key=f"download_mc_tab_{problema}")

    with opciones[7]:
        try:
            mc_i=leer_montecarlo(FUENTE,problema)
        except Exception as error:
            st.error(f"No se pudo cargar la simulación de {problema}: {error}")
        else:
            mostrar_simulador(problema,fila,mc_i,"tab")

    with opciones[8]:
        mostrar_modelo_simulacion(problema,fila,inf,FUENTE,"tab_modelo")

    with opciones[9]:
        mostrar_flujo_economico(problema,fila,"tab_flujo")

elif pagina == "Ishikawa interactivo":
    st.subheader("Diagrama de Ishikawa · Análisis de causas")
    st.caption("Diagnóstico causa-efecto de los siete problemas del informe final.")
    seleccionado = st.selectbox("Selecciona el problema",list(NOMBRES),
                              format_func=lambda p:f"{p} · {NOMBRES[p]}",index=6,key="ishikawa_seleccion")
    mostrar_ishikawa(seleccionado,"menu")

elif pagina == "Cronograma Gantt":
    st.subheader("Cronograma Gantt de implementación")
    st.caption("Plan de trabajo del informe final · Grupo 4. Actividades y responsables por semana.")
    opcion_gantt=st.radio("Vista del cronograma",["Consolidado · 7 propuestas","Por problema"],horizontal=True)
    if opcion_gantt == "Consolidado · 7 propuestas":
        mostrar_gantt(None,vista_completa=True)
    else:
        seleccionado=st.selectbox("Problema a analizar",list(NOMBRES),format_func=lambda p:f"{p} · {NOMBRES[p]}",key="gantt_problema_menu")
        mostrar_gantt(seleccionado)

elif pagina == "Modelo de simulación":
    st.subheader("Modelo de simulación de las siete propuestas")
    seleccionado=st.selectbox("Selecciona un problema",list(NOMBRES),
                              format_func=lambda p:f"{p} · {NOMBRES[p]}",
                              index=6,key="modelo_problema_menu")
    fila_modelo=resumen.loc[resumen["Problema"]==seleccionado].iloc[0]
    mostrar_modelo_simulacion(seleccionado,fila_modelo,DATOS_INFORME[seleccionado],FUENTE,"menu_modelo")

elif pagina == "Simulación Monte Carlo":
    problema = st.selectbox("Simulación por problema", list(NOMBRES),
                           format_func=lambda p: f"{p} · {NOMBRES[p]}")
    try:
        mc=leer_montecarlo(FUENTE,problema)
        sens=leer_sensibilidad(FUENTE,problema)
    except Exception as error:
        st.error(f"No pude leer las 2,000 iteraciones de {problema}: {error}")
        st.stop()
    van=mc["VAN (S/)"]
    a,b,c,d=st.columns(4)
    a.metric("Iteraciones cargadas",f"{len(mc):,}")
    b.metric("VAN promedio",moneda(van.mean()))
    c.metric("Probabilidad VAN positivo",f"{(van.gt(0).mean()*100):.1f} %")
    d.metric("Percentil 5 / 95",f"{van.quantile(.05):,.0f} / {van.quantile(.95):,.0f}")
    izq,der=st.columns(2)
    with izq:
        fig=px.histogram(mc,x="VAN (S/)",nbins=42,color_discrete_sequence=[CYAN])
        fig.add_vline(x=0,line_color=PINK,line_dash="dash")
        fig.update_layout(title="Distribución del VAN de las 2,000 iteraciones",yaxis_title="Frecuencia")
        st.plotly_chart(formato_fig(fig),width="stretch")
    with der:
        fig=px.bar(sens,x="Correlación",y="Variable",orientation="h",
                   color="Correlación",color_continuous_scale=[PINK,"#B3B8D0",CYAN])
        fig.update_layout(title="Sensibilidad del VAN",yaxis_title="",coloraxis_showscale=False)
        st.plotly_chart(formato_fig(fig),width="stretch")
    fig=px.scatter(mc,x="Iter.",y="VAN (S/)",opacity=.5,
                   color_discrete_sequence=[GREEN])
    fig.add_hline(y=0,line_dash="dash",line_color=PINK)
    fig.update_layout(title="VAN de cada iteración",xaxis_title="Número de iteración")
    st.plotly_chart(formato_fig(fig,320),width="stretch")
    with st.expander("Ver y descargar las 2,000 iteraciones originales"):
        st.dataframe(mc,width="stretch",height=280)
        st.download_button("Descargar CSV",mc.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"MonteCarlo_{problema}_2000.csv",mime="text/csv")

elif pagina == "Evaluación económica":
    a,b,c=st.columns(3)
    a.metric("Inversión conjunta",moneda(resumen["Inversión (S/)"].sum()))
    b.metric("VAN positivo",f"{(resumen['VAN base (S/)'] > 0).sum()} de 7")
    c.metric("Problemas evaluados", "7")
    sel=st.selectbox("Ver alternativa",list(NOMBRES),
                     format_func=lambda p:f"{p} · {NOMBRES[p]}")
    f=resumen.loc[resumen["Problema"]==sel].iloc[0]
    a,b,c,d=st.columns(4)
    a.metric("VAN base",moneda(f["VAN base (S/)"]))
    b.metric("TIR por periodo",f"{f['TIR por periodo']*100:,.1f}%")
    c.metric("Relación B/C",f"{f['B/C']:.2f}")
    d.metric("Horizonte",f"{int(f['Horizonte N (meses)'])} meses")
    izq,der=st.columns(2)
    with izq:
        datos=resumen[["Problema","VAN base (S/)","VAN medio MC (S/)"]].melt(
            id_vars="Problema",var_name="Tipo",value_name="VAN")
        fig=px.bar(datos,x="Problema",y="VAN",color="Tipo",barmode="group",
                   color_discrete_sequence=[CYAN,PINK])
        st.plotly_chart(formato_fig(fig,360),width="stretch")
    with der:
        datos=resumen[["Problema","B/C"]]
        fig=px.bar(datos,x="Problema",y="B/C",color_discrete_sequence=[GREEN])
        fig.add_hline(y=1,line_color=PINK,line_dash="dash")
        st.plotly_chart(formato_fig(fig,360),width="stretch")
    st.dataframe(resumen[["Problema","Alternativa","Inversión (S/)","VAN base (S/)",
                          "TIR por periodo","B/C","Decisión"]],hide_index=True,width="stretch")

    st.divider()
    st.subheader("Flujos económicos por problema")
    elegido_flujo=st.selectbox("Selecciona el flujo económico",list(NOMBRES),
                               format_func=lambda p:f"{p} · {NOMBRES[p]}",key="flujo_global")
    fila_flujo=resumen.loc[resumen["Problema"]==elegido_flujo].iloc[0]
    mostrar_flujo_economico(elegido_flujo,fila_flujo,"global_flujo")

elif pagina == "Simulador interactivo":
    st.subheader("Laboratorio de escenarios · Monte Carlo")
    problema=st.selectbox("Selecciona un problema",list(NOMBRES),
                          format_func=lambda p: f"{p} · {NOMBRES[p]}")
    fila=resumen.loc[resumen["Problema"]==problema].iloc[0]
    try:
        mc=leer_montecarlo(FUENTE,problema)
    except Exception as error:
        st.error(f"No se pudieron cargar los datos de Monte Carlo de {problema}: {error}")
    else:
        mostrar_simulador(problema,fila,mc,"main")

elif pagina == "📂 Datos y Excel":
    st.subheader("Administración de datos")
    st.markdown("**Libro de Excel activo**")
    st.info(excel_subido.name if excel_subido is not None else "datos_agua_felice.xlsx · archivo base")
    st.caption("Si deseas actualizar costos, beneficios, VAN y simulaciones, utiliza el cargador del menú lateral. "
               "Solo acepta la estructura de hojas Resumen y Sim P1...Sim P7 del modelo del grupo. "
               "Guarda y recalcula las fórmulas de Excel antes de subirlo.")
    st.dataframe(resumen,hide_index=True,width="stretch")
    checks=[]
    for p in NOMBRES:
        try:
            datos_mc=leer_montecarlo(FUENTE,p)
            checks.append({"Problema":p,"Filas Monte Carlo":len(datos_mc),"Estado":"Correcto" if len(datos_mc)==2000 else "Revisar número de filas"})
        except Exception as error:
            checks.append({"Problema":p,"Filas Monte Carlo":0,"Estado":f"Error: {str(error)[:90]}"})
    st.markdown("##### Revisión de hojas de simulación")
    st.dataframe(pd.DataFrame(checks),hide_index=True,width="stretch")
    st.caption("Al subir un Excel nuevo cambian los datos económicos y las simulaciones del modelo. "
               "Las series históricas, causas, metas y conclusiones que se transcribieron del Word "
               "permanecen como referencia del informe final y no se reemplazan automáticamente.")

elif pagina == "Conclusiones":
    st.subheader("Conclusiones del informe final · Grupo 4")
    st.caption("Basadas en las 11 conclusiones del documento Word original. Los indicadores de las tarjetas económicas sí usan el Excel activo.")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Problemas evaluados","7")
    c2.metric("VAN base positivo",f"{int((resumen['VAN base (S/)']>0).sum())} de 7")
    c3.metric("Inversión total",moneda(float(resumen["Inversión (S/)"].sum())))
    c4.metric("Iteraciones originales","2,000 × 7")
    tab1,tab2,tab3=st.tabs(["Resumen visual","Por problema","Conclusiones y límites del informe"])
    with tab1:
        izq,der=st.columns([1.15,1])
        with izq:
            fig=px.bar(resumen,x="Problema",y="VAN base (S/)",
                       color="Decisión",text="VAN base (S/)",
                       color_discrete_sequence=[GREEN,PINK])
            fig.update_traces(texttemplate="%{y:,.0f}",textposition="outside")
            fig.add_hline(y=0,line_dash="dash",line_color=PINK)
            fig.update_layout(title="VAN de cada alternativa · caso base",yaxis_title="Soles (S/)",xaxis_title="")
            st.plotly_chart(formato_fig(fig,330),width="stretch")
        with der:
            positivos=int((resumen["VAN base (S/)"]>0).sum())
            fig=go.Figure(go.Pie(labels=["VAN positivo","VAN no positivo"],values=[positivos,7-positivos],
                                 hole=.65,marker=dict(colors=[GREEN,PINK]),textinfo="label+value"))
            fig.update_layout(title="Decisión económica reportada")
            st.plotly_chart(formato_fig(fig,330),width="stretch")
        st.markdown("##### Comparación de indicadores de gestión")
        rows=[]
        for p in NOMBRES:
            rows.append({"Problema":p,"Indicador":NOMBRES[p],"Histórico (%)":ACTUAL[p],
                         "Meta (%)":META[p],"Tendencia buscada":TIPO[p]})
        st.dataframe(pd.DataFrame(rows),hide_index=True,width="stretch")
    with tab2:
        problema=st.selectbox("Consultar conclusión",list(NOMBRES),
                              format_func=lambda p:f"{p} · {NOMBRES[p]}",key="conclusion_p")
        fila=resumen.loc[resumen["Problema"]==problema].iloc[0]
        x1,x2,x3=st.columns(3)
        x1.metric("VAN base",moneda(float(fila["VAN base (S/)"])))
        x2.metric("Inversión",moneda(float(fila["Inversión (S/)"])))
        x3.metric("Meta",f"{META[problema]:.2f} %")
        st.markdown("##### Conclusión del documento original")
        st.markdown(CONCLUSIONES_ORIGINALES[int(problema[1:])])
        st.caption("Texto literal de la conclusión correspondiente en el informe final. "
                   "Si cargas otro Excel, el texto original no cambia; las tarjetas sí.")
    with tab3:
        for titulo,indices in [("Visión general",[0,8,9]),
                               ("Problemas 1 a 4",[1,2,3,4]),
                               ("Problemas 5 a 7",[5,6,7]),
                               ("Limitaciones y seguimiento",[10])]:
            with st.expander(titulo,expanded=(titulo in ("Visión general", "Limitaciones y seguimiento"))):
                for i in indices:
                    st.markdown(CONCLUSIONES_ORIGINALES[i])
                    st.divider()
    st.caption("Fuente: Conclusiones, páginas 180–181 del Informe Final Agua Felice Grupo 4.")
