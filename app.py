"""Informe Final | Grupo 4 – Agua Felice.

Dashboard Streamlit basado en el Excel final del grupo.
Mantén datos_agua_felice.xlsx en la misma carpeta que este script.
"""
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Agua Felice | Informe Final – Grupo 4",
    page_icon="💧",
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
/* Mejorar visibilidad de los desplegables */
[data-testid="stExpander"] summary {
    background-color: #20314F !important;
    border-radius: 8px;
}


}""", unsafe_allow_html=True)

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

@st.cache_data(show_spinner=False)
def leer_resumen(ruta: str) -> pd.DataFrame:
    tabla = pd.read_excel(ruta, sheet_name="Resumen", header=2, nrows=7, engine="openpyxl")
    tabla = tabla.loc[tabla["Problema"].isin(NOMBRES)].copy()
    if len(tabla) != 7:
        raise ValueError("No se encontraron los siete problemas en la hoja Resumen.")
    return tabla

@st.cache_data(show_spinner=False)
def leer_montecarlo(ruta: str, problema: str) -> pd.DataFrame:
    # En las hojas Sim P1–Sim P7, los títulos están en la fila 30.
    tabla = pd.read_excel(ruta, sheet_name=f"Sim {problema}", header=29,
                          nrows=2000, usecols="A:N", engine="openpyxl")
    tabla["VAN (S/)"] = pd.to_numeric(tabla["VAN (S/)"], errors="coerce")
    tabla["Iter."] = pd.to_numeric(tabla["Iter."], errors="coerce")
    return tabla.dropna(subset=["Iter.", "VAN (S/)"]).copy()

@st.cache_data(show_spinner=False)
def leer_sensibilidad(ruta: str, problema: str) -> pd.DataFrame:
    tabla = pd.read_excel(ruta, sheet_name=f"Sim {problema}",
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

st.sidebar.title("💧 AGUA FELICE")
st.sidebar.caption("INFORME FINAL · GRUPO 4")
pagina = st.sidebar.radio("MENÚ PRINCIPAL", [
    "Dashboard general", "Problemas", "Simulación Monte Carlo",
    "Evaluación económica", "Conclusiones"
])

st.title("INFORME FINAL | GRUPO 4")
st.caption("Dashboard interactivo de Ingeniería Industrial · Agua Felice")

if not EXCEL.exists():
    st.error("No encuentro datos_agua_felice.xlsx junto a este archivo Python.")
    st.code(str(EXCEL.parent), language="text")
    st.stop()

try:
    resumen = leer_resumen(str(EXCEL))
except Exception as error:
    st.error(f"No pude leer la hoja Resumen del Excel: {error}")
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
        st.plotly_chart(formato_fig(fig,370),use_container_width=True)
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
                st.plotly_chart(formato_fig(fig,260),use_container_width=True)
                st.caption(f"Meta propuesta: {META[clave]:.2f} %")

    izq, der = st.columns([1,1])
    with izq:
        st.markdown("### Distribución de la inversión")
        fig=px.pie(resumen,names="Problema",values="Inversión (S/)",hole=.68,
                   color_discrete_sequence=[CYAN,PINK,GREEN,PURPLE,"#FDBA74","#59A7F8","#F7D96B"])
        fig.update_traces(textinfo="percent+label",textposition="outside",textfont=dict(color=TEXT,size=12))
        fig.update_layout(showlegend=False,annotations=[dict(text="S/ 10,357.93",x=.5,y=.5,
                            font=dict(color=TEXT,size=16),showarrow=False)])
        st.plotly_chart(formato_fig(fig,345),use_container_width=True)
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
        st.plotly_chart(formato_fig(fig,345),use_container_width=True)

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
        st.plotly_chart(formato_fig(fig,370),use_container_width=True)
    with der:
        fig=px.bar(resumen,x="Problema",y="B/C",color="B/C",
                   color_continuous_scale=[PINK,PURPLE,CYAN])
        fig.add_hline(y=1,line_color="#F7D96B",line_dash="dash",
                      annotation_text="Referencia B/C = 1",annotation_font_color=TEXT)
        fig.update_layout(title="Relación beneficio/costo por alternativa",
                          coloraxis_showscale=False,xaxis_title="",yaxis_title="B/C (escala logarítmica)")
        fig.update_yaxes(type="log", range=[-0.5, 2.7])
        st.plotly_chart(formato_fig(fig,370),use_container_width=True)
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
    st.plotly_chart(formato_fig(fig,325),use_container_width=True)
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

    opciones = st.tabs(["📊 Diagnóstico", "🔎 Análisis de causas", "🧩 Metodología y restricciones",
                        "🎯 Indicadores y AS-IS / TO-BE", "💰 Evaluación económica"])

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
            st.plotly_chart(formato_fig(fig,385),use_container_width=True)
        with der:
            labels=["Situación histórica","Meta propuesta"]
            fig=go.Figure(go.Bar(x=labels,y=[ACTUAL[problema],META[problema]],
                                 marker_color=[CYAN,PINK],
                                 text=[f"{ACTUAL[problema]:.2f}%",f"{META[problema]:.2f}%"],
                                 textposition="outside"))
            fig.update_layout(title="Comparación del indicador",yaxis_title="Porcentaje (%)",
                              yaxis_range=[0,max(110,1.16*max(ACTUAL[problema],META[problema]))])
            st.plotly_chart(formato_fig(fig,385),use_container_width=True)
        st.caption(f"Fuente: registro mensual 2025, Tabla {inf['tabla_mensual']} del informe final. "
                   "La meta corresponde al diseño de mejora, no a un resultado histórico.")

    with opciones[1]:
        izq,der=st.columns([1.2,1])
        with izq:
            causas=pd.DataFrame(inf["prioridades"])
            fig=px.bar(causas.sort_values("prioridad"), x="prioridad",y="categoria",orientation="h",
                       text="prioridad",color="nivel",
                       color_discrete_map={"Alta":PINK,"Media":CYAN,"Baja":PURPLE})
            fig.update_layout(title=f"Priorización de causas (Tabla {inf['tabla_prioridades']})",
                              xaxis_title="Frecuencia × impacto",yaxis_title="",legend_title_text="Nivel")
            fig.update_traces(textposition="outside",textfont=dict(color=TEXT))
            st.plotly_chart(formato_fig(fig,370),use_container_width=True)
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
            st.plotly_chart(formato_fig(fig,370),use_container_width=True)
        st.markdown("##### Matriz de restricciones")
        st.dataframe(df,hide_index=True,use_container_width=True)
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
        st.plotly_chart(formato_fig(fig,370),use_container_width=True)
        st.caption("Histórico = registros 2025; Meta = objetivo de la propuesta; "
                   "Esperado = resultado proyectado con supuestos de cumplimiento del modelo, no una observación física.")
        st.markdown("##### Fórmulas e indicadores")
        for f in inf["formulas"]:
            with st.expander(f["indicador"]+" · "+f["valor"],expanded=(f==inf["formulas"][0])):
                st.markdown("**Fórmula**")
                st.code(f["formula"],language=None)
                st.caption("Cálculo reportado: "+f["calculo"])
        with st.expander("Ver indicadores de proceso del informe"):
            proc=inf["procesos"]
            st.dataframe(pd.DataFrame(proc["filas"],columns=proc["cabeceras"]),
                         hide_index=True,use_container_width=True)
        with st.expander("Comparación completa del escenario propuesto"):
            comp=inf["antes_despues"]
            st.dataframe(pd.DataFrame(comp["filas"],columns=comp["cabeceras"]),
                         hide_index=True,use_container_width=True)

    with opciones[4]:
        ec1,ec2,ec3,ec4=st.columns(4)
        ec1.metric("VAN base",moneda(fila["VAN base (S/)"]))
        ec2.metric("VAN medio Monte Carlo",moneda(fila["VAN medio MC (S/)"]))
        ec3.metric("Relación B/C",f"{float(fila['B/C']):.2f}")
        ec4.metric("Prob. VAN positivo",f"{float(fila['P(VAN>0)'])*100:.1f}%")
        if fila["VAN base (S/)"]<0:
            st.caption("Decisión del caso base reportada por el grupo: no viable económicamente en el horizonte evaluado.")
        izq,der=st.columns(2)
        with izq:
            fig=go.Figure(go.Bar(x=["VAN base","VAN medio MC"],
                                 y=[fila["VAN base (S/)"],fila["VAN medio MC (S/)"]],
                                 marker_color=[CYAN,PURPLE],
                                 text=[moneda(fila["VAN base (S/)"]),moneda(fila["VAN medio MC (S/)"])],
                                 textposition="outside"))
            fig.add_hline(y=0,line_color=PINK,line_dash="dash")
            fig.update_layout(title="Comparación de VAN",yaxis_title="Soles (S/)")
            st.plotly_chart(formato_fig(fig,330),use_container_width=True)
        with der:
            fig=go.Figure(go.Indicator(mode="gauge+number",value=float(fila["P(VAN>0)"])*100,
                                      number=dict(suffix=" %",font=dict(color=TEXT)),
                                      gauge=dict(axis=dict(range=[0,100]),bar=dict(color=GREEN),
                                                 bgcolor=PANEL)))
            fig.update_layout(title=dict(text="Probabilidad simulada de VAN positivo"))
            st.plotly_chart(formato_fig(fig,330),use_container_width=True)
        st.caption("Fuente: hoja Resumen del Excel final; Monte Carlo se analiza en detalle desde "
                   "el menú lateral «Simulación Monte Carlo».")

elif pagina == "Simulación Monte Carlo":
    problema = st.selectbox("Simulación por problema", list(NOMBRES),
                           format_func=lambda p: f"{p} · {NOMBRES[p]}")
    try:
        mc=leer_montecarlo(str(EXCEL),problema)
        sens=leer_sensibilidad(str(EXCEL),problema)
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
        st.plotly_chart(formato_fig(fig),use_container_width=True)
    with der:
        fig=px.bar(sens,x="Correlación",y="Variable",orientation="h",
                   color="Correlación",color_continuous_scale=[PINK,"#B3B8D0",CYAN])
        fig.update_layout(title="Sensibilidad del VAN",yaxis_title="",coloraxis_showscale=False)
        st.plotly_chart(formato_fig(fig),use_container_width=True)
    fig=px.scatter(mc,x="Iter.",y="VAN (S/)",opacity=.5,
                   color_discrete_sequence=[GREEN])
    fig.add_hline(y=0,line_dash="dash",line_color=PINK)
    fig.update_layout(title="VAN de cada iteración",xaxis_title="Número de iteración")
    st.plotly_chart(formato_fig(fig,320),use_container_width=True)
    with st.expander("Ver y descargar las 2,000 iteraciones originales"):
        st.dataframe(mc,use_container_width=True,height=280)
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
        st.plotly_chart(formato_fig(fig,360),use_container_width=True)
    with der:
        datos=resumen[["Problema","B/C"]]
        fig=px.bar(datos,x="Problema",y="B/C",color_discrete_sequence=[GREEN])
        fig.add_hline(y=1,line_color=PINK,line_dash="dash")
        st.plotly_chart(formato_fig(fig,360),use_container_width=True)
    st.dataframe(resumen[["Problema","Alternativa","Inversión (S/)","VAN base (S/)",
                          "TIR por periodo","B/C","Decisión"]],hide_index=True,use_container_width=True)

elif pagina == "Conclusiones":
    st.subheader("Síntesis de resultados")
    st.metric("Alternativas con VAN base positivo",f"{(resumen['VAN base (S/)']>0).sum()} de 7")
    for _,fila in resumen.iterrows():
        color="🟢" if fila["VAN base (S/)"]>0 else "🔴"
        st.markdown(f"{color} **{fila['Problema']} · {NOMBRES[fila['Problema']]}:** "
                    f"{fila['Alternativa']}; VAN base {moneda(fila['VAN base (S/)'])}.")
    st.caption("Valores procedentes del archivo de evaluación económica original del grupo. Las metas son propuestas de mejora.")
