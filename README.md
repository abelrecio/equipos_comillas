# Generador de Equipos de Debate

## Formato del Excel

El archivo Excel debe tener las siguientes columnas en este orden:

1. **NOMBRE**: Nombre del participante
2. **PREFERENCIA** (5 columnas): Personas con las que SÍ quiere debatir
3. **NO PREFERENCIA** (3 columnas): Personas con las que NO quiere debatir

Ejemplo:

| NOMBRE | PREFERENCIA | PREFERENCIA | PREFERENCIA | PREFERENCIA | PREFERENCIA | NO PREFERENCIA | NO PREFERENCIA | NO PREFERENCIA |
|--------|-------------|-------------|-------------|-------------|-------------|----------------|----------------|----------------|
| MANU   | LARA        | PEPE        | MINGO       | TERE        | SALINO      | FARUNO         | PESTO          | CALA           |
| PEPE   | JOSELUIS    | RABINO      | CUALQUIERA  | CUALQUIERA  | CUALQUIERA  | CUALQUIERA     | CUALQUIERA     | CUALQUIERA     |

- Si no tiene preferencia específica, escribir **CUALQUIERA**
- Los nombres deben ser consistentes (evitar errores tipográficos)

## Cómo usar el programa

1. **Abrir el programa**: Doble clic en `GeneradorEquipos.exe`

2. **Importar Excel**: 
   - Click en "Importar Excel"
   - Seleccionar el archivo con los datos
   - El programa mostrará los grupos formados

3. **Exportar resultados**:
   - Click en "Exportar Resultados Excel"
   - Elegir dónde guardar el archivo
   - Se creará un Excel con:
     - Todos los miembros de cada grupo
     - Top 5 equipos de 4 personas con puntuación

## Reglas de formación de grupos

- Personas se unen al mismo grupo si aparecen en las preferencias mutuas
- "CUALQUIERA" significa sin preferencia → va a grupo INDIFERENTE
- Puntuación: +1 por cada preferencia, -3 por cada no-preferencia
