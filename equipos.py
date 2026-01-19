import csv
import pandas as pd
import glob
import os
from itertools import combinations
from difflib import SequenceMatcher

def similitud_nombres(nombre1, nombre2):
	return SequenceMatcher(None, nombre1.upper(), nombre2.upper()).ratio()

def normalizar_nombre(nombre, todos_los_nombres, umbral=0.85):
	nombre = nombre.strip().upper()
	
	for nombre_existente in todos_los_nombres:
		if similitud_nombres(nombre, nombre_existente) >= umbral:
			return nombre_existente
	
	return nombre

def excel_a_csv(archivo_excel, archivo_csv):
	extension = archivo_excel.split('.')[-1].lower()
	
	if extension in ['xlsx', 'xls', 'xltx']:
		df = pd.read_excel(archivo_excel, engine='openpyxl')
	elif extension == 'ods':
		df = pd.read_excel(archivo_excel, engine='odf')
	else:
		raise ValueError(f"Formato no soportado: {extension}")
	
	df.to_csv(archivo_csv, index=False, encoding='utf-8')
	print(f"Convertido: {archivo_excel} → {archivo_csv}")

def leer_csv(archivo):
	with open(archivo, 'r', encoding='utf-8') as f:
		reader = csv.reader(f)
		headers = next(reader)
		
		idx_cambio = None
		for i, header in enumerate(headers):
			if 'NO PREFERENCIA' in header.upper():
				idx_cambio = i
				break
		
		if idx_cambio is None:
			raise ValueError("No se encontró la columna 'NO PREFERENCIA'")
		
		col_nombre = 0
		cols_preferencias = list(range(1, idx_cambio))
		cols_no_preferencias = list(range(idx_cambio, len(headers)))
		
		return headers, cols_preferencias, cols_no_preferencias

def construir_grupos(archivo_csv, cols_preferencias, cols_no_preferencias):
	grupos = []
	grupo_indiferente = {
		'nombre': 'INDIFERENTE',
		'miembros': [],
		'datos_miembros': {}
	}
	
	todos_los_nombres = []
	
	with open(archivo_csv, 'r', encoding='utf-8') as f:
		reader = csv.reader(f)
		next(reader)
		
		for fila in reader:
			nombre_raw = fila[0].strip()
			nombre = normalizar_nombre(nombre_raw, todos_los_nombres)
			
			if nombre not in todos_los_nombres:
				todos_los_nombres.append(nombre)
			
			preferencias_raw = [fila[i].strip() for i in cols_preferencias if fila[i].strip() and fila[i].strip().upper() != 'CUALQUIERA']
			preferencias = [normalizar_nombre(p, todos_los_nombres) for p in preferencias_raw]
			
			for p in preferencias:
				if p not in todos_los_nombres:
					todos_los_nombres.append(p)
			
			no_preferencias_raw = [fila[i].strip() for i in cols_no_preferencias if fila[i].strip() and fila[i].strip().upper() != 'CUALQUIERA']
			no_preferencias = [normalizar_nombre(np, todos_los_nombres) for np in no_preferencias_raw]
			
			print(f"\nProcesando: {nombre}")
			print(f"  Preferencias: {preferencias}")
			print(f"  No preferencias: {no_preferencias}")
			
			grupo_encontrado = None
			
			for grupo in grupos:
				if nombre in grupo['miembros']:
					grupo_encontrado = grupo
					print(f"  → {nombre} ya está en un grupo")
					break
			
			if not grupo_encontrado:
				for grupo in grupos:
					for miembro in grupo['miembros']:
						if miembro in grupo['datos_miembros']:
							datos_miembro = grupo['datos_miembros'][miembro]
							if nombre in datos_miembro['preferencias']:
								grupo_encontrado = grupo
								print(f"  → {miembro} quiere debatir con {nombre}")
								break
					if grupo_encontrado:
						break
			
			if not grupo_encontrado and preferencias:
				for grupo in grupos:
					for pref in preferencias:
						if pref in grupo['miembros']:
							grupo_encontrado = grupo
							print(f"  → {nombre} quiere debatir con {pref} que está en un grupo")
							break
					if grupo_encontrado:
						break
			
			if grupo_encontrado:
				print(f"  → Se une al grupo existente")
				if nombre not in grupo_encontrado['miembros']:
					grupo_encontrado['miembros'].append(nombre)
				
				for pref in preferencias:
					if pref not in grupo_encontrado['miembros']:
						grupo_encontrado['miembros'].append(pref)
				
				grupo_encontrado['datos_miembros'][nombre] = {
					'preferencias': preferencias,
					'no_preferencias': no_preferencias
				}
			elif not preferencias:
				print(f"  → Sin preferencias específicas, va a INDIFERENTE")
				grupo_indiferente['miembros'].append(nombre)
				grupo_indiferente['datos_miembros'][nombre] = {
					'preferencias': preferencias,
					'no_preferencias': no_preferencias
				}
			else:
				print(f"  → Crea nuevo grupo")
				miembros_iniciales = [nombre] + preferencias
				miembros_unicos = list(dict.fromkeys(miembros_iniciales))
				
				nuevo_grupo = {
					'miembros': miembros_unicos,
					'datos_miembros': {
						nombre: {
							'preferencias': preferencias,
							'no_preferencias': no_preferencias
						}
					}
				}
				grupos.append(nuevo_grupo)
	
	if grupo_indiferente['miembros']:
		grupos.append(grupo_indiferente)
	
	return grupos

def calcular_puntuacion_equipo(equipo, datos_miembros):
	puntuacion = 0

	for persona in equipo:
		if persona not in datos_miembros:
			continue
		
		datos = datos_miembros[persona]
		
		for otro in equipo:
			if otro == persona:
				continue
			
			if otro in datos['preferencias']:
				puntuacion += 1
			
			if otro in datos['no_preferencias']:
				puntuacion -= 3
	
	return puntuacion

def generar_equipos(grupos):
	resultados = []
	
	for idx, grupo in enumerate(grupos):
		print(f"\n{'='*50}")
		if 'nombre' in grupo and grupo['nombre'] == 'INDIFERENTE':
			print(f"Procesando GRUPO INDIFERENTE")
			nombre_grupo = 'INDIFERENTE'
		else:
			print(f"Procesando GRUPO {idx + 1}")
			nombre_grupo = f"GRUPO {idx + 1}"
		
		print(f"Miembros: {len(grupo['miembros'])}")
		
		if len(grupo['miembros']) < 4:
			print(f"  ADVERTENCIA: Grupo muy pequeño, se necesitan al menos 4 miembros")
			continue
		
		equipos_con_puntuacion = []
		
		for equipo in combinations(grupo['miembros'], 4):
			puntuacion = calcular_puntuacion_equipo(equipo, grupo['datos_miembros'])
			equipos_con_puntuacion.append((list(equipo), puntuacion))
		
		equipos_con_puntuacion.sort(key=lambda x: x[1], reverse=True)
		
		top_5 = equipos_con_puntuacion[:5]
		
		print(f"\nTOP 5 EQUIPOS:")
		for i, (equipo, punt) in enumerate(top_5, 1):
			print(f"  {i}. {equipo} - Puntuacion: {punt}")
		
		resultados.append({
			'grupo': nombre_grupo,
			'top_equipos': top_5
		})
	
	return resultados

def main():
	carpeta = 'meter_excel_aqui'
	
	if not os.path.exists(carpeta):
		print(f"Error: No existe la carpeta '{carpeta}'")
		return
	
	archivos = glob.glob(os.path.join(carpeta, '*'))
	archivos = [f for f in archivos if os.path.isfile(f)]
	
	if not archivos:
		print(f"Error: No se encontró ningún archivo en '{carpeta}'")
		return
	
	archivo_excel = archivos[0]
	archivo_csv = 'debates.csv'
	
	print(f"Usando archivo: {archivo_excel}")
	
	excel_a_csv(archivo_excel, archivo_csv)
	
	headers, cols_pref, cols_no_pref = leer_csv(archivo_csv)
	
	grupos = construir_grupos(archivo_csv, cols_pref, cols_no_pref)
	
	print(f"\n\n{'='*50}")
	print(f"TOTAL GRUPOS CREADOS: {len(grupos)}")
	print(f"{'='*50}\n")
	
	for i, grupo in enumerate(grupos, 1):
		if 'nombre' in grupo and grupo['nombre'] == 'INDIFERENTE':
			print(f"GRUPO INDIFERENTE: {grupo['miembros']}")
		else:
			print(f"GRUPO {i}: {grupo['miembros']}")
	
	resultados = generar_equipos(grupos)

if __name__ == '__main__':
	main()
