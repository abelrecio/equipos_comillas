import tkinter as tk
from tkinter import filedialog, scrolledtext
import sys
from io import StringIO
import pandas as pd
import os
from equipos import excel_a_csv, leer_csv, construir_grupos, generar_equipos

class Aplicacion:
	def __init__(self, root):
		self.root = root
		self.root.title("Generador de Equipos de Debate")
		self.root.geometry("800x650")
		
		btn_importar = tk.Button(root, text="Importar Excel", command=self.importar_excel, height=2, width=20)
		btn_importar.pack(pady=10)
		
		btn_exportar = tk.Button(root, text="Exportar Resultados Excel", command=self.exportar_excel, height=2, width=20)
		btn_exportar.pack(pady=5)
		
		self.texto = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=90, height=28)
		self.texto.pack(pady=10, padx=10, expand=True, fill='both')
		
		self.resultados = None
		self.grupos = None
		
	def importar_excel(self):
		archivo_excel = filedialog.askopenfilename(
			title="Seleccionar archivo Excel",
			filetypes=[("Excel files", "*.xlsx *.xls *.xltx *.ods"), ("Todos", "*.*")]
		)
		
		if not archivo_excel:
			return
		
		self.texto.delete(1.0, tk.END)
		
		old_stdout = sys.stdout
		sys.stdout = StringIO()
		
		try:
			archivo_csv = 'debates.csv'
			
			excel_a_csv(archivo_excel, archivo_csv)
			headers, cols_pref, cols_no_pref = leer_csv(archivo_csv)
			self.grupos = construir_grupos(archivo_csv, cols_pref, cols_no_pref)
			self.resultados = generar_equipos(self.grupos)
			
			output = sys.stdout.getvalue()
			self.texto.insert(tk.END, output)
			
		except Exception as e:
			self.texto.insert(tk.END, f"ERROR: {str(e)}")
		finally:
			sys.stdout = old_stdout
	
	def exportar_excel(self):
		if not self.resultados or not self.grupos:
			self.texto.insert(tk.END, "\nNo hay resultados para exportar. Primero importa un archivo.\n")
			return
		
		archivo_salida = filedialog.asksaveasfilename(
			title="Guardar resultados",
			initialdir=os.path.expanduser("~/Desktop"),
			defaultextension=".xlsx",
			filetypes=[("Excel files", "*.xlsx"), ("Todos", "*.*")]
		)
		
		if not archivo_salida:
			return
		
		todas_las_filas = []
		
		for idx, grupo in enumerate(self.grupos):
			if 'nombre' in grupo and grupo['nombre'] == 'INDIFERENTE':
				nombre_grupo = 'GRUPO INDIFERENTE'
			else:
				nombre_grupo = f"GRUPO {idx + 1}"
			
			todos_miembros = grupo['miembros']
			
			todas_las_filas.append({
				'Grupo': nombre_grupo,
				'Miembros': ', '.join(todos_miembros),
				'Ranking': '',
				'Equipo': '',
				'Puntuacion': ''
			})
			
			resultado = next((r for r in self.resultados if r['grupo'] == nombre_grupo), None)
			
			if resultado and resultado['top_equipos']:
				for i, (equipo, puntuacion) in enumerate(resultado['top_equipos'], 1):
					todas_las_filas.append({
						'Grupo': '',
						'Miembros': '',
						'Ranking': i,
						'Equipo': ', '.join(equipo),
						'Puntuacion': puntuacion
					})
			
			todas_las_filas.append({
				'Grupo': '',
				'Miembros': '',
				'Ranking': '',
				'Equipo': '',
				'Puntuacion': ''
			})
		
		df = pd.DataFrame(todas_las_filas)
		df.to_excel(archivo_salida, index=False, sheet_name='Resultados')
		
		self.texto.insert(tk.END, f"\nResultados exportados a: {archivo_salida}\n")

if __name__ == '__main__':
	root = tk.Tk()
	app = Aplicacion(root)
	root.mainloop()
