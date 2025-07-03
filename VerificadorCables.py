import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
from datetime import datetime
from collections import defaultdict
import json

class VerificadorCables:
    def __init__(self):
        self.root = None
        self.ot_entry = None
        self.serie_entry = None
        self.resultado_text = None
        self.ruta_ilrl_label = None
        self.ruta_geo_label = None
        
        # Rutas base configuradas (ahora se cargarán de config.json)
        self.ruta_base_ilrl = r"C:\Users\Paulo\Desktop\ILRL JWS1-1" # Valor por defecto
        self.ruta_base_geo = r"C:\Users\Paulo\Desktop\Geometria JWS1-1" # Valor por defecto
        
        self.config_file = "config.json"
        self.password = "admin123" # Contraseña para acceder a la configuración
        
        # Variables para almacenar la última información analizada (para detalles)
        self.last_ilrl_analysis_data = None
        self.last_ilrl_file_path = None
        self.last_geo_analysis_data = None
        self.last_geo_file_path = None
        
        self.cargar_rutas() # Cargar las rutas al iniciar la aplicación

    def cargar_rutas(self):
        """Carga las rutas de los archivos desde un archivo de configuración JSON."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.ruta_base_ilrl = config.get('ruta_ilrl', self.ruta_base_ilrl)
                    self.ruta_base_geo = config.get('ruta_geo', self.ruta_base_geo)
            except Exception as e:
                messagebox.showerror("Error de Configuración", f"No se pudo cargar la configuración: {e}. Usando rutas por defecto.")
                self.guardar_rutas() # Guardar rutas por defecto si falla la carga
        else:
            self.guardar_rutas() # Guardar las rutas por defecto si el archivo no existe

    def guardar_rutas(self):
        """Guarda las rutas actuales en un archivo de configuración JSON."""
        config = {
            'ruta_ilrl': self.ruta_base_ilrl,
            'ruta_geo': self.ruta_base_geo
        }
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            messagebox.showinfo("Configuración Guardada", "Las rutas se han guardado correctamente.")
        except Exception as e:
            messagebox.showerror("Error al Guardar", f"No se pudieron guardar las rutas: {e}")

    def extraer_clave_ilrl(self, archivo):
        """Método para extraer clave de archivo ILRL"""
        base = os.path.splitext(archivo)[0]
        patron = r'JMO-(\d+)-(?:LC|SC|SCLC)-(\d{4})'
        m = re.match(patron, base)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
        return None

    def leer_resultado_ilrl(self, ruta):
        """
        Método para leer resultados ILRL.
        Retorna: resultado_final, ultima_fecha, df_original, col_resultado, col_fecha
        """
        try:
            df = pd.read_excel(ruta, header=None)
            inicio = 12

            # Intentar determinar la columna de resultados dinámicamente
            col7_vals = df.iloc[inicio:, 7].dropna().astype(str).str.upper()
            col8_vals = df.iloc[inicio:, 8].dropna().astype(str).str.upper()

            count_col7_pass_fail = col7_vals.isin(['PASS', 'FAIL']).sum()
            count_col8_pass_fail = col8_vals.isin(['PASS', 'FAIL']).sum()

            if count_col8_pass_fail >= count_col7_pass_fail and count_col8_pass_fail > 0:
                col_resultado = 8
                col_fecha = 10
            elif count_col7_pass_fail > 0:
                col_resultado = 7
                col_fecha = 9
            else: # No se encontraron 'PASS'/'FAIL' en las columnas esperadas
                return None, None, None, None, None

            resultados = df.iloc[inicio:, col_resultado].dropna().astype(str).str.upper().tolist()
            fechas_raw = df.iloc[inicio:, col_fecha].dropna().tolist()

            if not resultados:
                return None, None, None, None, None

            resultado_final = 'APROBADO' if all(r == 'PASS' for r in resultados) else 'RECHAZADO'

            fechas_datetime = []
            for f in fechas_raw:
                try:
                    if isinstance(f, datetime):
                        fechas_datetime.append(f)
                    else:
                        fechas_datetime.append(datetime.strptime(str(f).split('.')[0], "%d/%m/%Y %H:%M")) # Ajuste para formato
                except ValueError:
                    try:
                        fechas_datetime.append(datetime.strptime(str(f).split('.')[0], "%Y-%m-%d %H:%M:%S")) # Otro formato común
                    except:
                        pass # Ignorar fechas no parseables

            ultima_fecha = max(fechas_datetime).strftime("%d/%m/%Y %H:%M") if fechas_datetime else 'N/A'
            return resultado_final, ultima_fecha, df, col_resultado, col_fecha
        except Exception as e:
            print(f"Error leyendo {os.path.basename(ruta)}: {e}")
            return None, None, None, None, None

    def normalizar_serie_geo(self, serie_completo):
        """Método para normalizar serie de geometría"""
        texto = str(serie_completo).strip().upper()
        texto = re.sub(r'^\s*[JM]O\s*[\-\s]*', '', texto)
        match = re.search(r'(\d{13})', texto)
        if not match:
            return None, None
        serie = match.group(1)
        punta = texto[match.end():].strip()
        punta = re.sub(r'^[\-\s]+', '', punta)
        if not punta:
            return serie, None
        if 'R' in punta:
            punta = 'R' + punta.replace('R', '').replace('-', '')
        return serie, punta if punta in {'1','2','3','4','R1','R2','R3','R4'} else None

    def leer_resultado_geo(self, ruta):
        """
        Método para leer resultados de geometría.
        Retorna: resultados_por_serie, ultima_fecha, df_procesado
        """
        try:
            df = pd.read_excel(ruta, header=None, skiprows=12)
            datos = []
            
            for idx, row in df.iterrows():
                serie, punta = self.normalizar_serie_geo(row[0])
                if not serie or not punta:
                    continue
                    
                resultado = str(row[6]).strip().upper() if len(row) > 6 and not pd.isna(row[6]) else None
                fecha = row[3] if len(row) > 3 and not pd.isna(row[3]) else None
                hora = row[4] if len(row) > 4 and not pd.isna(row[4]) else None
                
                timestamp = None
                try:
                    if fecha and hora:
                        if isinstance(fecha, datetime):
                            if isinstance(hora, datetime): # Hora ya es datetime
                                timestamp = datetime.combine(fecha.date(), hora.time())
                            elif isinstance(hora, (float, int)): # Hora como número de serie de Excel
                                timestamp = fecha + pd.to_timedelta(hora, unit='D')
                            else: # Hora como string
                                hora_str = str(hora).split('.')[0] # Eliminar milisegundos si hay
                                timestamp = datetime.strptime(f"{fecha.strftime('%Y-%m-%d')} {hora_str}", "%Y-%m-%d %H:%M:%S")
                        elif isinstance(fecha, (float, int)): # Fecha como número de serie de Excel
                            base_date = datetime(1899, 12, 30) # Excel epoch date for Windows
                            if fecha < 60: # Handling Excel leap year bug for 1900
                                fecha -= 1
                            timestamp = base_date + pd.to_timedelta(fecha, unit='D')
                            if isinstance(hora, (float, int)):
                                timestamp += pd.to_timedelta(hora, unit='D')
                            else:
                                hora_str = str(hora).split('.')[0]
                                timestamp = datetime.strptime(f"{timestamp.strftime('%Y-%m-%d')} {hora_str}", "%Y-%m-%d %H:%M:%S")

                        elif isinstance(fecha, str) and isinstance(hora, str):
                            timestamp = datetime.strptime(f"{fecha.split('.')[0]} {hora.split('.')[0]}", "%Y-%m-%d %H:%M:%S")
                except Exception as e:
                    pass
                    
                datos.append({
                    'Serie': serie,
                    'Punta': punta,
                    'Resultado': resultado,
                    'Timestamp': timestamp
                })
            
            df_procesado = pd.DataFrame(datos)
            if df_procesado.empty:
                return None, None, None
            
            # Determinar estado final por serie
            resultados_por_serie = {}
            for serie, grupo in df_procesado.groupby('Serie'):
                ultima_medicion_por_punta = {}
                
                # Obtener la última medición para cada punta física (1, 2, 3, 4)
                for _, medicion in grupo.iterrows():
                    punta = medicion['Punta']
                    punta_fisica = punta.replace('R', '') # Ignorar 'R' para identificar la punta física
                    
                    # Si no existe o la nueva es más reciente
                    if punta_fisica not in ultima_medicion_por_punta or \
                       (medicion['Timestamp'] and ultima_medicion_por_punta[punta_fisica]['Timestamp'] and \
                        medicion['Timestamp'] > ultima_medicion_por_punta[punta_fisica]['Timestamp']):
                        ultima_medicion_por_punta[punta_fisica] = {
                            'Punta': punta, # Guardar la original (ej. R1)
                            'Resultado': medicion['Resultado'] == 'PASS',
                            'Timestamp': medicion['Timestamp']
                        }
                    elif not ultima_medicion_por_punta[punta_fisica]['Timestamp'] and medicion['Timestamp']:
                         ultima_medicion_por_punta[punta_fisica] = {
                            'Punta': punta, # Guardar la original (ej. R1)
                            'Resultado': medicion['Resultado'] == 'PASS',
                            'Timestamp': medicion['Timestamp']
                        }
                
                estado_final = "APROBADO"
                for p in ['1', '2', '3', '4']:
                    if p in ultima_medicion_por_punta:
                        if not ultima_medicion_por_punta[p]['Resultado']:
                            estado_final = "RECHAZADO"
                    else: # Si falta alguna punta
                        estado_final = "RECHAZADO"
                        
                resultados_por_serie[serie] = estado_final
            
            ultima_fecha_total = df_procesado['Timestamp'].max()
            return resultados_por_serie, ultima_fecha_total, df_procesado
        except Exception as e:
            print(f"Error leyendo {os.path.basename(ruta)}: {e}")
            return None, None, None

    def buscar_archivos_ilrl(self, ot_numero):
        """Busca archivos ILRL para la OT especificada"""
        ruta_ot = os.path.join(self.ruta_base_ilrl, ot_numero)
        if not os.path.exists(ruta_ot):
            return []
        
        archivos = []
        for f in os.listdir(ruta_ot):
            if f.endswith('.xlsx'):
                archivos.append(os.path.join(ruta_ot, f))
        return archivos

    def buscar_archivos_geo(self, ot_numero):
        """Busca archivos de Geometría para la OT especificada"""
        archivos = []
        for f in os.listdir(self.ruta_base_geo):
            if f.endswith('.xlsx') and ot_numero in f:
                archivos.append(os.path.join(self.ruta_base_geo, f))
        return archivos

    def verificar_cable_automatico(self, event=None):
        """Método que se llama automáticamente al escribir en el campo de serie."""
        serie_cable = self.serie_entry.get().strip()
        if len(serie_cable) == 13:
            self.verificar_cable()
        elif len(serie_cable) < 13:
            # Limpiar resultados si el número de serie es incompleto
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "Esperando un número de serie de 13 dígitos para iniciar la verificación...", "normal")
            # Quitar bindings de tags anteriores
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)

    def verificar_cable(self):
        ot_numero = self.ot_entry.get().strip().upper()
        serie_cable = self.serie_entry.get().strip()
        
        # Limpiar datos de análisis previos
        self.last_ilrl_analysis_data = None
        self.last_ilrl_file_path = None
        self.last_geo_analysis_data = None
        self.last_geo_file_path = None
        
        # Actualizar información de rutas en la interfaz
        self.ruta_ilrl_label.config(text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}")
        self.ruta_geo_label.config(text=f"📂 Ruta Geometría: {self.ruta_base_geo}")
        
        if not ot_numero or not serie_cable:
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "Por favor, ingrese OT y Número de Serie para verificar.", "normal")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            return
        
        if not re.match(r'^\d{13}$', serie_cable):
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "El número de serie debe tener 13 dígitos para realizar la verificación.", "rojo")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            return
        
        archivos_ilrl = self.buscar_archivos_ilrl(ot_numero)
        if not archivos_ilrl:
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, f"NO SE ENCONTRARON ARCHIVOS ILRL PARA LA OT {ot_numero}\n", "rojo")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            return
        
        archivos_geo = self.buscar_archivos_geo(ot_numero)
        if not archivos_geo:
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, f"NO SE ENCONTRARON ARCHIVOS DE GEOMETRÍA PARA LA OT {ot_numero}\n", "rojo")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            return
        
        # Procesar ILRL (usamos los últimos 4 dígitos del número de serie)
        serie_buscar_ilrl = serie_cable[-4:]
        resultado_ilrl = None
        fecha_ilrl = None
        df_ilrl_original = None
        col_ilrl_resultado = None
        col_ilrl_fecha = None
        
        for archivo in archivos_ilrl:
            clave = self.extraer_clave_ilrl(os.path.basename(archivo))
            if clave and clave.split('-')[1] == serie_buscar_ilrl:
                res, fecha, df_orig, col_res, col_fecha = self.leer_resultado_ilrl(archivo)
                if res:
                    resultado_ilrl = res
                    fecha_ilrl = fecha
                    df_ilrl_original = df_orig
                    col_ilrl_resultado = col_res
                    col_ilrl_fecha = col_fecha
                    self.last_ilrl_file_path = archivo # Almacenar la ruta
                    break
        
        # Almacenar datos para ILRL
        self.last_ilrl_analysis_data = {
            'df': df_ilrl_original,
            'col_resultado': col_ilrl_resultado,
            'col_fecha': col_ilrl_fecha,
            'resultado_general': resultado_ilrl,
            'fecha_general': fecha_ilrl
        }
        
        # Procesar Geometría (buscamos el número de serie completo)
        resultado_geo = None
        fecha_geo = None
        df_geo_procesado = None
        
        for archivo in archivos_geo:
            res_dict, fecha, df_proc = self.leer_resultado_geo(archivo)
            if res_dict and serie_cable in res_dict:
                resultado_geo = res_dict[serie_cable]
                fecha_geo = fecha
                df_geo_procesado = df_proc
                self.last_geo_file_path = archivo # Almacenar la ruta
                break
        
        # Almacenar datos para Geometría
        self.last_geo_analysis_data = {
            'df_procesado': df_geo_procesado,
            'resultado_general': resultado_geo,
            'fecha_general': fecha_geo,
            'serie_cable': serie_cable
        }

        # Mostrar resultados con formato
        self.resultado_text.config(state=tk.NORMAL)
        self.resultado_text.delete(1.0, tk.END)
        
        # Quitar bindings de tags anteriores para evitar múltiples llamadas
        self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
        self.resultado_text.tag_unbind("geo_click", "<Button-1>")

        # Encabezado
        self.resultado_text.insert(tk.END, f"🔍 Resultados para cable {serie_cable} en OT {ot_numero}:\n\n", "header")
        
        # Resultado ILRL
        self.resultado_text.insert(tk.END, "📊 ILRL: ", "bold")
        if resultado_ilrl:
            color_tag = "verde" if resultado_ilrl == "APROBADO" else "rojo"
            # Aplica el tag de color y el tag de click
            self.resultado_text.insert(tk.END, f"{resultado_ilrl}", (color_tag, "ilrl_click"))
            if fecha_ilrl:
                self.resultado_text.insert(tk.END, f" (📅 {fecha_ilrl})", "normal")
            # Vincula el tag a la función de detalles
            self.resultado_text.tag_bind("ilrl_click", "<Button-1>", lambda e: self.mostrar_detalles_ilrl())
            self.resultado_text.tag_config("ilrl_click", underline=1) # SOLO SUBRAYADO, SIN CAMBIAR EL COLOR
        else:
            self.resultado_text.insert(tk.END, f"NO ENCONTRADO (buscando terminación {serie_buscar_ilrl})", "rojo")
        self.resultado_text.insert(tk.END, "\n")
        
        # Resultado Geometría
        self.resultado_text.insert(tk.END, "📐 Geometría: ", "bold")
        if resultado_geo:
            color_tag = "verde" if resultado_geo == "APROBADO" else "rojo"
            # Aplica el tag de color y el tag de click
            self.resultado_text.insert(tk.END, f"{resultado_geo}", (color_tag, "geo_click"))
            if fecha_geo:
                fecha_str = fecha_geo.strftime('%d/%m/%Y %H:%M') if hasattr(fecha_geo, 'strftime') else str(fecha_geo)
                self.resultado_text.insert(tk.END, f" (📅 {fecha_str})", "normal")
            # Vincula el tag a la función de detalles
            self.resultado_text.tag_bind("geo_click", "<Button-1>", lambda e: self.mostrar_detalles_geo())
            self.resultado_text.tag_config("geo_click", underline=1) # SOLO SUBRAYADO, SIN CAMBIAR EL COLOR
        else:
            self.resultado_text.insert(tk.END, "NO ENCONTRADA", "rojo")
        self.resultado_text.insert(tk.END, "\n\n")
        
        # Estado final
        if resultado_ilrl and resultado_geo:
            estado_final = "APROBADO" if resultado_ilrl == "APROBADO" and resultado_geo == "APROBADO" else "RECHAZADO"
            self.resultado_text.insert(tk.END, "🏁 ESTADO FINAL: ", "bold")
            color = "verde" if estado_final == "APROBADO" else "rojo"
            self.resultado_text.insert(tk.END, f"{estado_final}\n", color)
            
            # Emoji adicional según resultado
            if estado_final == "APROBADO":
                self.resultado_text.insert(tk.END, "✅ ¡El cable cumple con todos los requisitos!\n", "verde")
            else:
                self.resultado_text.insert(tk.END, "❌ El cable no cumple con los requisitos\n", "rojo")
        else:
            # Si uno de los resultados no fue encontrado, el estado final es rechazado implícitamente
            self.resultado_text.insert(tk.END, "🏁 ESTADO FINAL: ", "bold")
            self.resultado_text.insert(tk.END, "RECHAZADO\n", "rojo")
            self.resultado_text.insert(tk.END, "❌ No se pudo verificar completamente el cable.\n", "rojo")

        self.resultado_text.config(state=tk.DISABLED)

    def mostrar_detalles_ilrl(self):
        if not self.last_ilrl_analysis_data or not self.last_ilrl_file_path:
            messagebox.showinfo("Detalles ILRL", "No hay datos de ILRL para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación ILRL")
        detalles_window.geometry("700x500")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        frame = ttk.Frame(detalles_window, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        # Ruta del archivo
        ttk.Label(frame, text="📁 Archivo Analizado:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        ttk.Label(frame, text=self.last_ilrl_file_path, wraplength=650, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

        # Resultado general y fecha
        ttk.Label(frame, text="📈 Resultado General ILRL:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        
        resultado_general = self.last_ilrl_analysis_data['resultado_general']
        fecha_general = self.last_ilrl_analysis_data['fecha_general']
        color = "green" if resultado_general == "APROBADO" else "red"
        
        info_label = ttk.Label(frame, text=f"{resultado_general} (Fecha de medición más reciente: {fecha_general})", 
                               font=("Arial", 10, "bold"), foreground=color, background="#F0F4F8")
        info_label.pack(anchor="w", pady=(0, 10))

        ttk.Label(frame, text="📊 Mediciones Detalladas por Línea:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))

        # Tabla de mediciones detalladas
        tree = ttk.Treeview(frame, columns=("Línea", "Resultado", "Fecha"), show="headings", height=10)
        tree.heading("Línea", text="Línea", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha", anchor=tk.W)

        tree.column("Línea", width=70, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=180, stretch=tk.NO)

        df = self.last_ilrl_analysis_data['df']
        col_resultado = self.last_ilrl_analysis_data['col_resultado']
        col_fecha = self.last_ilrl_analysis_data['col_fecha']
        inicio = 12 # Fila de inicio de los datos

        if df is not None and col_resultado is not None and col_fecha is not None:
            for i in range(inicio, len(df)):
                try:
                    resultado = str(df.iloc[i, col_resultado]).strip().upper()
                    fecha_raw = df.iloc[i, col_fecha]
                    fecha_str = 'N/A'
                    if isinstance(fecha_raw, datetime):
                        fecha_str = fecha_raw.strftime("%d/%m/%Y %H:%M")
                    else:
                        try:
                            fecha_str = datetime.strptime(str(fecha_raw).split('.')[0], "%d/%m/%Y %H:%M").strftime("%d/%m/%Y %H:%M")
                        except:
                            try:
                                fecha_str = datetime.strptime(str(fecha_raw).split('.')[0], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
                            except:
                                pass # No se pudo parsear
                    
                    if resultado in ['PASS', 'FAIL']:
                        tree.insert("", tk.END, values=(i - inicio + 1, resultado, fecha_str), 
                                    tags=('pass_style' if resultado == 'PASS' else 'fail_style'))
                except IndexError:
                    continue # Saltar si la fila no tiene suficientes columnas
        
        # Configurar estilos para las filas de Treeview
        tree.tag_configure('pass_style', foreground='green')
        tree.tag_configure('fail_style', foreground='red')

        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

        detalles_window.mainloop()

    def mostrar_detalles_geo(self):
        if not self.last_geo_analysis_data or not self.last_geo_file_path:
            messagebox.showinfo("Detalles Geometría", "No hay datos de Geometría para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación Geometría")
        detalles_window.geometry("700x500")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        frame = ttk.Frame(detalles_window, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        # Ruta del archivo
        ttk.Label(frame, text="📁 Archivo Analizado:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        ttk.Label(frame, text=self.last_geo_file_path, wraplength=650, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

        # Resultado general y fecha
        ttk.Label(frame, text=f"📈 Resultado General para Serie {self.last_geo_analysis_data['serie_cable']}:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        
        resultado_general = self.last_geo_analysis_data['resultado_general']
        fecha_general = self.last_geo_analysis_data['fecha_general']
        color = "green" if resultado_general == "APROBADO" else "red"
        
        fecha_str_disp = fecha_general.strftime("%d/%m/%Y %H:%M") if hasattr(fecha_general, 'strftime') else str(fecha_general)
        info_label = ttk.Label(frame, text=f"{resultado_general} (Fecha de medición más reciente: {fecha_str_disp})", 
                               font=("Arial", 10, "bold"), foreground=color, background="#F0F4F8")
        info_label.pack(anchor="w", pady=(0, 10))

        ttk.Label(frame, text="📐 Mediciones Detalladas por Punta:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))

        # Tabla de mediciones detalladas
        tree = ttk.Treeview(frame, columns=("Serie", "Punta", "Resultado", "Fecha"), show="headings", height=10)
        tree.heading("Serie", text="Serie", anchor=tk.W)
        tree.heading("Punta", text="Punta", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha y Hora", anchor=tk.W)

        tree.column("Serie", width=120, stretch=tk.NO)
        tree.column("Punta", width=70, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=180, stretch=tk.NO)

        df_procesado = self.last_geo_analysis_data['df_procesado']
        serie_cable_actual = self.last_geo_analysis_data['serie_cable']

        if df_procesado is not None and not df_procesado.empty:
            df_filtrado_por_serie = df_procesado[df_procesado['Serie'] == serie_cable_actual].copy()
            # Ordenar para mostrar la más reciente por punta al final o destacar
            df_filtrado_por_serie = df_filtrado_por_serie.sort_values(by=['Punta', 'Timestamp'], ascending=[True, True])

            for index, row in df_filtrado_por_serie.iterrows():
                resultado = str(row['Resultado']).strip().upper() if row['Resultado'] else 'N/A'
                fecha_str = row['Timestamp'].strftime("%d/%m/%Y %H:%M:%S") if row['Timestamp'] else 'N/A'
                
                tree.insert("", tk.END, values=(row['Serie'], row['Punta'], resultado, fecha_str), 
                            tags=('pass_style' if resultado == 'PASS' else 'fail_style'))
        
        tree.tag_configure('pass_style', foreground='green')
        tree.tag_configure('fail_style', foreground='red')

        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

        detalles_window.mainloop()

    def solicitar_contrasena(self):
        """Solicita la contraseña para acceder a la configuración de rutas."""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", "Ingrese la contraseña para cambiar las rutas:", show='*')
        if password_ingresada == self.password:
            self.mostrar_ventana_configuracion_rutas()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")

    def mostrar_ventana_configuracion_rutas(self):
        """Muestra la ventana para configurar las rutas de ILRL y Geometría."""
        config_window = tk.Toplevel(self.root)
        config_window.title("Configurar Rutas de Archivos")
        config_window.geometry("600x250")
        config_window.transient(self.root) # Hacerla modal respecto a la ventana principal
        config_window.grab_set() # Bloquear interacción con la ventana principal
        
        frame = ttk.Frame(config_window, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Ruta Base ILRL:", font=("Arial", 10, "bold"), foreground="#2C3E50").grid(row=0, column=0, sticky=tk.W, pady=5)
        ilrl_entry = ttk.Entry(frame, width=60, font=("Arial", 10), style="TEntry")
        ilrl_entry.insert(0, self.ruta_base_ilrl)
        ilrl_entry.grid(row=0, column=1, pady=5, padx=10, sticky="ew")

        ttk.Label(frame, text="Ruta Base Geometría:", font=("Arial", 10, "bold"), foreground="#2C3E50").grid(row=1, column=0, sticky=tk.W, pady=5)
        geo_entry = ttk.Entry(frame, width=60, font=("Arial", 10), style="TEntry")
        geo_entry.insert(0, self.ruta_base_geo)
        geo_entry.grid(row=1, column=1, pady=5, padx=10, sticky="ew")

        def guardar_nuevas_rutas():
            nueva_ilrl = ilrl_entry.get().strip()
            nueva_geo = geo_entry.get().strip()

            if not os.path.isdir(nueva_ilrl):
                messagebox.showwarning("Ruta Inválida", "La ruta de ILRL no es un directorio válido.")
                return
            if not os.path.isdir(nueva_geo):
                messagebox.showwarning("Ruta Inválida", "La ruta de Geometría no es un directorio válido.")
                return

            self.ruta_base_ilrl = nueva_ilrl
            self.ruta_base_geo = nueva_geo
            self.guardar_rutas()
            
            # Actualizar las etiquetas en la ventana principal
            self.ruta_ilrl_label.config(text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}")
            self.ruta_geo_label.config(text=f"📂 Ruta Geometría: {self.ruta_base_geo}")
            
            config_window.destroy()

        save_button = ttk.Button(frame, text="Guardar Rutas", command=guardar_nuevas_rutas, style="Primary.TButton")
        save_button.grid(row=2, column=0, columnspan=2, pady=20)
        
        config_window.columnconfigure(1, weight=1) # Hacer que la columna de entrada se expanda
        config_window.mainloop()


    def iniciar(self):
        self.root = tk.Tk()
        self.root.title("Verificador de Estado de Cables - Versión 1.2")
        self.root.geometry("800x650")
        self.root.configure(bg="#F0F4F8") # Fondo principal más suave
        
        # Crear barra de menú
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configuración", menu=config_menu)
        config_menu.add_command(label="Cambiar Rutas", command=self.solicitar_contrasena)
        
        # Configurar estilos
        self.style = ttk.Style()
        
        # Estilo general
        self.style.configure(".", background="#F0F4F8", font=("Arial", 10))
        self.style.configure("TFrame", background="#F0F4F8")
        self.style.configure("TLabel", background="#F0F4F8", foreground="#2C3E50") # Color de texto más oscuro
        self.style.configure("TButton", font=("Arial", 10, "bold"), padding=8, relief="flat", borderwidth=0)
        self.style.map("TButton", 
                       background=[('active', "#A0615E"), ('!disabled', "#CA364E")], # Azul para botones
                       foreground=[('active', 'black'), ('!disabled', 'black')])

        # Estilo para el botón principal (si se volviera a usar)
        self.style.configure("Primary.TButton", 
                           background="#28A745", # Verde para acciones principales
                           foreground="white",
                           font=("Arial", 11, "bold"),
                           borderwidth=0,
                           focusthickness=2,
                           focuscolor="#28A745")
        self.style.map("Primary.TButton", 
                       background=[('active', '#218838'), ('!disabled', '#28A745')])
        
        # Estilo para los campos de entrada
        self.style.configure("TEntry", 
                           fieldbackground="white", 
                           foreground="#2C3E50", 
                           borderwidth=1, 
                           relief="solid")
        
        # Estilo para los frames de contenido (tarjetas)
        self.style.configure("Card.TFrame", 
                           background="#FFFFFF", # Fondo blanco para las "tarjetas"
                           relief="flat", 
                           borderwidth=1, 
                           bordercolor="#E0E0E0") # Borde sutil
        
        # Estilo para el área de resultados
        self.style.configure("Result.TFrame", 
                           background="#FFFFFF", # Fondo blanco
                           relief="flat", # Borde plano
                           borderwidth=1, 
                           bordercolor="#E0E0E0") # Borde sutil
        
        # Estilo para las rutas
        self.style.configure("Path.TLabel", 
                           font=("Arial", 9), 
                           foreground="#6C757D", # Gris más oscuro
                           background="#F8F9FA", # Fondo para etiquetas de ruta
                           padding=(5,2)) # Pequeño padding para visualización

        # Estilo para el frame de entrada
        self.style.configure("Input.TFrame",
                            background="#FFFFFF",
                            relief="flat",
                            borderwidth=1,
                            bordercolor="#E0E0E0")

        # Estilo para Treeview (tablas en ventanas de detalles)
        self.style.configure("Treeview.Heading", font=("Arial", 9, "bold"), background="#E0E0E0", foreground="#2C3E50")
        self.style.configure("Treeview", font=("Arial", 9), rowheight=25)
        self.style.map("Treeview", background=[('selected', '#B0D7FF')]) # Selección de fila
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=(20, 15), style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_frame = ttk.Frame(main_frame, style="TFrame")
        title_frame.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        ttk.Label(title_frame, 
                text="🔍 Verificador de Cables de Fibra Óptica", 
                font=("Arial", 18, "bold"), 
                foreground="#2C3E50", # Azul oscuro
                background="#F0F4F8").pack()
        
        ttk.Label(title_frame, 
                text="Sistema de verificación de resultados ILRL y Geometría", 
                font=("Arial", 11), 
                foreground="#3498DB", # Azul vibrante
                background="#F0F4F8").pack()
        
        # Frame de entrada de datos
        input_frame = ttk.Frame(main_frame, padding=(20, 15), style="Input.TFrame")
        input_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        
        # Entrada OT
        ttk.Label(input_frame, text="📋 Orden de Trabajo (ej. JMO-250500001):", 
                 font=("Arial", 10, "bold"), foreground="#2C3E50").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ot_entry = ttk.Entry(input_frame, width=30, font=("Arial", 10), style="TEntry")
        self.ot_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=10)
        
        # Entrada Número de Serie
        ttk.Label(input_frame, text="🔢 Número de Serie (13 dígitos):", 
                 font=("Arial", 10, "bold"), foreground="#2C3E50").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.serie_entry = ttk.Entry(input_frame, width=30, font=("Arial", 10), style="TEntry")
        self.serie_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=10)
        # Bindea la función verificar_cable_automatico al evento de soltar cualquier tecla
        self.serie_entry.bind("<KeyRelease>", self.verificar_cable_automatico)
        
        input_frame.columnconfigure(1, weight=1) # Hacer que la columna de entrada se expanda
        
        # Sección de rutas de análisis
        path_frame = ttk.Frame(main_frame, padding=(15, 10), style="Card.TFrame")
        path_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Label(path_frame, 
                 text="🔎 RUTAS DE ANÁLISIS", 
                 font=("Arial", 10, "bold"), 
                 foreground="#3498DB", # Azul vibrante
                 background="#FFFFFF").pack(anchor="w", pady=(0, 5))
        
        self.ruta_ilrl_label = ttk.Label(path_frame, 
                                       text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}", # Actualizar al cargar
                                       style="Path.TLabel")
        self.ruta_ilrl_label.pack(anchor="w", padx=5, pady=2, fill="x")
        
        self.ruta_geo_label = ttk.Label(path_frame, 
                                      text=f"📂 Ruta Geometría: {self.ruta_base_geo}", # Actualizar al cargar
                                      style="Path.TLabel")
        self.ruta_geo_label.pack(anchor="w", padx=5, pady=2, fill="x")
        
        # Área de resultados (se ajusta la fila a la 3, ya que la 3 original era del botón)
        result_frame = ttk.Frame(main_frame, style="Result.TFrame", padding=15)
        result_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky="nsew")
        
        ttk.Label(result_frame, 
                text="📋 RESULTADOS DE LA VERIFICACIÓN", 
                font=("Arial", 11, "bold"), 
                foreground="#2C3E50", # Azul oscuro
                background="#FFFFFF").pack(anchor="w", pady=(0, 5))
        
        self.resultado_text = tk.Text(result_frame, 
                                    height=12, 
                                    width=80, 
                                    wrap=tk.WORD, 
                                    padx=10, 
                                    pady=10, 
                                    font=("Arial", 10),
                                    bg="white",
                                    bd=0,
                                    highlightthickness=0)
        self.resultado_text.pack(fill=tk.BOTH, expand=True)
        
        # Configurar tags para formato de texto
        self.resultado_text.tag_config("normal", foreground="#2C3E50") # Color de texto normal
        self.resultado_text.tag_config("bold", font=("Arial", 10, "bold"), foreground="#2C3E50")
        self.resultado_text.tag_config("header", 
                                     font=("Arial", 13, "bold"), 
                                     foreground="#2C3E50",
                                     justify="center")
        self.resultado_text.tag_config("verde", 
                                     foreground="#28A745", # Verde para PASS
                                     font=("Arial", 10, "bold"))
        self.resultado_text.tag_config("rojo", 
                                     foreground="#DC3545", # Rojo para FAIL
                                     font=("Arial", 10, "bold"))
        # Tags para los enlaces clickeables (solo subrayado)
        self.resultado_text.tag_config("ilrl_click", underline=1, font=("Arial", 10, "bold"))
        self.resultado_text.tag_config("geo_click", underline=1, font=("Arial", 10, "bold"))
        
        # Mensaje inicial en el área de resultados
        self.resultado_text.insert(tk.END, "Bienvenido al Verificador de Cables.\n\n"
                                     "Ingrese la Orden de Trabajo y el Número de Serie para iniciar.\n"
                                     "La verificación se realizará automáticamente al completar los 13 dígitos del número de serie.\n\n"
                                     "Haga click en los resultados 'APROBADO' o 'RECHAZADO' de ILRL o Geometría para ver los detalles.", "normal")
        self.resultado_text.config(state=tk.DISABLED) # Deshabilitar edición inicial
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(result_frame, 
                                command=self.resultado_text.yview,
                                style="TScrollbar")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.resultado_text.config(yscrollcommand=scrollbar.set)
        
        # Instrucciones (se ajusta la fila a la 4)
        instrucciones_frame = ttk.Frame(main_frame, padding=(15, 10), style="Card.TFrame")
        instrucciones_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="ew")
        
        instrucciones = (
            "📝 INSTRUCCIONES:\n"
            "1. Ingrese el número completo de la OT (ej. JMO-250500001)\n"
            "2. Ingrese el número de serie completo del cable (13 dígitos)\n"
            "3. Revise las rutas de análisis que se mostrarán arriba\n"
            "4. La verificación se realizará automáticamente al completar el número de serie (13 dígitos).\n"
            "5. Revise los resultados en la sección inferior, y haga click en el estatus para ver detalles."
        )
        ttk.Label(instrucciones_frame, 
                text=instrucciones, 
                wraplength=700, 
                justify=tk.LEFT,
                font=("Arial", 9),
                foreground="#6C757D", # Gris oscuro
                background="#FFFFFF").pack(anchor="w")

        # Frame para botones (Nuevo)
        button_exit_frame = ttk.Frame(main_frame, style="TFrame")
        button_exit_frame.grid(row=5, column=0, columnspan=2, pady=(15, 5))
        
        # Botón para salir (Nuevo)
        exit_button = ttk.Button(button_exit_frame, 
                                 text="🚫 Salir del Programa", 
                                 command=self.root.destroy, 
                                 style="TButton") # Usar estilo de botón general
        exit_button.pack(pady=5, ipadx=10, ipady=5)
        
        # Footer (se ajusta la fila a la 6)
        footer_frame = ttk.Frame(main_frame, style="TFrame")
        footer_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Label(footer_frame, 
                text="Sistema de Verificación de Cables v1.2 | Desarrollado por Paulo", 
                font=("Arial", 8), 
                foreground="#6C757D", # Gris oscuro
                background="#F0F4F8").pack()
        
        # Configurar el grid para que sea responsivo
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1) # El resultado_text expandirá verticalmente
        
        self.root.mainloop()

if __name__ == "__main__":
    app = VerificadorCables()
    app.iniciar()