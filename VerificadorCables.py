import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
from datetime import datetime
from collections import defaultdict
import json
import sqlite3

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
    
        # Variables para almacenar la última información analizada
        self.last_ilrl_analysis_data = None
        self.last_geo_analysis_data = None
        self.last_ilrl_file_path = None
        self.last_geo_file_path = None
    
        # Base de datos - ahora con ruta absoluta en el directorio del programa
        self.db_name = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cable_verifications.db")
    
        # Asegurarse de que el directorio existe
        os.makedirs(os.path.dirname(self.db_name), exist_ok=True)
    
        self._init_database() # Inicializar la base de datos al inicio
        self.cargar_rutas() # Cargar las rutas al iniciar la aplicación

        # Nuevo caché para almacenar los detalles de los elementos de Treeview
        self.item_data_cache = {}

    def _init_database(self):
        """Inicializa la base de datos SQLite y crea la tabla si no existe."""
        conn = None
        try:
            # Usar check_same_thread=False si hay problemas de hilos (común en Tkinter)
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
        
            # Verificar si la tabla ya existe
            cursor.execute("""
                SELECT count(name) FROM sqlite_master 
                WHERE type='table' AND name='cable_verifications'
            """)
        
            if cursor.fetchone()[0] == 0:
                # Crear tabla solo si no existe
                cursor.execute("""
                    CREATE TABLE cable_verifications (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        entry_date TEXT NOT NULL,
                        serial_number TEXT NOT NULL,
                        ot_number TEXT NOT NULL,
                        overall_status TEXT NOT NULL,
                        ilrl_status TEXT,
                        ilrl_date TEXT,
                        geo_status TEXT,
                        geo_date TEXT,
                        ilrl_details_json TEXT,
                        geo_details_json TEXT
                    )
                """)
                conn.commit()
            
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", 
                            f"No se pudo inicializar la base de datos: {e}")
            # Intentar crear el archivo si no existe y falló la conexión
            if not os.path.exists(self.db_name):
                try:
                    open(self.db_name, 'w').close()
                    # No reintentar init_database aquí para evitar bucles si el error es persistente
                    messagebox.showinfo("Base de Datos", "Archivo de base de datos creado. Intente reiniciar la aplicación.")
                except Exception as e:
                    messagebox.showerror("Error Crítico", 
                                    f"No se pudo crear el archivo de base de datos: {e}")
        finally:
            if conn:
                conn.close()

    def _log_verification_result(self, serial_number, ot_number, overall_status, 
                           ilrl_status, ilrl_date, ilrl_details, 
                           geo_status, geo_date, geo_details):
        """Registra el resultado de la verificación de un cable en la base de datos."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name, check_same_thread=False)
            cursor = conn.cursor()
            entry_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Convertir detalles a JSON strings para almacenamiento
            ilrl_details_json = json.dumps(ilrl_details, ensure_ascii=False) if ilrl_details else None
            geo_details_json = json.dumps(geo_details, ensure_ascii=False) if geo_details else None

            cursor.execute("""
                INSERT INTO cable_verifications (
                    entry_date, serial_number, ot_number, overall_status,
                    ilrl_status, ilrl_date, ilrl_details_json,
                    geo_status, geo_date, geo_details_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                entry_date, serial_number, ot_number, overall_status,
                ilrl_status, ilrl_date, ilrl_details_json,
                geo_status, geo_date, geo_details_json
            ))
        
            # Asegurarse de hacer commit explícito
            conn.commit()
        
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", 
                            f"No se pudo registrar el resultado: {e}\n"
                            f"Base de datos: {os.path.abspath(self.db_name)}")
        finally:
            if conn:
                conn.close()

    def verificar_ruta_db(self):
        """Muestra la ruta real de la base de datos para diagnóstico."""
        ruta_absoluta = os.path.abspath(self.db_name)
        messagebox.showinfo(
            "Ubicación de la Base de Datos",
            f"La base de datos se está guardando en:\n\n{ruta_absoluta}\n\n"
            f"Tamaño del archivo: {os.path.getsize(self.db_name) if os.path.exists(self.db_name) else 0} bytes"
        )

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
        """Método mejorado para extraer clave de archivo ILRL"""
        archivo = os.path.normpath(archivo)
        base = os.path.splitext(os.path.basename(archivo))[0]
    
        # Eliminar sufijo -F si está presente (para retrabajos)
        if base.endswith('-F'):
            base = base[:-2]
    
        # Patrón para todos los casos posibles:
        patron = r'JMO-(\d+)-(?:LC|SC|SCLC|LCSC)-(\d{4})'
        m = re.match(patron, base)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
        return None

    def leer_resultado_ilrl(self, ruta):
        """
        Método mejorado para leer resultados ILRL que maneja todos los casos.
        Retorna: resultado_final, ultima_fecha, lista_detalles_ilrl (para JSON)
        
        Modificado para devolver los resultados encontrados, incluso si no son 4,
        y el estado general del archivo basado en esos resultados.
        """
        try:
            if not os.path.exists(ruta):
                return None, None, None
            if os.path.basename(ruta).startswith('~$'):
                return None, None, None
            
            df = pd.read_excel(ruta, header=None)
            inicio = 12

            es_combinado = any(x in os.path.basename(ruta).upper() for x in ['SCLC', 'LCSC'])
        
            col_resultado = -1
            col_fecha = -1

            if es_combinado:
                pass_counts = []
                for col in [7, 8, 9, 10]:
                    col_vals = df.iloc[inicio:, col].dropna().astype(str).str.upper()
                    pass_count = col_vals.isin(['PASS']).sum()
                    pass_counts.append(pass_count)

                if max(pass_counts) > 0:
                    col_resultado = pass_counts.index(max(pass_counts)) + 7
                    col_fecha = col_resultado + 2
                else:
                    return None, None, None

            else: # Procesamiento normal para archivos no combinados (LC o SC)
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
                else:
                    return None, None, None

            if col_resultado == -1:
                return None, None, None

            resultados_raw = df.iloc[inicio:, col_resultado].dropna().astype(str).str.upper().tolist()
            valid_results = [r for r in resultados_raw if r in ['PASS', 'FAIL']]

            if not valid_results: # Si no hay resultados válidos en el archivo
                return None, None, None

            # Determinar el resultado final para ESTE ARCHIVO
            resultado_final = 'APROBADO' if all(r == 'PASS' for r in valid_results) else 'RECHAZADO'

            fechas_raw = df.iloc[inicio:, col_fecha].dropna().tolist()
            fechas_datetime = []
            for f in fechas_raw:
                try:
                    if isinstance(f, datetime):
                        fechas_datetime.append(f)
                    else:
                        fechas_datetime.append(datetime.strptime(str(f).split('.')[0], "%d/%m/%Y %H:%M"))
                except ValueError:
                    try:
                        fechas_datetime.append(datetime.strptime(str(f).split('.')[0], "%Y-%m-%d %H:%M:%S"))
                    except:
                        pass

            ultima_fecha = max(fechas_datetime).strftime("%d/%m/%Y %H:%M") if fechas_datetime else 'N/A'

            lista_detalles_ilrl = []
            for i, res_val in enumerate(valid_results):
                fecha_str_linea = 'N/A'
                if i < len(fechas_raw):
                    f_raw_linea = fechas_raw[i]
                    if isinstance(f_raw_linea, datetime):
                        fecha_str_linea = f_raw_linea.strftime("%d/%m/%Y %H:%M")
                    else:
                        try:
                            fecha_str_linea = datetime.strptime(str(f_raw_linea).split('.')[0], "%d/%m/%Y %H:%M").strftime("%d/%m/%Y %H:%M")
                        except:
                            try:
                                fecha_str_linea = datetime.strptime(str(f_raw_linea).split('.')[0], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
                            except:
                                pass

                lista_detalles_ilrl.append({
                    'linea': i + 1,
                    'resultado': res_val,
                    'fecha': fecha_str_linea,
                    'origen_archivo': os.path.basename(ruta),
                    'tipo_archivo': 'COMBINADO' if es_combinado else ('LC' if '-LC-' in os.path.basename(ruta).upper() else 'SC')
                })
            
            return resultado_final, ultima_fecha, lista_detalles_ilrl
        except Exception as e:
            print(f"Error leyendo {os.path.basename(ruta)}: {e}")
            return None, None, None

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
        Retorna: resultados_por_serie, ultima_fecha, detalles_geo_por_serie (para JSON)
        """
        try:
            # Verificar si el archivo existe y es accesible
            if not os.path.exists(ruta):
                print(f"Archivo no encontrado: {ruta}")
                return None, None, None
            
            # Verificar si el archivo está bloqueado o es temporal
            if os.path.basename(ruta).startswith('~$'):
                print(f"Ignorando archivo temporal de Excel: {ruta}")
                return None, None, None
            
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
                    if pd.isna(fecha) or pd.isna(hora):
                        continue
                    
                    if isinstance(fecha, datetime):
                        if isinstance(hora, datetime):
                            timestamp = datetime.combine(fecha.date(), hora.time())
                        elif isinstance(hora, (float, int)):
                            timestamp = fecha + pd.to_timedelta(hora, unit='D')
                        else:
                            hora_str = str(hora).split('.')[0]
                            timestamp = datetime.strptime(f"{fecha.strftime('%Y-%m-%d')} {hora_str}", "%Y-%m-%d %H:%M:%S")
                    elif isinstance(fecha, (float, int)):
                        base_date = datetime(1899, 12, 30)
                        if fecha < 60:
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
                    print(f"Error procesando fecha/hora en archivo {os.path.basename(ruta)}: {e}")
                    continue
                
                datos.append({
                    'Serie': serie,
                    'Punta': punta,
                    'Resultado': resultado,
                    'Timestamp': timestamp
                })
        
            df_procesado = pd.DataFrame(datos)
            if df_procesado.empty:
                return None, None, None
        
            resultados_por_serie = {}
            detalles_geo_por_serie = defaultdict(list)
        
            # Filtrar registros con timestamp válido
            df_procesado = df_procesado[df_procesado['Timestamp'].notna()]
        
            if df_procesado.empty:
                return None, None, None
            
            ultima_fecha_total = df_procesado['Timestamp'].max()

            for serie, grupo in df_procesado.groupby('Serie'):
                ultima_medicion_por_punta = {}
            
                for _, medicion in grupo.iterrows():
                    punta = medicion['Punta']
                    punta_fisica = punta.replace('R', '')
                
                    if punta_fisica not in ultima_medicion_por_punta or \
                    (medicion['Timestamp'] and ultima_medicion_por_punta[punta_fisica]['Timestamp'] and \
                        medicion['Timestamp'] > ultima_medicion_por_punta[punta_fisica]['Timestamp']):
                        ultima_medicion_por_punta[punta_fisica] = {
                            'Punta': punta,
                            'Resultado': medicion['Resultado'] == 'PASS',
                            'Timestamp': medicion['Timestamp']
                        }
                    elif not ultima_medicion_por_punta[punta_fisica]['Timestamp'] and medicion['Timestamp']:
                        ultima_medicion_por_punta[punta_fisica] = {
                            'Punta': punta,
                            'Resultado': medicion['Resultado'] == 'PASS',
                            'Timestamp': medicion['Timestamp']
                        }
            
                estado_final = "APROBADO"
                for p in ['1', '2', '3', '4']:
                    if p in ultima_medicion_por_punta:
                        if not ultima_medicion_por_punta[p]['Resultado']:
                            estado_final = "RECHAZADO"
                    else:
                        estado_final = "RECHAZADO"
                    
                resultados_por_serie[serie] = estado_final

                # Almacenar detalles para la serie actual (todas las mediciones para esa serie)
                for _, medicion in grupo.iterrows():
                    detalles_geo_por_serie[serie].append({
                        'serie': medicion['Serie'],
                        'punta': medicion['Punta'],
                        'resultado': medicion['Resultado'],
                        'timestamp': medicion['Timestamp'].strftime("%d/%m/%Y %H:%M:%S") if medicion['Timestamp'] else 'N/A'
                    })
        
            return resultados_por_serie, ultima_fecha_total, dict(detalles_geo_por_serie)
        except Exception as e:
            print(f"Error leyendo {os.path.basename(ruta)}: {e}")
            return None, None, None

    def buscar_archivos_ilrl(self, ot_numero):
        """Busca archivos ILRL para la OT especificada, incluyendo todos los casos"""
        ruta_ot = os.path.join(self.ruta_base_ilrl, ot_numero)
        archivos = []
    
        # Buscar en la carpeta principal de la OT
        if os.path.exists(ruta_ot):
            for f in os.listdir(ruta_ot):
                if f.endswith('.xlsx') and not f.startswith('~$'):
                    # Verificar si el nombre coincide con los patrones esperados
                    base_name = os.path.splitext(f)[0]
                    if any(x in base_name.upper() for x in ['-SC-', '-LC-', '-SCLC-', '-LCSC-']):
                        archivos.append(os.path.join(ruta_ot, f))
    
        # Buscar en la subcarpeta F si existe (para retrabajos)
        ruta_ot_f = os.path.join(ruta_ot, "F")
        if os.path.exists(ruta_ot_f):
            for f in os.listdir(ruta_ot_f):
                if f.endswith('.xlsx') and not f.startswith('~$'):
                    base_name = os.path.splitext(f)[0]
                    if any(x in base_name.upper() for x in ['-SC-', '-LC-', '-SCLC-', '-LCSC-']):
                        archivos.append(os.path.join(ruta_ot_f, f))
    
        return archivos

    def buscar_archivos_geo(self, ot_numero):
        """Busca archivos de Geometría para la OT especificada"""
        archivos = []
        for f in os.listdir(self.ruta_base_geo):
            # Excluir archivos temporales de Excel y asegurarse de que es .xlsx
            if f.endswith('.xlsx') and ot_numero in f and not f.startswith('~$'):
                archivos.append(os.path.join(self.ruta_base_geo, f))
        return archivos

    def verificar_cable_automatico(self, event=None):
        """Método que se llama automáticamente al escribir en el campo de serie."""
        serie_cable = self.serie_entry.get().strip()
        if len(serie_cable) == 13:
            self.verificar_cable()
        elif len(serie_cable) < 13:
            # Limpiar resultados if the serial number is incomplete
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
        
        # --- Poka-Yoke: Validación de coincidencia OT y Número de Serie ---
        match_ot = re.search(r'(\d+)', ot_numero)
        ot_numerico_parte = match_ot.group(1) if match_ot else None

        serie_ot_parte = serie_cable[:9] if len(serie_cable) >= 9 else None

        if ot_numerico_parte is None or serie_ot_parte is None or ot_numerico_parte != serie_ot_parte:
            messagebox.showwarning(
                "Error de Coincidencia",
                f"La Orden de Trabajo '{ot_numero}' no coincide con la parte inicial "
                f"del Número de Serie '{serie_cable}'.\n\n"
                "Verifique que los datos sean correctos. No se realizará la verificación ni el registro."
            )
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, "⚠️ ERROR: La OT y el Número de Serie no coinciden.\n"
                                         "Por favor, verifique los datos.", "rojo")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            return

        # --- Procesamiento ILRL ---
        serie_buscar_ilrl = serie_cable[-4:]
        all_ilrl_details_collected = [] # Lista para recolectar detalles de todas las puntas encontradas
        ilrl_file_paths_for_display = [] # Para almacenar los nombres de archivo para mostrar

        archivos_ilrl = self.buscar_archivos_ilrl(ot_numero)

        if not archivos_ilrl:
            resultado_ilrl = "NO ENCONTRADO"
            fecha_ilrl = None
            ilrl_detalles_para_db = None
        else:
            for archivo in archivos_ilrl:
                base_name = os.path.basename(archivo)
                clave = self.extraer_clave_ilrl(base_name)
        
                if clave and clave.split('-')[1] == serie_buscar_ilrl:
                    res_file, fecha_file, detalles_ilrl_list_file = self.leer_resultado_ilrl(archivo)
                    if detalles_ilrl_list_file: # Si la función devolvió detalles válidos
                        all_ilrl_details_collected.extend(detalles_ilrl_list_file)
                        ilrl_file_paths_for_display.append(archivo) # Añadir a la lista de archivos procesados

            # --- Consolidar resultados ILRL de todas las puntas recolectadas ---
            resultado_ilrl = "NO ENCONTRADO"
            fecha_ilrl = None
            ilrl_detalles_para_db = {
                'lc_file': None, # Estos campos ahora serán más informativos, no solo booleanos
                'sc_file': None,
                'combinado_file': None,
                'overall_ilrl_status': None,
                'latest_ilrl_date': None,
                'combined_details': []
            }

            if not all_ilrl_details_collected:
                resultado_ilrl = "NO ENCONTRADO"
            else:
                # Filtrar detalles duplicados si una punta aparece en múltiples archivos (mantener la más reciente si hay fechas)
                # Para simplificar, asumiremos que 'linea' + 'tipo_archivo' es un identificador de punta única
                unique_ilrl_details = {} 
                latest_date_overall = datetime.min

                for detail in all_ilrl_details_collected:
                    # Crear un identificador único para cada "punta"
                    # Asumimos que 'linea' es la punta (1, 2, 3, 4) y 'tipo_archivo' (LC, SC, COMBINADO) ayuda a la unicidad
                    # Si un cable es LC-0001 (punta 1,2) y otro LC-0001 (punta 3,4) de la misma OT,
                    # necesitamos que se identifiquen como 4 puntas distintas.
                    # Para esto, usaremos una combinación de origen_archivo y línea.
                    
                    # Un identificador de punta más robusto podría ser (origen_archivo, linea)
                    # O si las puntas tienen nombres específicos (ej. 'Punta A', 'Punta B'), usar eso.
                    # Dado que 'linea' es un número, y puede repetirse entre archivos,
                    # usaremos una combinación de archivo + línea como identificador único.
                    
                    unique_id = (detail.get('origen_archivo'), detail.get('linea'))

                    # Si ya tenemos esta punta, solo la actualizamos si la nueva es más reciente
                    current_detail = unique_ilrl_details.get(unique_id)
                    if current_detail:
                        try:
                            current_date = datetime.strptime(current_detail.get('fecha'), "%d/%m/%Y %H:%M")
                            new_date = datetime.strptime(detail.get('fecha'), "%d/%m/%Y %H:%M")
                            if new_date > current_date:
                                unique_ilrl_details[unique_id] = detail
                        except (ValueError, TypeError):
                            unique_ilrl_details[unique_id] = detail # Si la fecha no es parseable, simplemente reemplazamos
                    else:
                        unique_ilrl_details[unique_id] = detail

                    # Actualizar la fecha más reciente general
                    try:
                        detail_date = datetime.strptime(detail.get('fecha'), "%d/%m/%Y %H:%M")
                        if detail_date > latest_date_overall:
                            latest_date_overall = detail_date
                    except (ValueError, TypeError):
                        pass # Ignorar fechas no válidas

                final_consolidated_details = list(unique_ilrl_details.values())
                
                # Verificar el estado final basado en las puntas consolidadas
                total_puntas_encontradas = len(final_consolidated_details)
                all_puntas_pass = all(d.get('resultado') == 'PASS' for d in final_consolidated_details)

                if total_puntas_encontradas == 4 and all_puntas_pass:
                    resultado_ilrl = "APROBADO"
                elif total_puntas_encontradas > 0: # Si se encontraron puntas pero no 4 o no todas PASS
                    resultado_ilrl = "RECHAZADO"
                else: # No se encontraron puntas válidas en absoluto
                    resultado_ilrl = "NO ENCONTRADO"

                fecha_ilrl = latest_date_overall.strftime("%d/%m/%Y %H:%M") if latest_date_overall != datetime.min else 'N/A'
                
                ilrl_detalles_para_db['overall_ilrl_status'] = resultado_ilrl
                ilrl_detalles_para_db['latest_ilrl_date'] = fecha_ilrl
                ilrl_detalles_para_db['combined_details'] = final_consolidated_details

                # Actualizar las rutas de archivo en last_ilrl_file_path para la interfaz
                self.last_ilrl_file_path = "\n".join(ilrl_file_paths_for_display) if ilrl_file_paths_for_display else "N/A"

        self.last_ilrl_analysis_data = ilrl_detalles_para_db

        # --- Procesamiento Geometría ---
        resultado_geo = "NO ENCONTRADO"
        fecha_geo = None
        geo_detalles_para_db = None
        
        archivos_geo = self.buscar_archivos_geo(ot_numero)
        if not archivos_geo:
            # No se muestra mensaje de error aquí, se maneja en el resultado final
            pass
        else:
            for archivo in archivos_geo:
                try:
                    res_dict, fecha, detalles_geo_dict = self.leer_resultado_geo(archivo)
                    if res_dict and serie_cable in res_dict:
                        resultado_geo = res_dict[serie_cable]
                        fecha_geo = fecha
                        self.last_geo_file_path = archivo
                        geo_detalles_para_db = {
                            'file_path': archivo,
                            'resultado_general': resultado_geo,
                            'fecha_general': fecha.strftime("%d/%m/%Y %H:%M:%S") if hasattr(fecha, 'strftime') else str(fecha),
                            'detalles_puntas': detalles_geo_dict.get(serie_cable, [])
                        }
                        break
                except Exception as e:
                    print(f"Error procesando archivo {archivo}: {e}")
                    continue
        
        self.last_geo_analysis_data = geo_detalles_para_db
        
        # --- Determinación del Estatus General y Registro en DB ---
        overall_status_db = "NO ENCONTRADO"
        if resultado_ilrl != "NO ENCONTRADO" and resultado_geo != "NO ENCONTRADO":
            overall_status_db = "APROBADO" if resultado_ilrl == "APROBADO" and resultado_geo == "APROBADO" else "RECHAZADO"
        elif resultado_ilrl != "NO ENCONTRADO" and resultado_geo == "NO ENCONTRADO":
            overall_status_db = "RECHAZADO" # Si ILRL está y GEO no, se rechaza
        elif resultado_ilrl == "NO ENCONTRADO" and resultado_geo != "NO ENCONTRADO":
            overall_status_db = "RECHAZADO" # Si GEO está y ILRL no, se rechaza
        
        # Log results to database
        self._log_verification_result(
            serial_number=serie_cable,
            ot_number=ot_numero,
            overall_status=overall_status_db,
            ilrl_status=resultado_ilrl,
            ilrl_date=fecha_ilrl,
            geo_status=resultado_geo,
            geo_date=fecha_geo.strftime("%d/%m/%Y %H:%M:%S") if hasattr(fecha_geo, 'strftime') else str(fecha_geo),
            ilrl_details=ilrl_detalles_para_db,
            geo_details=geo_detalles_para_db
        )

        # --- Mostrar resultados en la interfaz ---
        self.resultado_text.config(state=tk.NORMAL)
        self.resultado_text.delete(1.0, tk.END)
        
        # Quitar bindings de tags anteriores para evitar múltiples llamadas
        self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
        self.resultado_text.tag_unbind("geo_click", "<Button-1>")

        self.resultado_text.insert(tk.END, f"🔍 Resultados para cable {serie_cable} en OT {ot_numero}:\n\n", "header")
        
        self.resultado_text.insert(tk.END, "📊 ILRL: ", "bold")
        if resultado_ilrl != "NO ENCONTRADO":
            color_tag = "verde" if resultado_ilrl == "APROBADO" else "rojo"
            self.resultado_text.insert(tk.END, f"{resultado_ilrl}", (color_tag, "ilrl_click"))
            if fecha_ilrl:
                self.resultado_text.insert(tk.END, f" (📅 {fecha_ilrl})", "normal")
            self.resultado_text.tag_bind("ilrl_click", "<Button-1>", lambda e: self.mostrar_detalles_ilrl(self.last_ilrl_analysis_data))
            self.resultado_text.tag_config("ilrl_click", underline=1)
        else:
            self.resultado_text.insert(tk.END, f"NO ENCONTRADO (buscando terminación {serie_buscar_ilrl})", "rojo")
        self.resultado_text.insert(tk.END, "\n")
        
        self.resultado_text.insert(tk.END, "📐 Geometría: ", "bold")
        if resultado_geo != "NO ENCONTRADO":
            color_tag = "verde" if resultado_geo == "APROBADO" else "rojo"
            self.resultado_text.insert(tk.END, f"{resultado_geo}", (color_tag, "geo_click"))
            if fecha_geo:
                fecha_str = fecha_geo.strftime('%d/%m/%Y %H:%M') if hasattr(fecha_geo, 'strftime') else str(fecha_geo)
                self.resultado_text.insert(tk.END, f" (📅 {fecha_str})", "normal")
            self.resultado_text.tag_bind("geo_click", "<Button-1>", lambda e: self.mostrar_detalles_geo(self.last_geo_analysis_data))
            self.resultado_text.tag_config("geo_click", underline=1)
        else:
            self.resultado_text.insert(tk.END, "NO ENCONTRADA", "rojo")
        self.resultado_text.insert(tk.END, "\n\n")
        
        # Estado final para la interfaz (usa overall_status_db para consistencia)
        self.resultado_text.insert(tk.END, "🏁 ESTADO FINAL: ", "bold")
        color = "verde" if overall_status_db == "APROBADO" else "rojo" if overall_status_db == "RECHAZADO" else "orange"
        self.resultado_text.insert(tk.END, f"{overall_status_db}\n", color)
            
        if overall_status_db == "APROBADO":
            self.resultado_text.insert(tk.END, "✅ ¡El cable cumple con todos los requisitos!\n", "verde")
        elif overall_status_db == "RECHAZADO":
            self.resultado_text.insert(tk.END, "❌ El cable no cumple con los requisitos\n", "rojo")
        else: # NO ENCONTRADO
            self.resultado_text.insert(tk.END, "⚠️ No se pudo verificar completamente el cable.\n", "orange")

        self.resultado_text.config(state=tk.DISABLED)

    def mostrar_detalles_ilrl(self, data=None):
        """Muestra una ventana con los detalles completos del análisis ILRL"""
        details_to_show = data if data else self.last_ilrl_analysis_data

        if not details_to_show:
            messagebox.showinfo("Detalles ILRL", "No hay datos de ILRL para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación ILRL")
        detalles_window.geometry("800x600")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        # Configurar el estilo
        style = ttk.Style()
        style.configure("Detalles.TFrame", background="#F0F4F8")
        style.configure("Detalles.TLabel", background="#F0F4F8", foreground="#2C3E50")
        style.configure("Detalles.Treeview", background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Detalles.Treeview.Heading", background="#E9ECEF", foreground="#2C3E50", font=('Arial', 9, 'bold'))
        style.map("Detalles.Treeview", background=[('selected', '#007BFF')])

        # Frame principal con scrollbar
        main_frame = ttk.Frame(detalles_window, style="Detalles.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame, background="#F0F4F8", highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Detalles.TFrame")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Función para el mouse wheel
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * int((event.delta / 120)), "units")
    
        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        # Contenido del frame desplazable
        content_frame = ttk.Frame(scrollable_frame, style="Detalles.TFrame", padding=(20, 20))
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        ttk.Label(content_frame, 
                text="📊 Detalles Completos de Verificación ILRL", 
                font=("Arial", 12, "bold"), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 15))

        # Sección de archivos analizados (ahora muestra los archivos que contribuyeron)
        files_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
        files_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(files_frame, 
                text="📁 Archivos Analizados:", 
                font=("Arial", 10, "bold"), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))
        
        processed_files_display = details_to_show.get('ilrl_analizado_paths', [])
        if processed_files_display:
            for file_path in processed_files_display:
                origen = "(Subcarpeta F)" if "\\F\\" in file_path else "(Carpeta principal)"
                ttk.Label(files_frame, 
                        text=f"• {os.path.basename(file_path)} {origen}", 
                        wraplength=700, 
                        font=("Arial", 9), 
                        foreground="#6C757D", 
                        background="#F0F4F8").pack(anchor="w")
        else:
            ttk.Label(files_frame, 
                    text="• N/A (Ningún archivo ILRL encontrado para esta serie)", 
                    wraplength=700, 
                    font=("Arial", 9), 
                    foreground="#6C757D", 
                    background="#F0F4F8").pack(anchor="w")


    # Resultado general
        result_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
        result_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(result_frame, 
                text="📈 Resultado General ILRL:", 
                font=("Arial", 10, "bold"), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))
    
        resultado_general = details_to_show.get('overall_ilrl_status', 'N/A')
        fecha_general = details_to_show.get('latest_ilrl_date', 'N/A')
        color = "green" if resultado_general == "APROBADO" else "red" if resultado_general == "RECHAZADO" else "orange"
    
        ttk.Label(result_frame, 
                text=f"• Estado: {resultado_general}", 
                font=("Arial", 10), 
                foreground=color, 
                background="#F0F4F8").pack(anchor="w")
    
        ttk.Label(result_frame, 
                text=f"• Fecha de medición más reciente: {fecha_general}", 
                font=("Arial", 9), 
                foreground="#6C757D", 
                background="#F0F4F8").pack(anchor="w")

    # Detalles de las mediciones
        ttk.Label(content_frame, 
                text="📊 Mediciones Detalladas:", 
                font=("Arial", 10, "bold"), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))

    # Crear Treeview con scrollbar
        tree_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True)

    # Configurar Treeview
        tree = ttk.Treeview(
            tree_frame,
            columns=("Línea", "Resultado", "Fecha", "Tipo Archivo", "Origen"),
            show="headings",
            height=10,
            style="Detalles.Treeview"
        )

    # Configurar columnas
        tree.heading("Línea", text="Línea", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha", anchor=tk.W)
        tree.heading("Tipo Archivo", text="Tipo Archivo", anchor=tk.W)
        tree.heading("Origen", text="Origen", anchor=tk.W)

        tree.column("Línea", width=50, stretch=tk.NO, anchor=tk.W)
        tree.column("Resultado", width=80, stretch=tk.NO, anchor=tk.W)
        tree.column("Fecha", width=120, stretch=tk.NO, anchor=tk.W)
        tree.column("Tipo Archivo", width=100, stretch=tk.NO, anchor=tk.W)
        tree.column("Origen", width=100, stretch=tk.NO, anchor=tk.W)

    # Scrollbar para el Treeview
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree.pack(side="left", fill=tk.BOTH, expand=True)
        tree_scroll.pack(side="right", fill="y")

    # Configurar tags para colores
        tree.tag_configure('PASS', foreground='green')
        tree.tag_configure('FAIL', foreground='red')

    # Llenar el Treeview con los datos
        detalles_lineas = details_to_show.get('combined_details', [])
        for detalle in detalles_lineas:
            resultado = detalle.get('resultado', 'N/A')
            origen = "(Subcarpeta F)" if "\\F\\" in detalle.get('origen_archivo', '') else "(Carpeta principal)"
            tipo_archivo = detalle.get('tipo_archivo', 'N/A')
        
            tree.insert(
                "", 
                tk.END, 
                values=(
                    detalle.get('linea', 'N/A'),
                    resultado,
                    detalle.get('fecha', 'N/A'),
                    tipo_archivo,
                    origen
                ), 
                tags=(resultado,)
            )

    # Estadísticas resumen
        stats_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
        stats_frame.pack(fill=tk.X, pady=(15, 0))

    # Contar PASS/FAIL
        total = len(detalles_lineas)
        pass_count = sum(1 for d in detalles_lineas if d.get('resultado') == 'PASS')
        fail_count = total - pass_count

        ttk.Label(stats_frame, 
                text="📝 Resumen Estadístico:", 
                font=("Arial", 10, "bold"), 
                style="Detalles.TLabel").pack(anchor="w", pady=(0, 5))

        stats_text = (
            f"• Total de mediciones: {total}\n"
            f"• Aprobadas (PASS): {pass_count} ({pass_count/total*100:.1f}%)\n"
            f"• Rechazadas (FAIL): {fail_count} ({fail_count/total*100:.1f}%)"
        )
    
        ttk.Label(stats_frame, 
                text=stats_text, 
                justify=tk.LEFT,
                font=("Arial", 9), 
                foreground="#6C757D", 
                background="#F0F4F8").pack(anchor="w")

    # Botón de cierre
        btn_frame = ttk.Frame(content_frame, style="Detalles.TFrame")
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(btn_frame, 
                text="Cerrar", 
                command=detalles_window.destroy,
                style="TButton").pack(pady=10)

        detalles_window.mainloop()

    def mostrar_detalles_geo(self, data=None):
        details_to_show = data if data else self.last_geo_analysis_data
        
        if not details_to_show:
            messagebox.showinfo("Detalles Geometría", "No hay datos de Geometría para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación Geometría")
        detalles_window.geometry("700x500")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        frame = ttk.Frame(detalles_window, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="📁 Archivo Analizado:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        origen = "(Subcarpeta F)" if "\\F\\" in details_to_show.get('file_path', '') else "(Carpeta principal)"
        ttk.Label(frame, text=f"{details_to_show.get('file_path', 'N/A')} {origen}", wraplength=650, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

        ttk.Label(frame, text=f"📈 Resultado General para Geometría:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        resultado_general = details_to_show.get('resultado_general', 'N/A')
        fecha_general = details_to_show.get('fecha_general', 'N/A')
        color = "green" if resultado_general == "APROBADO" else "red"
        info_label = ttk.Label(frame, text=f"{resultado_general} (Fecha de medición más reciente: {fecha_general})", 
                               font=("Arial", 10, "bold"), foreground=color, background="#F0F4F8")
        info_label.pack(anchor="w", pady=(0, 10))

        ttk.Label(frame, text="📐 Mediciones Detalladas por Punta:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))

        tree = ttk.Treeview(frame, columns=("Serie", "Punta", "Resultado", "Fecha"), show="headings", height=10)
        tree.heading("Serie", text="Serie", anchor=tk.W)
        tree.heading("Punta", text="Punta", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha y Hora", anchor=tk.W)

        tree.column("Serie", width=120, stretch=tk.NO)
        tree.column("Punta", width=70, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=180, stretch=tk.NO)

        detalles_puntas = details_to_show.get('detalles_puntas', [])
        for detalle in detalles_puntas:
            resultado = detalle.get('resultado', 'N/A')
            tree.insert("", tk.END, values=(detalle.get('serie', 'N/A'), detalle.get('punta', 'N/A'), resultado, detalle.get('timestamp', 'N/A')), 
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

    def solicitar_contrasena_registros(self):
        """Solicita la contraseña para acceder a la vista de registros."""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", "Ingrese la contraseña para ver los registros:", show='*')
        if password_ingresada == self.password:
            self.mostrar_vista_registros()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")

    def mostrar_ventana_configuracion_rutas(self):
        """Muestra la ventana para configurar las rutas de ILRL y Geometría."""
        config_window = tk.Toplevel(self.root)
        config_window.title("Configurar Rutas de Archivos")
        config_window.geometry("600x250")
        config_window.transient(self.root)
        config_window.grab_set()

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
            self.ruta_ilrl_label.config(text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}")
            self.ruta_geo_label.config(text=f"📂 Ruta Geometría: {self.ruta_base_geo}")
            config_window.destroy()

        save_button = ttk.Button(frame, text="Guardar Rutas", command=guardar_nuevas_rutas, style="Primary.TButton")
        save_button.grid(row=2, column=0, columnspan=2, pady=20)

        config_window.columnconfigure(1, weight=1)
        config_window.mainloop()

    def _borrar_todos_los_registros(self):
        """Borra todos los registros de la tabla cable_verifications."""
        if not messagebox.askyesno("Confirmar Eliminación", 
                                   "¿Está seguro de que desea borrar TODOS los registros de la base de datos?\n\n"
                                   "Esta acción es irreversible."):
            return

        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM cable_verifications")
            conn.commit()
            messagebox.showinfo("Éxito", "Todos los registros han sido eliminados correctamente.")
            if hasattr(self, 'tree_registros'):
                self.cargar_registros()
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron borrar los registros: {e}")
        finally:
            if conn:
                conn.close()

    def solicitar_contrasena_borrar_datos(self):
        """Solicita la contraseña para borrar todos los datos de la base de datos."""
        password_ingresada = simpledialog.askstring("Contraseña Requerida", 
                                                     "Ingrese la contraseña para borrar TODOS los datos:", 
                                                     show='*')
        if password_ingresada == self.password:
            self._borrar_todos_los_registros()
        else:
            messagebox.showerror("Acceso Denegado", "Contraseña incorrecta.")

    def mostrar_vista_registros(self):
        """Muestra la ventana para que un ingeniero visualice los registros de cables."""
        registros_window = tk.Toplevel(self.root)
        registros_window.title("Vista de Registros de Cables")
        registros_window.geometry("1000x700")
        registros_window.transient(self.root)
        registros_window.grab_set()

        main_frame = ttk.Frame(registros_window, padding=(20, 20), style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Frame para filtros y botones
        filter_frame = ttk.Frame(main_frame, style="TFrame")
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(filter_frame, text="Filtrar por OT o Serie:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(side=tk.LEFT, padx=(0, 5))
        self.filtro_entry = ttk.Entry(filter_frame, width=30, font=("Arial", 10), style="TEntry")
        self.filtro_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.filtro_entry.bind("<KeyRelease>", self.aplicar_filtro_registros)

        btn_aplicar_filtro = ttk.Button(filter_frame, text="Aplicar Filtro", command=self.aplicar_filtro_registros, style="TButton")
        btn_aplicar_filtro.pack(side=tk.LEFT, padx=(0, 10))

        btn_limpiar_filtro = ttk.Button(filter_frame, text="Limpiar Filtro", command=self.limpiar_filtro_registros, style="TButton")
        btn_limpiar_filtro.pack(side=tk.LEFT, padx=(0, 20))

        # Nuevo botón para borrar todos los datos
        btn_borrar_todos = ttk.Button(filter_frame, text="🗑️ Borrar Todos los Registros", 
                                      command=self.solicitar_contrasena_borrar_datos, style="Danger.TButton")
        btn_borrar_todos.pack(side=tk.RIGHT)

        # Treeview para mostrar los registros
        columns = ("ID", "Fecha Entrada", "Número Serie", "Número OT", "Estado General", 
                   "ILRL Estatus", "ILRL Fecha", "Geo Estatus", "Geo Fecha")
        self.tree_registros = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        for col in columns:
            self.tree_registros.heading(col, text=col, anchor=tk.W)
            self.tree_registros.column(col, width=100, anchor=tk.W)

        self.tree_registros.column("ID", width=50, stretch=tk.NO)
        self.tree_registros.column("Fecha Entrada", width=140, stretch=tk.NO)
        self.tree_registros.column("Número Serie", width=120, stretch=tk.NO)
        self.tree_registros.column("Número OT", width=120, stretch=tk.NO)
        self.tree_registros.column("Estado General", width=100, stretch=tk.NO)
        self.tree_registros.column("ILRL Estatus", width=90, stretch=tk.NO)
        self.tree_registros.column("ILRL Fecha", width=120, stretch=tk.NO)
        self.tree_registros.column("Geo Estatus", width=90, stretch=tk.NO)
        self.tree_registros.column("Geo Fecha", width=120, stretch=tk.NO)

        self.tree_registros.pack(fill=tk.BOTH, expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree_registros.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_registros.configure(yscrollcommand=scrollbar.set)

        # Configurar tags para colores de estado
        self.tree_registros.tag_configure('APROBADO', foreground='green')
        self.tree_registros.tag_configure('RECHAZADO', foreground='red')
        self.tree_registros.tag_configure('NO ENCONTRADO', foreground='orange')

        # Asociar evento de clic a las filas para mostrar detalles
        self.tree_registros.bind("<Double-1>", self.mostrar_detalles_registro_bd)

        self.cargar_registros()
        registros_window.mainloop()

    def cargar_registros(self):
        """Carga los registros de la base de datos en el Treeview."""
        for item in self.tree_registros.get_children():
            self.tree_registros.delete(item)
        
        self.item_data_cache = {}

        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM cable_verifications ORDER BY entry_date DESC")
            registros = cursor.fetchall()

            for i, row in enumerate(registros):
                ilrl_details = json.loads(row[9]) if row[9] else None
                geo_details = json.loads(row[10]) if row[10] else None

                self.item_data_cache[row[0]] = {
                    "id": row[0],
                    "entry_date": row[1],
                    "serial_number": row[2],
                    "ot_number": row[3],
                    "overall_status": row[4],
                    "ilrl_status": row[5],
                    "ilrl_date": row[6],
                    "geo_status": row[7],
                    "geo_date": row[8],
                    "ilrl_details": ilrl_details,
                    "geo_details": geo_details
                }

                self.tree_registros.insert("", tk.END, iid=row[0], values=(
                    row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]
                ), tags=(row[4],))
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los registros: {e}")
        finally:
            if conn:
                conn.close()

    def aplicar_filtro_registros(self, event=None):
        """Aplica un filtro a los registros mostrados en el Treeview."""
        filtro = self.filtro_entry.get().strip().upper()
        
        for item in self.tree_registros.get_children():
            self.tree_registros.delete(item)
            
        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            if filtro:
                cursor.execute("""
                    SELECT * FROM cable_verifications 
                    WHERE UPPER(ot_number) LIKE ? OR serial_number LIKE ?
                    ORDER BY entry_date DESC
                """, (f"%{filtro}%", f"%{filtro}%"))
            else:
                cursor.execute("SELECT * FROM cable_verifications ORDER BY entry_date DESC")
                
            registros = cursor.fetchall()

            for i, row in enumerate(registros):
                ilrl_details = json.loads(row[9]) if row[9] else None
                geo_details = json.loads(row[10]) if row[10] else None

                self.item_data_cache[row[0]] = {
                    "id": row[0],
                    "entry_date": row[1],
                    "serial_number": row[2],
                    "ot_number": row[3],
                    "overall_status": row[4],
                    "ilrl_status": row[5],
                    "ilrl_date": row[6],
                    "geo_status": row[7],
                    "geo_date": row[8],
                    "ilrl_details": ilrl_details,
                    "geo_details": geo_details
                }
                
                self.tree_registros.insert("", tk.END, iid=row[0], values=(
                    row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]
                ), tags=(row[4],))
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los registros filtrados: {e}")
        finally:
            if conn:
                conn.close()

    def limpiar_filtro_registros(self):
        """Limpia el campo de filtro y recarga todos los registros."""
        self.filtro_entry.delete(0, tk.END)
        self.cargar_registros()

    def mostrar_detalles_registro_bd(self, event):
        """Muestra una ventana de detalles para el registro seleccionado en la base de datos."""
        selected_item_id = self.tree_registros.focus()
        if not selected_item_id:
            return

        record_id = int(selected_item_id)
        record_data = self.item_data_cache.get(record_id)

        if not record_data:
            messagebox.showerror("Error", "No se encontraron los detalles del registro.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title(f"Detalles del Registro #{record_data['id']}")
        detalles_window.geometry("800x600")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        frame = ttk.Frame(detalles_window, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="📋 Información General:", font=("Arial", 12, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 10))
        
        info_general_text = (
            f"   • ID de Registro: {record_data['id']}\n"
            f"   • Fecha de Entrada: {record_data['entry_date']}\n"
            f"   • Número de Serie: {record_data['serial_number']}\n"
            f"   • Número de OT: {record_data['ot_number']}\n"
        )
        ttk.Label(frame, text=info_general_text, justify=tk.LEFT, font=("Arial", 10), foreground="#6C757D", background="#F0F4F8").pack(anchor="w")

        # Estado general
        ttk.Label(frame, text="🏁 Estado General:", font=("Arial", 12, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(10, 5))
        overall_status_color = "green" if record_data['overall_status'] == "APROBADO" else "red" if record_data['overall_status'] == "RECHAZADO" else "orange"
        ttk.Label(frame, text=f"   • {record_data['overall_status']}", font=("Arial", 10, "bold"), foreground=overall_status_color, background="#F0F4F8").pack(anchor="w")

        # Detalles ILRL
        ttk.Label(frame, text="📊 Detalles ILRL:", font=("Arial", 12, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(10, 5))
        ilrl_status_color = "green" if record_data['ilrl_status'] == "APROBADO" else "red" if record_data['ilrl_status'] == "RECHAZADO" else "orange"
        ttk.Label(frame, text=f"   • Estado: {record_data['ilrl_status']}", font=("Arial", 10, "bold"), foreground=ilrl_status_color, background="#F0F4F8").pack(anchor="w")
        ttk.Label(frame, text=f"   • Fecha: {record_data['ilrl_date'] if record_data['ilrl_date'] else 'N/A'}", font=("Arial", 10), foreground="#6C757D", background="#F0F4F8").pack(anchor="w")
        
        ilrl_details_from_db = record_data['ilrl_details']
        if ilrl_details_from_db:
            # Ahora mostramos los archivos que contribuyeron a la verificación ILRL
            processed_files_display = details_to_show.get('ilrl_analizado_paths', [])
            if processed_files_display:
                ttk.Label(frame, text="   • Archivos procesados:", font=("Arial", 9, "bold"), foreground="#6C757D", background="#F0F4F8").pack(anchor="w")
                for file_path in processed_files_display:
                    origen = "(Subcarpeta F)" if "\\F\\" in file_path else "(Carpeta principal)"
                    ttk.Label(frame, text=f"     - {os.path.basename(file_path)} {origen}", font=("Arial", 9), foreground="#6C757D", background="#F0F4F8", wraplength=700).pack(anchor="w")
            
            btn_ver_detalles_ilrl = ttk.Button(frame, text="Ver Detalles ILRL (Ventana Completa)", 
                                            command=lambda: self.mostrar_detalles_ilrl(ilrl_details_from_db), 
                                            style="Secondary.TButton")
            btn_ver_detalles_ilrl.pack(anchor="w", pady=(5, 5))
        else:
            ttk.Label(frame, text="   • No hay detalles ILRL disponibles.", font=("Arial", 10), foreground="#999999", background="#F0F4F8").pack(anchor="w")

        # Detalles Geometría
        ttk.Label(frame, text="📐 Detalles Geometría:", font=("Arial", 12, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(10, 5))
        geo_status_color = "green" if record_data['geo_status'] == "APROBADO" else "red" if record_data['geo_status'] == "RECHAZADO" else "orange"
        ttk.Label(frame, text=f"   • Estado: {record_data['geo_status']}", font=("Arial", 10, "bold"), foreground=geo_status_color, background="#F0F4F8").pack(anchor="w")
        geo_date_str = record_data['geo_date'] if record_data['geo_date'] else 'N/A'
        ttk.Label(frame, text=f"   • Fecha: {geo_date_str}", font=("Arial", 10), foreground="#6C757D", background="#F0F4F8").pack(anchor="w")

        if record_data['geo_details'] and record_data['geo_details'].get('file_path'):
            origen = "(Subcarpeta F)" if "\\F\\" in record_data['geo_details']['file_path'] else "(Carpeta principal)"
            ttk.Label(frame, text=f"   • Archivo: {record_data['geo_details']['file_path']} {origen}", font=("Arial", 9), foreground="#6C757D", background="#F0F4F8", wraplength=700).pack(anchor="w")
            btn_ver_detalles_geo = ttk.Button(frame, text="Ver Detalles Geometría (Ventana Completa)", 
                                              command=lambda: self.mostrar_detalles_geo(record_data['geo_details']), 
                                              style="Secondary.TButton")
            btn_ver_detalles_geo.pack(anchor="w", pady=(5, 5))
        else:
            ttk.Label(frame, text="   • No hay detalles de Geometría disponibles.", font=("Arial", 10), foreground="#999999", background="#F0F4F8").pack(anchor="w")
        
        detalles_window.mainloop()

    def create_main_window(self):
        self.root = tk.Tk()
        self.root.title("Sistema de Verificación de Cables JWS1-1")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        # Configuración de estilos ttk
        style = ttk.Style()
        style.theme_use('clam')

        style.configure("TFrame", background="#F0F4F8")
        style.configure("TLabel", background="#F0F4F8", foreground="#333333")
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#333333")
        style.configure("TButton", background="#007BFF", foreground="#FFFFFF", font=("Arial", 10, "bold"), padding=6)
        style.map("TButton", background=[('active', '#0056b3')])

        # Estilo para botones primarios (ej. Guardar Rutas)
        style.configure("Primary.TButton", background="#28A745", foreground="#FFFFFF")
        style.map("Primary.TButton", background=[('active', '#218838')])

        # Estilo para botones secundarios (ej. Ver Detalles)
        style.configure("Secondary.TButton", background="#6C757D", foreground="#FFFFFF")
        style.map("Secondary.TButton", background=[('active', '#5A6268')])

        # Estilo para botón de peligro (ej. Borrar Datos)
        style.configure("Danger.TButton", background="#DC3545", foreground="#FFFFFF")
        style.map("Danger.TButton", background=[('active', '#C82333')])

        # Create a Canvas and a Scrollbar
        canvas = tk.Canvas(self.root, background="#F0F4F8")
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Create a frame inside the canvas to hold the content
        scrollable_content_frame = ttk.Frame(canvas, padding=(20, 20, 20, 10), style="TFrame")
        canvas.create_window((0, 0), window=scrollable_content_frame, anchor="nw")

        # Configure canvas scrolling
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scrollable_content_frame.bind("<Configure>", on_frame_configure)
        
        # Make the scrollbar work with mouse wheel
        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * int((event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        # Título
        ttk.Label(scrollable_content_frame, 
                  text="⚙️ Sistema de Verificación de Cables", 
                  font=("Arial", 16, "bold"), 
                  foreground="#0056b3",
                  background="#F0F4F8").grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Sección de Entrada
        input_frame = ttk.Frame(scrollable_content_frame, padding=10, relief="solid", borderwidth=1, style="TFrame")
        input_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Label(input_frame, text="Número de OT:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ot_entry = ttk.Entry(input_frame, width=40, font=("Arial", 10), style="TEntry")
        self.ot_entry.grid(row=0, column=1, pady=5, padx=10, sticky="ew")
        
        ttk.Label(input_frame, text="Número de Serie del Cable (13 dígitos):", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.serie_entry = ttk.Entry(input_frame, width=40, font=("Arial", 10), style="TEntry")
        self.serie_entry.grid(row=1, column=1, pady=5, padx=10, sticky="ew")
        self.serie_entry.bind("<KeyRelease>", self.verificar_cable_automatico)

        # Botones de Acción
        button_frame = ttk.Frame(scrollable_content_frame, padding=10, style="TFrame")
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        btn_verificar = ttk.Button(button_frame, text="✅ Verificar Cable", command=self.verificar_cable, style="Primary.TButton")
        btn_verificar.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)

        btn_config_rutas = ttk.Button(button_frame, text="⚙️ Configurar Rutas", command=self.solicitar_contrasena, style="TButton")
        btn_config_rutas.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        
        btn_ver_registros = ttk.Button(button_frame, text="📊 Ver Registros", command=self.solicitar_contrasena_registros, style="TButton")
        btn_ver_registros.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)

        # En la sección de botones, después de btn_ver_registros
        btn_diagnostico_db = ttk.Button(
            button_frame, 
            text="🛠️ Diagnóstico DB", 
            command=self.verificar_ruta_db, 
            style="TButton"
    )
        btn_diagnostico_db.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)

        # --- Nuevo diseño para rutas e instrucciones ---
        info_area_frame = ttk.Frame(scrollable_content_frame, style="TFrame")
        info_area_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")
        
        # Sub-frame para las Rutas de Análisis (a la izquierda de las instrucciones)
        rutas_frame = ttk.Frame(info_area_frame, padding=10, style="TFrame")
        rutas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        self.ruta_ilrl_label = ttk.Label(rutas_frame, text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}", font=("Arial", 9), foreground="#666666")
        self.ruta_ilrl_label.pack(anchor="w")

        self.ruta_geo_label = ttk.Label(rutas_frame, text=f"📂 Ruta Geometría: {self.ruta_base_geo}", font=("Arial", 9), foreground="#666666")
        self.ruta_geo_label.pack(anchor="w")

        # Sub-frame para las Instrucciones (a la derecha de las rutas)
        instrucciones_frame = ttk.Frame(info_area_frame, padding=10, style="TFrame")
        instrucciones_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        instrucciones = (
            "📝 INSTRUCCIONES:\n"
            "1. Ingrese el número completo de la OT (ej. JMO-250500001).\n"
            "2. Ingrese el número de serie completo del cable (13 dígitos).\n"
            "3. Revise las rutas de análisis que se mostrarán arriba.\n"
            "4. Haga clic en 'Verificar Cable' (o espere la verificación automática).\n"
            "5. Revise los resultados en la sección inferior, y haga click en el estatus para ver detalles."
        )
        ttk.Label(instrucciones_frame, 
                  text=instrucciones, 
                  wraplength=350,
                  justify=tk.LEFT,
                  font=("Arial", 9),
                  foreground="#6C757D",
                  background="#F0F4F8").pack(anchor="w")

        # Sección de Resultados
        resultado_frame = ttk.Frame(scrollable_content_frame, padding=10, relief="solid", borderwidth=1, style="TFrame")
        resultado_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Label(resultado_frame, text="Resultados de Verificación:", font=("Arial", 10, "bold"), foreground="#2C3E50").pack(anchor="w")
        self.resultado_text = tk.Text(resultado_frame, height=10, width=80, wrap="word", 
                                       font=("Arial", 10), state=tk.DISABLED, 
                                       background="#FFFFFF", foreground="#333333")
        self.resultado_text.pack(pady=5, fill=tk.BOTH, expand=True)

        # Configuración de estilos para el widget Text
        self.resultado_text.tag_configure("normal", font=("Arial", 10), foreground="#333333")
        self.resultado_text.tag_configure("header", font=("Arial", 12, "bold"), foreground="#0056b3")
        self.resultado_text.tag_configure("bold", font=("Arial", 10, "bold"), foreground="#333333")
        self.resultado_text.tag_configure("verde", foreground="#28A745")
        self.resultado_text.tag_configure("rojo", foreground="#DC3545")
        self.resultado_text.tag_configure("orange", foreground="#FFC107")
        
        button_exit_frame = ttk.Frame(scrollable_content_frame, style="TFrame")
        button_exit_frame.grid(row=5, column=0, columnspan=2, pady=(15, 5))
        
        exit_button = ttk.Button(button_exit_frame, 
                                 text="🚫 Salir del Programa", 
                                 command=self.root.destroy, 
                                 style="TButton")
        exit_button.pack(pady=5, ipadx=10, ipady=5)
        
        # Footer
        footer_frame = ttk.Frame(scrollable_content_frame, style="TFrame")
        footer_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Label(footer_frame, 
                  text="Sistema de Verificación de Cables v1.2", 
                  font=("Arial", 8), 
                  foreground="#6C757D",
                  background="#F0F4F8").pack()
        
        scrollable_content_frame.grid_columnconfigure(0, weight=1)
        scrollable_content_frame.grid_columnconfigure(1, weight=1)

        self.root.mainloop()

if __name__ == "__main__":
    app = VerificadorCables()
    app.create_main_window()
