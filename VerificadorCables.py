import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
from datetime import datetime
from collections import defaultdict
import json
import sqlite3 # Importar el módulo SQLite

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
        self.last_ilrl_analysis_data = None # Almacena un dict con 'file_path', 'resultado_general', 'fecha_general', 'detalles_lineas'
        self.last_geo_analysis_data = None # Almacena un dict con 'file_path', 'resultado_general', 'fecha_general', 'detalles_puntas', 'serie_cable'
        
        self.db_name = "cable_verifications.db"
        self._init_database() # Inicializar la base de datos al inicio

        self.cargar_rutas() # Cargar las rutas al iniciar la aplicación

    def _init_database(self):
        """Inicializa la base de datos SQLite y crea la tabla si no existe."""
        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS cable_verifications (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    entry_date TEXT NOT NULL,
                    serial_number TEXT UNIQUE NOT NULL, -- UNIQUE para el comportamiento de upsert
                    ot_number TEXT NOT NULL,
                    overall_status TEXT NOT NULL,
                    ilrl_status TEXT,
                    ilrl_date TEXT,
                    geo_status TEXT,
                    geo_date TEXT,
                    ilrl_details_json TEXT, -- Almacenar detalles como JSON string
                    geo_details_json TEXT   -- Almacenar detalles como JSON string
                )
            """)
            conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo inicializar la base de datos: {e}")
        finally:
            if conn:
                conn.close()

    def _log_verification_result(self, serial_number, ot_number, overall_status, 
                                 ilrl_status, ilrl_date, ilrl_details, 
                                 geo_status, geo_date, geo_details):
        """
        Registra o actualiza el resultado de la verificación de un cable en la base de datos.
        Realiza un 'upsert' (UPDATE si existe, INSERT si no).
        """
        conn = None
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            entry_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Convertir detalles a JSON strings para almacenamiento
            ilrl_details_json = json.dumps(ilrl_details) if ilrl_details else None
            geo_details_json = json.dumps(geo_details) if geo_details else None

            cursor.execute("""
                INSERT INTO cable_verifications (entry_date, serial_number, ot_number, overall_status,
                                                 ilrl_status, ilrl_date, ilrl_details_json,
                                                 geo_status, geo_date, geo_details_json)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(serial_number) DO UPDATE SET
                    entry_date = EXCLUDED.entry_date,
                    ot_number = EXCLUDED.ot_number,
                    overall_status = EXCLUDED.overall_status,
                    ilrl_status = EXCLUDED.ilrl_status,
                    ilrl_date = EXCLUDED.ilrl_date,
                    ilrl_details_json = EXCLUDED.ilrl_details_json,
                    geo_status = EXCLUDED.geo_status,
                    geo_date = EXCLUDED.geo_date,
                    geo_details_json = EXCLUDED.geo_details_json
            """, (entry_date, serial_number, ot_number, overall_status,
                  ilrl_status, ilrl_date, ilrl_details_json,
                  geo_status, geo_date, geo_details_json))
            conn.commit()
            print(f"Registro para {serial_number} actualizado/insertado correctamente.")
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo registrar/actualizar el resultado: {e}")
        finally:
            if conn:
                conn.close()

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
        Retorna: resultado_final, ultima_fecha, lista_detalles_ilrl (para JSON)
        """
        try:
            df = pd.read_excel(ruta, header=None)
            inicio = 12

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
                return None, None, None # No se encontraron 'PASS'/'FAIL'

            resultados = df.iloc[inicio:, col_resultado].dropna().astype(str).str.upper().tolist()
            fechas_raw = df.iloc[inicio:, col_fecha].dropna().tolist()

            if not resultados:
                return None, None, None

            resultado_final = 'APROBADO' if all(r == 'PASS' for r in resultados) else 'RECHAZADO'

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

            # Preparar detalles ILRL para JSON
            lista_detalles_ilrl = []
            for i in range(inicio, len(df)):
                try:
                    resultado_linea = str(df.iloc[i, col_resultado]).strip().upper()
                    fecha_raw_linea = df.iloc[i, col_fecha]
                    fecha_str_linea = 'N/A'
                    if isinstance(fecha_raw_linea, datetime):
                        fecha_str_linea = fecha_raw_linea.strftime("%d/%m/%Y %H:%M")
                    else:
                        try:
                            fecha_str_linea = datetime.strptime(str(fecha_raw_linea).split('.')[0], "%d/%m/%Y %H:%M").strftime("%d/%m/%Y %H:%M")
                        except:
                            try:
                                fecha_str_linea = datetime.strptime(str(fecha_raw_linea).split('.')[0], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
                            except:
                                pass
                    if resultado_linea in ['PASS', 'FAIL']:
                        lista_detalles_ilrl.append({
                            'linea': i - inicio + 1,
                            'resultado': resultado_linea,
                            'fecha': fecha_str_linea
                        })
                except IndexError:
                    continue
            
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
            
            resultados_por_serie = {}
            detalles_geo_por_serie = defaultdict(list)
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
            
            return resultados_por_serie, ultima_fecha_total, dict(detalles_geo_por_serie) # Convertir defaultdict a dict
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
        
        # --- Procesamiento ILRL ---
        serie_buscar_ilrl = serie_cable[-4:]
        resultado_ilrl = "NO ENCONTRADO"
        fecha_ilrl = None
        ilrl_detalles_para_db = None
        
        archivos_ilrl = self.buscar_archivos_ilrl(ot_numero)
        if not archivos_ilrl:
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, f"NO SE ENCONTRARON ARCHIVOS ILRL PARA LA OT {ot_numero}\n", "rojo")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            # Log this as "NO ENCONTRADO" ILRL status, overall will be "NO ENCONTRADO" or "RECHAZADO" later
            # We will log the overall status at the end after checking both ILRL and GEO
        else:
            for archivo in archivos_ilrl:
                clave = self.extraer_clave_ilrl(os.path.basename(archivo))
                if clave and clave.split('-')[1] == serie_buscar_ilrl:
                    res, fecha, detalles_ilrl_list = self.leer_resultado_ilrl(archivo)
                    if res:
                        resultado_ilrl = res
                        fecha_ilrl = fecha
                        self.last_ilrl_file_path = archivo # Almacenar la ruta para la interfaz
                        ilrl_detalles_para_db = { # Preparar para DB
                            'file_path': archivo,
                            'resultado_general': res,
                            'fecha_general': fecha,
                            'detalles_lineas': detalles_ilrl_list
                        }
                        break
        
        self.last_ilrl_analysis_data = ilrl_detalles_para_db # Almacenar para la interfaz

        # --- Procesamiento Geometría ---
        resultado_geo = "NO ENCONTRADO"
        fecha_geo = None
        geo_detalles_para_db = None
        
        archivos_geo = self.buscar_archivos_geo(ot_numero)
        if not archivos_geo:
            self.resultado_text.config(state=tk.NORMAL)
            self.resultado_text.delete(1.0, tk.END)
            self.resultado_text.insert(tk.END, f"NO SE ENCONTRARON ARCHIVOS DE GEOMETRÍA PARA LA OT {ot_numero}\n", "rojo")
            self.resultado_text.tag_unbind("ilrl_click", "<Button-1>")
            self.resultado_text.tag_unbind("geo_click", "<Button-1>")
            self.resultado_text.config(state=tk.DISABLED)
            # We will log the overall status at the end after checking both ILRL and GEO
        else:
            for archivo in archivos_geo:
                res_dict, fecha, detalles_geo_dict = self.leer_resultado_geo(archivo)
                if res_dict and serie_cable in res_dict:
                    resultado_geo = res_dict[serie_cable]
                    fecha_geo = fecha
                    self.last_geo_file_path = archivo # Almacenar la ruta para la interfaz
                    geo_detalles_para_db = { # Preparar para DB
                        'file_path': archivo,
                        'resultado_general': resultado_geo,
                        'fecha_general': fecha.strftime("%d/%m/%Y %H:%M:%S") if hasattr(fecha, 'strftime') else str(fecha),
                        'detalles_puntas': detalles_geo_dict.get(serie_cable, [])
                    }
                    break
        
        self.last_geo_analysis_data = geo_detalles_para_db # Almacenar para la interfaz

        # --- Determinación del Estatus General y Registro en DB ---
        overall_status_db = "NO ENCONTRADO"
        if resultado_ilrl != "NO ENCONTRADO" and resultado_geo != "NO ENCONTRADO":
            overall_status_db = "APROBADO" if resultado_ilrl == "APROBADO" and resultado_geo == "APROBADO" else "RECHAZADO"
        elif resultado_ilrl != "NO ENCONTRADO" and resultado_geo == "NO ENCONTRADO":
            overall_status_db = "RECHAZADO" # ILRL encontrado, pero GEO no -> estatus general rechazado
        elif resultado_ilrl == "NO ENCONTRADO" and resultado_geo != "NO ENCONTRADO":
            overall_status_db = "RECHAZADO" # GEO encontrado, pero ILRL no -> estatus general rechazado
        
        # Log results to database
        self._log_verification_result(
            serial_number=serie_cable,
            ot_number=ot_numero,
            overall_status=overall_status_db,
            ilrl_status=resultado_ilrl,
            ilrl_date=fecha_ilrl,
            ilrl_details=ilrl_detalles_para_db,
            geo_status=resultado_geo,
            geo_date=fecha_geo.strftime("%d/%m/%Y %H:%M:%S") if hasattr(fecha_geo, 'strftime') else str(fecha_geo),
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
        details_to_show = data if data else self.last_ilrl_analysis_data

        if not details_to_show:
            messagebox.showinfo("Detalles ILRL", "No hay datos de ILRL para mostrar detalles. Realice una verificación primero.")
            return

        detalles_window = tk.Toplevel(self.root)
        detalles_window.title("Detalles de Verificación ILRL")
        detalles_window.geometry("700x500")
        detalles_window.transient(self.root)
        detalles_window.grab_set()

        frame = ttk.Frame(detalles_window, padding=(20, 20), style="TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="📁 Archivo Analizado:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        ttk.Label(frame, text=details_to_show.get('file_path', 'N/A'), wraplength=650, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

        ttk.Label(frame, text="📈 Resultado General ILRL:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        
        resultado_general = details_to_show.get('resultado_general', 'N/A')
        fecha_general = details_to_show.get('fecha_general', 'N/A')
        color = "green" if resultado_general == "APROBADO" else "red"
        
        info_label = ttk.Label(frame, text=f"{resultado_general} (Fecha de medición más reciente: {fecha_general})", 
                               font=("Arial", 10, "bold"), foreground=color, background="#F0F4F8")
        info_label.pack(anchor="w", pady=(0, 10))

        ttk.Label(frame, text="📊 Mediciones Detalladas por Línea:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))

        tree = ttk.Treeview(frame, columns=("Línea", "Resultado", "Fecha"), show="headings", height=10)
        tree.heading("Línea", text="Línea", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha", anchor=tk.W)

        tree.column("Línea", width=70, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=180, stretch=tk.NO)

        detalles_lineas = details_to_show.get('detalles_lineas', [])
        for detalle in detalles_lineas:
            resultado = detalle.get('resultado', 'N/A')
            tree.insert("", tk.END, values=(detalle.get('linea', 'N/A'), resultado, detalle.get('fecha', 'N/A')), 
                        tags=('pass_style' if resultado == 'PASS' else 'fail_style'))
        
        tree.tag_configure('pass_style', foreground='green')
        tree.tag_configure('fail_style', foreground='red')

        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

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
        ttk.Label(frame, text=details_to_show.get('file_path', 'N/A'), wraplength=650, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

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

    def mostrar_vista_registros(self):
        """Muestra la ventana para que un ingeniero visualice los registros de cables."""
        registros_window = tk.Toplevel(self.root)
        registros_window.title("Vista de Registros de Cables")
        registros_window.geometry("1000x700")
        registros_window.transient(self.root)
        registros_window.grab_set()

        main_frame = ttk.Frame(registros_window, padding=(20, 20), style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Sección de filtro por OT
        filter_frame = ttk.Frame(main_frame, style="Card.TFrame", padding=(15, 10))
        filter_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(filter_frame, text="Filtrar por OT:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#FFFFFF").pack(side=tk.LEFT, padx=(0, 10))
        ot_filter_entry = ttk.Entry(filter_frame, width=25, font=("Arial", 10), style="TEntry")
        ot_filter_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        def cargar_registros(ot_filter=None):
            for item in tree.get_children():
                tree.delete(item)
            conn = None
            try:
                conn = sqlite3.connect(self.db_name)
                cursor = conn.cursor()
                if ot_filter:
                    cursor.execute("SELECT entry_date, serial_number, ot_number, overall_status, ilrl_details_json, geo_details_json FROM cable_verifications WHERE ot_number LIKE ? ORDER BY entry_date DESC", (f'%{ot_filter}%',))
                else:
                    cursor.execute("SELECT entry_date, serial_number, ot_number, overall_status, ilrl_details_json, geo_details_json FROM cable_verifications ORDER BY entry_date DESC")
                
                records = cursor.fetchall()
                for record in records:
                    entry_date, serial_number, ot_number, overall_status, ilrl_details_json, geo_details_json = record
                    
                    # Store full record for details access later
                    item_id = tree.insert("", tk.END, values=(entry_date, serial_number, ot_number, overall_status),
                                tags=('aprobado' if overall_status == 'APROBADO' else 'rechazado' if overall_status == 'RECHAZADO' else 'no_encontrado', serial_number))
                    
                    # Store details as a dictionary in the item's `open_data` for easy retrieval
                    tree.item(item_id, open_data={
                        'ilrl_details': json.loads(ilrl_details_json) if ilrl_details_json else None,
                        'geo_details': json.loads(geo_details_json) if geo_details_json else None,
                        'serial_number': serial_number, # Pass serial number for combined details window
                        'ot_number': ot_number,
                        'overall_status': overall_status
                    })
                                         
            except sqlite3.Error as e:
                messagebox.showerror("Error de Base de Datos", f"No se pudieron cargar los registros: {e}")
            finally:
                if conn:
                    conn.close()

        search_button = ttk.Button(filter_frame, text="Buscar", command=lambda: cargar_registros(ot_filter_entry.get().strip().upper()), style="Primary.TButton")
        search_button.pack(side=tk.LEFT)

        # Tabla de registros
        tree_frame = ttk.Frame(main_frame, style="Result.TFrame")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(tree_frame, columns=("Fecha", "Serie", "OT", "Estatus General"), show="headings", height=15)
        tree.heading("Fecha", text="Fecha Ingreso", anchor=tk.W)
        tree.heading("Serie", text="Número de Serie", anchor=tk.W)
        tree.heading("OT", text="Orden de Trabajo", anchor=tk.W)
        tree.heading("Estatus General", text="Estatus General", anchor=tk.W)

        tree.column("Fecha", width=150, stretch=tk.NO)
        tree.column("Serie", width=150, stretch=tk.NO)
        tree.column("OT", width=150, stretch=tk.NO)
        tree.column("Estatus General", width=120, stretch=tk.NO)

        # Configurar estilos de tags para Treeview
        tree.tag_configure('aprobado', foreground='green', font=('Arial', 9, 'bold'))
        tree.tag_configure('rechazado', foreground='red', font=('Arial', 9, 'bold'))
        tree.tag_configure('no_encontrado', foreground='orange', font=('Arial', 9, 'bold'))
        
        # Binding para el click en el número de serie
        def on_serial_number_click(event):
            item_id = tree.identify_row(event.y)
            if not item_id:
                return
            
            # Check if the clicked column is the "Serie" column (column index 1)
            column_id = tree.identify_column(event.x)
            if tree.column(column_id, 'id') == tree.column("#2", 'id'): # Column #2 is "Serie"
                data = tree.item(item_id, 'open_data')
                
                # Show combined details in a new window using the stored data
                self._show_combined_details(
                    serial_number=data['serial_number'],
                    ot_number=data['ot_number'],
                    overall_status=data['overall_status'],
                    ilrl_details=data['ilrl_details'],
                    geo_details=data['geo_details']
                )

        tree.bind("<Button-1>", on_serial_number_click)


        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        tree.configure(xscrollcommand=scrollbar_x.set)

        cargar_registros() # Cargar todos los registros al abrir la ventana

        registros_window.mainloop()

    def _show_combined_details(self, serial_number, ot_number, overall_status, ilrl_details, geo_details):
        """Muestra una ventana combinada de detalles ILRL y Geometría."""
        combined_details_window = tk.Toplevel(self.root)
        combined_details_window.title(f"Detalles del Cable: {serial_number}")
        combined_details_window.geometry("900x700")
        combined_details_window.transient(self.root)
        combined_details_window.grab_set()

        main_frame = ttk.Frame(combined_details_window, padding=(20, 20), style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text=f"Detalles para el Cable: {serial_number}", 
                  font=("Arial", 14, "bold"), foreground="#2C3E50").pack(pady=(0, 10))
        ttk.Label(main_frame, text=f"OT: {ot_number} | Estatus General: {overall_status}",
                  font=("Arial", 11, "bold"), foreground="#3498DB").pack(pady=(0, 15))

        # Pestañas para ILRL y Geometría
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # Pestaña ILRL
        ilrl_tab = ttk.Frame(notebook, style="TFrame")
        notebook.add(ilrl_tab, text="Detalles ILRL")
        self._populate_ilrl_details_tab(ilrl_tab, ilrl_details)

        # Pestaña Geometría
        geo_tab = ttk.Frame(notebook, style="TFrame")
        notebook.add(geo_tab, text="Detalles Geometría")
        self._populate_geo_details_tab(geo_tab, geo_details)
        
        combined_details_window.mainloop()

    def _populate_ilrl_details_tab(self, parent_frame, ilrl_details):
        """Pobla la pestaña de detalles ILRL."""
        if not ilrl_details:
            ttk.Label(parent_frame, text="No hay datos de ILRL disponibles para este registro.", font=("Arial", 10), foreground="#6C757D", background="#F0F4F8").pack(pady=20)
            return

        ttk.Label(parent_frame, text="📁 Archivo Analizado:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(10, 5))
        ttk.Label(parent_frame, text=ilrl_details.get('file_path', 'N/A'), wraplength=800, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

        ttk.Label(parent_frame, text="📈 Resultado General ILRL:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        
        resultado_general = ilrl_details.get('resultado_general', 'N/A')
        fecha_general = ilrl_details.get('fecha_general', 'N/A')
        color = "green" if resultado_general == "APROBADO" else "red"
        
        info_label = ttk.Label(parent_frame, text=f"{resultado_general} (Fecha de medición más reciente: {fecha_general})", 
                               font=("Arial", 10, "bold"), foreground=color, background="#F0F4F8")
        info_label.pack(anchor="w", pady=(0, 10))

        ttk.Label(parent_frame, text="📊 Mediciones Detalladas por Línea:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))

        tree = ttk.Treeview(parent_frame, columns=("Línea", "Resultado", "Fecha"), show="headings", height=10)
        tree.heading("Línea", text="Línea", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha", anchor=tk.W)

        tree.column("Línea", width=70, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=180, stretch=tk.NO)

        detalles_lineas = ilrl_details.get('detalles_lineas', [])
        for detalle in detalles_lineas:
            resultado = detalle.get('resultado', 'N/A')
            tree.insert("", tk.END, values=(detalle.get('linea', 'N/A'), resultado, detalle.get('fecha', 'N/A')), 
                        tags=('pass_style' if resultado == 'PASS' else 'fail_style'))
        
        tree.tag_configure('pass_style', foreground='green')
        tree.tag_configure('fail_style', foreground='red')

        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

    def _populate_geo_details_tab(self, parent_frame, geo_details):
        """Pobla la pestaña de detalles de Geometría."""
        if not geo_details:
            ttk.Label(parent_frame, text="No hay datos de Geometría disponibles para este registro.", font=("Arial", 10), foreground="#6C757D", background="#F0F4F8").pack(pady=20)
            return

        ttk.Label(parent_frame, text="📁 Archivo Analizado:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(10, 5))
        ttk.Label(parent_frame, text=geo_details.get('file_path', 'N/A'), wraplength=800, font=("Arial", 9), foreground="#6C757D", background="#F0F4F8").pack(anchor="w", pady=(0, 10))

        ttk.Label(parent_frame, text=f"📈 Resultado General para Geometría:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))
        
        resultado_general = geo_details.get('resultado_general', 'N/A')
        fecha_general = geo_details.get('fecha_general', 'N/A')
        color = "green" if resultado_general == "APROBADO" else "red"
        
        info_label = ttk.Label(parent_frame, text=f"{resultado_general} (Fecha de medición más reciente: {fecha_general})", 
                               font=("Arial", 10, "bold"), foreground=color, background="#F0F4F8")
        info_label.pack(anchor="w", pady=(0, 10))

        ttk.Label(parent_frame, text="📐 Mediciones Detalladas por Punta:", font=("Arial", 10, "bold"), foreground="#2C3E50", background="#F0F4F8").pack(anchor="w", pady=(0, 5))

        tree = ttk.Treeview(parent_frame, columns=("Serie", "Punta", "Resultado", "Fecha"), show="headings", height=10)
        tree.heading("Serie", text="Serie", anchor=tk.W)
        tree.heading("Punta", text="Punta", anchor=tk.W)
        tree.heading("Resultado", text="Resultado", anchor=tk.W)
        tree.heading("Fecha", text="Fecha y Hora", anchor=tk.W)

        tree.column("Serie", width=120, stretch=tk.NO)
        tree.column("Punta", width=70, stretch=tk.NO)
        tree.column("Resultado", width=100, stretch=tk.NO)
        tree.column("Fecha", width=180, stretch=tk.NO)

        detalles_puntas = geo_details.get('detalles_puntas', [])
        for detalle in detalles_puntas:
            resultado = detalle.get('resultado', 'N/A')
            tree.insert("", tk.END, values=(detalle.get('serie', 'N/A'), detalle.get('punta', 'N/A'), resultado, detalle.get('timestamp', 'N/A')), 
                        tags=('pass_style' if resultado == 'PASS' else 'fail_style'))
        
        tree.tag_configure('pass_style', foreground='green')
        tree.tag_configure('fail_style', foreground='red')

        tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

    def iniciar(self):
        self.root = tk.Tk()
        self.root.title("Verificador de Estado de Cables - Versión 1.2")
        self.root.geometry("800x650")
        self.root.configure(bg="#F0F4F8")
        
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configuración", menu=config_menu)
        config_menu.add_command(label="Cambiar Rutas", command=self.solicitar_contrasena)
        config_menu.add_command(label="Ver Registros", command=self.solicitar_contrasena_registros) # Nueva opción de menú
        
        self.style = ttk.Style()
        
        self.style.configure(".", background="#F0F4F8", font=("Arial", 10))
        self.style.configure("TFrame", background="#F0F4F8")
        self.style.configure("TLabel", background="#F0F4F8", foreground="#2C3E50")
        self.style.configure("TButton", font=("Arial", 10, "bold"), padding=8, relief="flat", borderwidth=0)
        self.style.map("TButton", 
                       background=[('active', "#AD8662"), ('!disabled', "#DB3F34")],
                       foreground=[('active', 'black'), ('!disabled', 'black')])

        self.style.configure("Primary.TButton", 
                           background="#28A745",
                           foreground="white",
                           font=("Arial", 11, "bold"),
                           borderwidth=0,
                           focusthickness=2,
                           focuscolor="#28A745")
        self.style.map("Primary.TButton", 
                       background=[('active', '#218838'), ('!disabled', '#28A745')])
        
        self.style.configure("TEntry", 
                           fieldbackground="white", 
                           foreground="#2C3E50", 
                           borderwidth=1, 
                           relief="solid")
        
        self.style.configure("Card.TFrame", 
                           background="#FFFFFF",
                           relief="flat", 
                           borderwidth=1, 
                           bordercolor="#E0E0E0")
        
        self.style.configure("Result.TFrame", 
                           background="#FFFFFF",
                           relief="flat",
                           borderwidth=1, 
                           bordercolor="#E0E0E0")
        
        self.style.configure("Path.TLabel", 
                           font=("Arial", 9), 
                           foreground="#6C757D",
                           background="#F8F9FA",
                           padding=(5,2))

        self.style.configure("Input.TFrame",
                            background="#FFFFFF",
                            relief="flat",
                            borderwidth=1,
                            bordercolor="#E0E0E0")

        self.style.configure("Treeview.Heading", font=("Arial", 9, "bold"), background="#E0E0E0", foreground="#2C3E50")
        self.style.configure("Treeview", font=("Arial", 9), rowheight=25)
        self.style.map("Treeview", background=[('selected', '#B0D7FF')])
        
        main_frame = ttk.Frame(self.root, padding=(20, 15), style="TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_frame = ttk.Frame(main_frame, style="TFrame")
        title_frame.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        ttk.Label(title_frame, 
                text="🔍 Verificador de Cables de Fibra Óptica", 
                font=("Arial", 18, "bold"), 
                foreground="#2C3E50",
                background="#F0F4F8").pack()
        
        ttk.Label(title_frame, 
                text="Sistema de verificación de resultados ILRL y Geometría", 
                font=("Arial", 11), 
                foreground="#3498DB",
                background="#F0F4F8").pack()
        
        input_frame = ttk.Frame(main_frame, padding=(20, 15), style="Input.TFrame")
        input_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Label(input_frame, text="📋 Orden de Trabajo (ej. JMO-250500001):", 
                 font=("Arial", 10, "bold"), foreground="#2C3E50").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ot_entry = ttk.Entry(input_frame, width=30, font=("Arial", 10), style="TEntry")
        self.ot_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=10)
        
        ttk.Label(input_frame, text="🔢 Número de Serie (13 dígitos):", 
                 font=("Arial", 10, "bold"), foreground="#2C3E50").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.serie_entry = ttk.Entry(input_frame, width=30, font=("Arial", 10), style="TEntry")
        self.serie_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=10)
        self.serie_entry.bind("<KeyRelease>", self.verificar_cable_automatico)
        
        input_frame.columnconfigure(1, weight=1)
        
        path_frame = ttk.Frame(main_frame, padding=(15, 10), style="Card.TFrame")
        path_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Label(path_frame, 
                 text="🔎 RUTAS DE ANÁLISIS", 
                 font=("Arial", 10, "bold"), 
                 foreground="#3498DB",
                 background="#FFFFFF").pack(anchor="w", pady=(0, 5))
        
        self.ruta_ilrl_label = ttk.Label(path_frame, 
                                       text=f"📂 Ruta ILRL: {self.ruta_base_ilrl}",
                                       style="Path.TLabel")
        self.ruta_ilrl_label.pack(anchor="w", padx=5, pady=2, fill="x")
        
        self.ruta_geo_label = ttk.Label(path_frame, 
                                      text=f"📂 Ruta Geometría: {self.ruta_base_geo}",
                                      style="Path.TLabel")
        self.ruta_geo_label.pack(anchor="w", padx=5, pady=2, fill="x")
        
        result_frame = ttk.Frame(main_frame, style="Result.TFrame", padding=15)
        result_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky="nsew")
        
        ttk.Label(result_frame, 
                text="📋 RESULTADOS DE LA VERIFICACIÓN", 
                font=("Arial", 11, "bold"), 
                foreground="#2C3E50",
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
        
        self.resultado_text.tag_config("normal", foreground="#2C3E50")
        self.resultado_text.tag_config("bold", font=("Arial", 10, "bold"), foreground="#2C3E50")
        self.resultado_text.tag_config("header", 
                                     font=("Arial", 13, "bold"), 
                                     foreground="#2C3E50",
                                     justify="center")
        self.resultado_text.tag_config("verde", 
                                     foreground="#28A745",
                                     font=("Arial", 10, "bold"))
        self.resultado_text.tag_config("rojo", 
                                     foreground="#DC3545",
                                     font=("Arial", 10, "bold"))
        self.resultado_text.tag_config("orange", # Nuevo tag para "NO ENCONTRADO"
                                     foreground="#FFA500",
                                     font=("Arial", 10, "bold"))
        self.resultado_text.tag_config("ilrl_click", underline=1, font=("Arial", 10, "bold"))
        self.resultado_text.tag_config("geo_click", underline=1, font=("Arial", 10, "bold"))
        
        self.resultado_text.insert(tk.END, "Bienvenido al Verificador de Cables.\n\n"
                                     "Ingrese la Orden de Trabajo y el Número de Serie para iniciar.\n"
                                     "La verificación se realizará automáticamente al completar los 13 dígitos del número de serie.\n\n"
                                     "Haga click en los resultados 'APROBADO' o 'RECHAZADO' de ILRL o Geometría para ver los detalles.", "normal")
        self.resultado_text.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(result_frame, 
                                command=self.resultado_text.yview,
                                style="TScrollbar")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.resultado_text.config(yscrollcommand=scrollbar.set)
        
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
                foreground="#6C757D",
                background="#FFFFFF").pack(anchor="w")

        button_exit_frame = ttk.Frame(main_frame, style="TFrame")
        button_exit_frame.grid(row=5, column=0, columnspan=2, pady=(15, 5))
        
        exit_button = ttk.Button(button_exit_frame, 
                                 text="🚫 Salir del Programa", 
                                 command=self.root.destroy, 
                                 style="TButton")
        exit_button.pack(pady=5, ipadx=10, ipady=5)
        
        footer_frame = ttk.Frame(main_frame, style="TFrame")
        footer_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Label(footer_frame, 
                text="Sistema de Verificación de Cables v1.2 | Desarrollado por Paulo", 
                font=("Arial", 8), 
                foreground="#6C757D",
                background="#F0F4F8").pack()
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        self.root.mainloop()

if __name__ == "__main__":
    app = VerificadorCables()
    app.iniciar()