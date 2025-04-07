import tkinter as tk
from tkinter import messagebox, ttk, Canvas, filedialog
import csv
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from PIL import Image, ImageTk

# Variables globales
trabajos = []
ID = 1
perfiles = {}
proveedor_global = ""

# Ruta del archivo perfiles.csv
ruta_csv = r"C:\Users\khriz\Desktop\Optimizacion_corte\perfiles.csv"

# Configuraciones de aplicaciones con perfiles asociados y ángulos de corte
CONFIG_APLICACIONES = {
    "Oscilobatiente": {
        "imagen": "oscilobatiente.png",
        "descuento": 80,
        "perfil_marco": "508001",
        "perfil_hoja": "508002",
        "perfil_travesano": None,
        "angulos": {
            "marco": {
                "ancho": (45, 135),  # (izquierda, derecha)
                "largo": (45, 135),  # (superior, inferior)
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            },
            "hoja": {
                "ancho": (45, 135),
                "largo": (45, 135),
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            }
        }
    },
    "Practicable": {
        "imagen": "practicable.png",
        "descuento": 80,
        "perfil_marco": "508001",
        "perfil_hoja": "508002",
        "perfil_travesano": None,
        "angulos": {
            "marco": {
                "ancho": (45, 135),
                "largo": (45, 135),
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            },
            "hoja": {
                "ancho": (45, 135),
                "largo": (45, 135),
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            }
        }
    },
    "Fija": {
        "imagen": "fija.png",
        "descuento": 0,
        "perfil_marco": "508001",
        "perfil_hoja": None,
        "perfil_travesano": None,
        "angulos": {
            "marco": {
                "ancho": (45, 135),
                "largo": (45, 135),
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            }
        }
    },
    "Balconera": {
        "imagen": "balconera.png",
        "descuento": 96,
        "perfil_marco": "500373",
        "perfil_hoja": "541170",
        "perfil_travesano": None,
        "angulos": {
            "marco": {
                "ancho": (45, 90),
                "largo": (90, 45),
                "aplicar": ["superior", "izquierda", "derecha"]
            },
            "hoja": {
                "ancho": (45, 135),
                "largo": (45, 135),
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            }
        }
    },
    "Corrediza": {
        "imagen": "corrediza.png",
        "descuento": 0,
        "perfil_marco": "500004",
        "perfil_hoja": "500033",
        "perfil_travesano": "500304",
        "angulos": {
            "marco": {
                "ancho": (45, 135),
                "largo": (45, 135),
                "aplicar": ["superior", "inferior", "izquierda", "derecha"]
            }
        }
    }
}

# Cargar los perfiles desde el archivo CSV
if os.path.exists(ruta_csv):
    with open(ruta_csv, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            codigo = row["Codigo"].strip()
            descripcion = row["Descripcion"].strip()
            alto = row["Alto (mm)"].strip()
            ancho = row["Ancho (mm)"].strip()
            refuerzo = row["Refuerzo (mm)"].strip()
            long_barra = row["Long_barra"].strip()
            perfiles[codigo] = {
                "codigo": codigo,
                "descripcion": descripcion,
                "alto": alto,
                "ancho": ancho,
                "refuerzo": refuerzo,
                "long_barra": long_barra
            }
else:
    messagebox.showerror("Error", f"No se encontró el archivo {ruta_csv}")

def obtener_codigo_por_descripcion(descripcion):
    """Busca el código de perfil por su descripción completa"""
    for codigo, datos in perfiles.items():
        if datos["descripcion"] == descripcion:
            return codigo
    return None

def agregar_trabajo(job_para_editar=None):
    global ID, trabajos
    
    # Inicializar variables
    cortes_marco = []
    cortes_hoja = []
    cortes_travesano = []
    es_edicion = False
    id_trabajo = ID
    
    if job_para_editar:
        es_edicion = True
        id_trabajo = job_para_editar[0]
        cliente_actual = job_para_editar[1]
        nombre_actual = job_para_editar[2]
        cortes_marco = list(job_para_editar[6])  # Marco
        cortes_hoja = list(job_para_editar[7])   # Hoja
        cortes_travesano = list(job_para_editar[8]) if len(job_para_editar) > 8 else []  # Travesaño
        aplicacion_actual = job_para_editar[9] if len(job_para_editar) > 9 else ""
    else:
        cliente_actual = trabajos[-1][1] if trabajos else ""
        nombre_actual = trabajos[-1][2] if trabajos else ""
        aplicacion_actual = ""

    def dibujar_angulo(canvas, x, y, angulo, posicion, color="red", tamaño=15):
        """Dibuja un ángulo en la posición especificada"""
        if angulo == 45:
            if posicion in ["izq_sup", "der_inf"]:
                canvas.create_line(x, y, x + tamaño, y + tamaño, width=2, fill=color)
            else:
                canvas.create_line(x, y, x - tamaño, y + tamaño, width=2, fill=color)
        elif angulo == 135:
            if posicion in ["der_sup", "izq_inf"]:
                canvas.create_line(x, y, x - tamaño, y + tamaño, width=2, fill=color)
            else:
                canvas.create_line(x, y, x + tamaño, y + tamaño, width=2, fill=color)
        elif angulo == 90:
            if posicion.startswith("izq"):
                canvas.create_line(x, y, x + tamaño, y, width=2, fill=color)
            elif posicion.startswith("der"):
                canvas.create_line(x, y, x - tamaño, y, width=2, fill=color)
            if posicion.endswith("sup"):
                canvas.create_line(x, y, x, y + tamaño, width=2, fill=color)
            elif posicion.endswith("inf"):
                canvas.create_line(x, y, x, y - tamaño, width=2, fill=color)

    def dibujar_aplicacion_dinamica():
        """Replica exactamente las imágenes originales de cada aplicación en el canvas"""
        try:
            # Obtener medidas
            ancho = int(entry_medida_horizontal.get()) if entry_medida_horizontal.get() else 500
            alto = int(entry_medida_vertical.get()) if entry_medida_vertical.get() else 500
            aplicacion = combo_aplicacion.get()
            config = CONFIG_APLICACIONES.get(aplicacion, {})
            
            # Limpiar canvas
            canvas_visualizacion.delete("all")
            
            # Tamaño fijo del canvas (300x200 como las imágenes originales)
            canvas_width = 300
            canvas_height = 200
            
            # Calcular escala manteniendo relación de aspecto
            escala_ancho = (canvas_width - 40) / max(ancho, 1)
            escala_alto = (canvas_height - 40) / max(alto, 1)
            escala = min(escala_ancho, escala_alto) * 0.85  # Pequeño margen
            
            # Centrar el dibujo
            margen_x = (canvas_width - (ancho * escala)) / 2
            margen_y = (canvas_height - (alto * escala)) / 2
            
            # Estilo común para todos los tipos
            grosor_marco = 3
            color_marco = "#2c3e50"
            color_relleno_marco = "#ecf0f1"
            color_hoja = "#3498db"
            color_relleno_hoja = "#d6eaf8"
            
            # Dibujar según el tipo de aplicación (réplica exacta de las imágenes)
            if aplicacion == "Oscilobatiente":
                # Marco con esquinas redondeadas (como en la imagen)
                radio_esquina = 8
                puntos_marco = [
                    margen_x + radio_esquina, margen_y,
                    margen_x + ancho*escala - radio_esquina, margen_y,
                    margen_x + ancho*escala, margen_y + radio_esquina,
                    margen_x + ancho*escala, margen_y + alto*escala - radio_esquina,
                    margen_x + ancho*escala - radio_esquina, margen_y + alto*escala,
                    margen_x + radio_esquina, margen_y + alto*escala,
                    margen_x, margen_y + alto*escala - radio_esquina,
                    margen_x, margen_y + radio_esquina,
                    margen_x + radio_esquina, margen_y
                ]
                canvas_visualizacion.create_polygon(
                    puntos_marco,
                    outline=color_marco, width=grosor_marco,
                    fill=color_relleno_marco, smooth=True
                )
                
                # Hoja con batiente inferior (como en la imagen)
                hoja_x1 = margen_x + 15
                hoja_y1 = margen_y + 15
                hoja_x2 = margen_x + ancho*escala - 15
                hoja_y2 = margen_y + alto*escala - 25
                
                # Parte principal de la hoja
                canvas_visualizacion.create_rectangle(
                    hoja_x1, hoja_y1, hoja_x2, hoja_y2,
                    outline=color_hoja, width=2, fill=color_relleno_hoja
                )
                
                # Batiente inferior (más grueso como en la imagen)
                batiente_y1 = hoja_y2 - 5
                batiente_y2 = hoja_y2 + 15
                canvas_visualizacion.create_rectangle(
                    hoja_x1, batiente_y1, hoja_x2, batiente_y2,
                    outline=color_hoja, width=2, fill=color_relleno_hoja
                )
                
                # Bisagra característica (como en la imagen)
                canvas_visualizacion.create_line(
                    hoja_x1 + 10, margen_y + alto*escala - 10,
                    hoja_x1 + 10, batiente_y2 - 2,
                    fill="#7f8c8d", width=3
                )
                
                # Manija (como en la imagen)
                canvas_visualizacion.create_oval(
                    hoja_x2 - 20, hoja_y1 + 30,
                    hoja_x2 - 10, hoja_y1 + 40,
                    fill="#bdc3c7", outline="#7f8c8d"
                )

            elif aplicacion == "Practicable":
                # Marco rectangular simple (como en la imagen)
                canvas_visualizacion.create_rectangle(
                    margen_x, margen_y,
                    margen_x + ancho*escala, margen_y + alto*escala,
                    outline=color_marco, width=grosor_marco,
                    fill=color_relleno_marco
                )
                
                # Hoja con sombra 3D (como en la imagen)
                hoja_x1 = margen_x + 12
                hoja_y1 = margen_y + 12
                hoja_x2 = margen_x + ancho*escala - 12
                hoja_y2 = margen_y + alto*escala - 12
                
                # Efecto 3D con gradiente (simulado)
                for i in range(3):
                    offset = i * 2
                    canvas_visualizacion.create_rectangle(
                        hoja_x1 + offset, hoja_y1 + offset,
                        hoja_x2 - offset, hoja_y2 - offset,
                        outline=color_hoja, width=1,
                        fill=color_relleno_hoja if i == 2 else ""
                    )
                
                # Manija prominente (como en la imagen)
                canvas_visualizacion.create_rectangle(
                    hoja_x2 - 25, hoja_y1 + 30,
                    hoja_x2 - 5, hoja_y1 + 50,
                    fill="#bdc3c7", outline="#7f8c8d", width=2
                )

            elif aplicacion == "Fija":
                # Marco con vidrio (como en la imagen)
                canvas_visualizacion.create_rectangle(
                    margen_x, margen_y,
                    margen_x + ancho*escala, margen_y + alto*escala,
                    outline=color_marco, width=grosor_marco,
                    fill=color_relleno_marco
                )
                
                # Líneas divisorias de vidrio (como en la imagen)
                divisiones = 3  # Número de divisiones como en la imagen
                for i in range(1, divisiones):
                    # Verticales
                    canvas_visualizacion.create_line(
                        margen_x + (ancho*escala/divisiones)*i, margen_y + 5,
                        margen_x + (ancho*escala/divisiones)*i, margen_y + alto*escala - 5,
                        fill="#bdc3c7", width=1
                    )
                    # Horizontales
                    canvas_visualizacion.create_line(
                        margen_x + 5, margen_y + (alto*escala/divisiones)*i,
                        margen_x + ancho*escala - 5, margen_y + (alto*escala/divisiones)*i,
                        fill="#bdc3c7", width=1
                    )
                
                # Puntos de fijación (como en la imagen)
                for x in [margen_x + 20, margen_x + ancho*escala - 20]:
                    for y in [margen_y + 20, margen_y + alto*escala - 20]:
                        canvas_visualizacion.create_oval(
                            x - 3, y - 3, x + 3, y + 3,
                            fill="#e74c3c", outline=""
                        )

            elif aplicacion == "Balconera":
                # Marco con forma característica (como en la imagen)
                puntos_marco = [
                    margen_x, margen_y + 20,
                    margen_x + 20, margen_y,
                    margen_x + ancho*escala - 20, margen_y,
                    margen_x + ancho*escala, margen_y + 20,
                    margen_x + ancho*escala, margen_y + alto*escala,
                    margen_x, margen_y + alto*escala,
                    margen_x, margen_y + 20
                ]
                canvas_visualizacion.create_polygon(
                    puntos_marco,
                    outline=color_marco, width=grosor_marco,
                    fill=color_relleno_marco
                )
                
                # Hoja inclinada (como en la imagen)
                puntos_hoja = [
                    margen_x + 25, margen_y + 25,
                    margen_x + ancho*escala - 25, margen_y + 25,
                    margen_x + ancho*escala - 35, margen_y + alto*escala - 25,
                    margen_x + 35, margen_y + alto*escala - 25,
                    margen_x + 25, margen_y + 25
                ]
                canvas_visualizacion.create_polygon(
                    puntos_hoja,
                    outline=color_hoja, width=2,
                    fill=color_relleno_hoja
                )
                
                # Manija vertical (como en la imagen)
                canvas_visualizacion.create_line(
                    margen_x + ancho*escala - 30, margen_y + 40,
                    margen_x + ancho*escala - 30, margen_y + alto*escala - 40,
                    fill="#7f8c8d", width=3
                )

            elif aplicacion == "Corrediza":
                # Marco con riel inferior destacado (como en la imagen)
                canvas_visualizacion.create_rectangle(
                    margen_x, margen_y,
                    margen_x + ancho*escala, margen_y + alto*escala,
                    outline=color_marco, width=grosor_marco,
                    fill=color_relleno_marco
                )
                
                # Riel inferior grueso (como en la imagen)
                canvas_visualizacion.create_rectangle(
                    margen_x + 5, margen_y + alto*escala - 15,
                    margen_x + ancho*escala - 5, margen_y + alto*escala - 5,
                    fill="#95a5a6", outline="#7f8c8d", width=2
                )
                
                # Dos hojas corredizas (como en la imagen)
                hoja1_x2 = margen_x + ancho*escala * 0.6
                hoja2_x1 = margen_x + ancho*escala * 0.4
                
                # Hoja 1 (trasera)
                canvas_visualizacion.create_rectangle(
                    margen_x + 10, margen_y + 10,
                    hoja1_x2, margen_y + alto*escala - 20,
                    outline=color_hoja, width=1, fill="#a5d8ff"
                )
                
                # Hoja 2 (delantera)
                canvas_visualizacion.create_rectangle(
                    hoja2_x1, margen_y + 15,
                    margen_x + ancho*escala - 10, margen_y + alto*escala - 15,
                    outline=color_hoja, width=2, fill=color_relleno_hoja
                )
                
                # Manijas (como en la imagen)
                canvas_visualizacion.create_rectangle(
                    hoja1_x2 - 20, margen_y + alto*escala//2 - 15,
                    hoja1_x2 - 5, margen_y + alto*escala//2 + 15,
                    fill="#bdc3c7", outline="#7f8c8d", width=1
                )
                canvas_visualizacion.create_rectangle(
                    hoja2_x1 + 5, margen_y + alto*escala//2 - 15,
                    hoja2_x1 + 20, margen_y + alto*escala//2 + 15,
                    fill="#bdc3c7", outline="#7f8c8d", width=1
                )

            # Dibujar ángulos según configuración (encima de la representación)
            dibujar_angulos_marco(config, margen_x, margen_y, 
                                margen_x + ancho*escala, margen_y + alto*escala, 
                                escala)
            
            if config.get("perfil_hoja"):
                # Coordenadas aproximadas de la hoja para los ángulos
                hoja_x1 = margen_x + 15
                hoja_y1 = margen_y + 15
                hoja_x2 = margen_x + ancho*escala - 15
                hoja_y2 = margen_y + alto*escala - 15
                dibujar_angulos_hoja(config, hoja_x1, hoja_y1, hoja_x2, hoja_y2, escala*0.7)
            
            # Mostrar medidas actuales (como en las imágenes originales)
            canvas_visualizacion.create_text(
                canvas_width/2, canvas_height-15,
                text=f"{ancho} x {alto} mm",
                font=("Arial", 9, "bold"),
                fill=color_marco
            )

        except ValueError:
            pass

    def dibujar_angulos_marco(config, x1, y1, x2, y2, escala):
        """Dibuja los ángulos del marco superpuestos a la representación visual"""
        angulos = config.get("angulos", {}).get("marco", {})
        tamaño = max(10, min(15, escala * 12))
        
        # Color rojo semitransparente para los ángulos
        color_angulo = "#e74c3c"
        
        if "superior" in angulos.get("aplicar", []):
            angulo_izq, angulo_der = angulos.get("ancho", (90, 90))
            dibujar_angulo(canvas_visualizacion, x1, y1, angulo_izq, "izq_sup", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x2, y1, angulo_der, "der_sup", color_angulo, tamaño)
        
        if "inferior" in angulos.get("aplicar", []):
            angulo_izq, angulo_der = angulos.get("ancho", (90, 90))
            dibujar_angulo(canvas_visualizacion, x1, y2, angulo_izq, "izq_inf", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x2, y2, angulo_der, "der_inf", color_angulo, tamaño)
        
        if "izquierda" in angulos.get("aplicar", []):
            angulo_sup, angulo_inf = angulos.get("largo", (90, 90))
            dibujar_angulo(canvas_visualizacion, x1, y1, angulo_sup, "izq_sup", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x1, y2, angulo_inf, "izq_inf", color_angulo, tamaño)
        
        if "derecha" in angulos.get("aplicar", []):
            angulo_sup, angulo_inf = angulos.get("largo", (90, 90))
            dibujar_angulo(canvas_visualizacion, x2, y1, angulo_sup, "der_sup", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x2, y2, angulo_inf, "der_inf", color_angulo, tamaño)

    def dibujar_angulos_hoja(config, x1, y1, x2, y2, escala):
        """Dibuja los ángulos de la hoja superpuestos a la representación visual"""
        if not config.get("perfil_hoja"):
            return
            
        angulos = config.get("angulos", {}).get("hoja", {})
        tamaño = max(8, min(12, escala * 10))
        color_angulo = "#2980b9"  # Azul para ángulos de hoja
        
        if "superior" in angulos.get("aplicar", []):
            angulo_izq, angulo_der = angulos.get("ancho", (90, 90))
            dibujar_angulo(canvas_visualizacion, x1, y1, angulo_izq, "izq_sup", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x2, y1, angulo_der, "der_sup", color_angulo, tamaño)
        
        if "inferior" in angulos.get("aplicar", []):
            angulo_izq, angulo_der = angulos.get("ancho", (90, 90))
            dibujar_angulo(canvas_visualizacion, x1, y2, angulo_izq, "izq_inf", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x2, y2, angulo_der, "der_inf", color_angulo, tamaño)
        
        if "izquierda" in angulos.get("aplicar", []):
            angulo_sup, angulo_inf = angulos.get("largo", (90, 90))
            dibujar_angulo(canvas_visualizacion, x1, y1, angulo_sup, "izq_sup", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x1, y2, angulo_inf, "izq_inf", color_angulo, tamaño)
        
        if "derecha" in angulos.get("aplicar", []):
            angulo_sup, angulo_inf = angulos.get("largo", (90, 90))
            dibujar_angulo(canvas_visualizacion, x2, y1, angulo_sup, "der_sup", color_angulo, tamaño)
            dibujar_angulo(canvas_visualizacion, x2, y2, angulo_inf, "der_inf", color_angulo, tamaño)

    def agregar_cortes_segun_orientacion(lista_cortes, medida_h, medida_v, tipo_perfil, config, listbox):
        tipo_str = tipo_perfil
        tag = "marco" if tipo_perfil == "Marco" else "hoja" if tipo_perfil == "Hoja" else "travesano"
        
        if tipo_perfil == "Hoja" and combo_aplicacion.get() == "Balconera":
            # Caso especial para hoja de Balconera - siempre 4 cortes
            descuento = config["descuento"]
            medida_h_hoja = medida_h - descuento
            medida_v_hoja = medida_v - descuento
            
            # Ángulos fijos para hoja de Balconera
            cortes_hoja.append((medida_h_hoja, "SUP", 45, 135))
            listbox.insert(tk.END, f"{medida_h_hoja} mm - SUP (Hoja) - 45°/135°", tag)
            
            cortes_hoja.append((medida_h_hoja, "INF", 45, 135))
            listbox.insert(tk.END, f"{medida_h_hoja} mm - INF (Hoja) - 45°/135°", tag)
            
            cortes_hoja.append((medida_v_hoja, "IZQ", 45, 135))
            listbox.insert(tk.END, f"{medida_v_hoja} mm - IZQ (Hoja) - 45°/135°", tag)
            
            cortes_hoja.append((medida_v_hoja, "DER", 45, 135))
            listbox.insert(tk.END, f"{medida_v_hoja} mm - DER (Hoja) - 45°/135°", tag)
        else:
            # Comportamiento normal para otros casos
            if medida_h > 0:
                if aplicar_superior:
                    corte_data = (medida_h, "SUP", angulo_sup_izq, angulo_sup_der)
                    lista_cortes.append(corte_data)
                    listbox.insert(tk.END, f"{medida_h} mm - SUP ({tipo_str}) - {corte_data[2]}°/{corte_data[3]}°", tag)
                    
                if aplicar_inferior:
                    corte_data = (medida_h, "INF", angulo_inf_izq, angulo_inf_der)
                    lista_cortes.append(corte_data)
                    listbox.insert(tk.END, f"{medida_h} mm - INF ({tipo_str}) - {corte_data[2]}°/{corte_data[3]}°", tag)
        
            if medida_v > 0:
                if aplicar_izquierda:
                    corte_data = (medida_v, "IZQ", angulo_izq_sup, angulo_izq_inf)
                    lista_cortes.append(corte_data)
                    listbox.insert(tk.END, f"{medida_v} mm - IZQ ({tipo_str}) - {corte_data[2]}°/{corte_data[3]}°", tag)
                
                if aplicar_derecha:
                    corte_data = (medida_v, "DER", angulo_der_sup, angulo_der_inf)
                    lista_cortes.append(corte_data)
                    listbox.insert(tk.END, f"{medida_v} mm - DER ({tipo_str}) - {corte_data[2]}°/{corte_data[3]}°", tag)

    def agregar_corte():
        try:
            # Obtener valores de los campos
            medida_horizontal = entry_medida_horizontal.get()
            medida_vertical = entry_medida_vertical.get()
            
            if not medida_horizontal and not medida_vertical:
                raise ValueError("Debe ingresar al menos una medida")
            
            medida_h = int(medida_horizontal) if medida_horizontal else 0
            medida_v = int(medida_vertical) if medida_vertical else 0
            
            # Validar checkboxes
            nonlocal aplicar_superior, aplicar_inferior, aplicar_izquierda, aplicar_derecha
            aplicar_superior = chk_superior_var.get()
            aplicar_inferior = chk_inferior_var.get()
            aplicar_izquierda = chk_izquierda_var.get()
            aplicar_derecha = chk_derecha_var.get()
            
            if not (aplicar_superior or aplicar_inferior or aplicar_izquierda or aplicar_derecha):
                raise ValueError("Debe seleccionar al menos una posición para el corte")
            
            # Obtener ángulos
            nonlocal angulo_sup_izq, angulo_sup_der, angulo_inf_izq, angulo_inf_der
            nonlocal angulo_izq_sup, angulo_izq_inf, angulo_der_sup, angulo_der_inf
            angulo_sup_izq = int(combo_angulo_sup_izq.get())
            angulo_sup_der = int(combo_angulo_sup_der.get())
            angulo_inf_izq = int(combo_angulo_inf_izq.get())
            angulo_inf_der = int(combo_angulo_inf_der.get())
            angulo_izq_sup = int(combo_angulo_izq_sup.get())
            angulo_izq_inf = int(combo_angulo_izq_inf.get())
            angulo_der_sup = int(combo_angulo_der_sup.get())
            angulo_der_inf = int(combo_angulo_der_inf.get())
            
            aplicacion = combo_aplicacion.get()
            config = CONFIG_APLICACIONES.get(aplicacion, {})
            
            # Validar que se haya seleccionado una aplicación
            if not aplicacion:
                raise ValueError("Debe seleccionar una aplicación primero")
            
            # Agregar cortes de marco
            agregar_cortes_segun_orientacion(
                cortes_marco, medida_h, medida_v, 
                "Marco", config, listbox_cortes
            )
            
            # Calcular y agregar cortes de hoja automáticamente si aplica
            if config.get("perfil_hoja"):
                descuento = config["descuento"]
                medida_h_hoja = medida_h - descuento
                medida_v_hoja = medida_v - descuento
                
                # Validar medidas de hoja
                if medida_h_hoja <= 0 or medida_v_hoja <= 0:
                    raise ValueError("Medidas de hoja inválidas tras aplicar descuento")
                
                agregar_cortes_segun_orientacion(
                    cortes_hoja, medida_h_hoja, medida_v_hoja,
                    "Hoja", config, listbox_cortes
                )

            # Actualizar visualización
            dibujar_aplicacion_dinamica()
            
            # Limpiar campos
            entry_medida_horizontal.delete(0, tk.END)
            entry_medida_vertical.delete(0, tk.END)
            
        except ValueError as e:
            messagebox.showerror("Error", f"Dato inválido: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al agregar corte: {str(e)}")

    def eliminar_corte():
        seleccionado = listbox_cortes.curselection()
        if seleccionado:
            index = seleccionado[0]
            texto_corte = listbox_cortes.get(index)
            
            # Determinar de qué lista eliminar el corte
            if "(Marco)" in texto_corte:
                cortes_marco.pop(index)
            elif "(Hoja)" in texto_corte:
                cortes_hoja.pop(index)
            elif "(Travesaño)" in texto_corte:
                cortes_travesano.pop(index)
                
            listbox_cortes.delete(index)
            dibujar_aplicacion_dinamica()
        else:
            messagebox.showwarning("Advertencia", "Seleccione un corte para eliminar")

    def guardar_trabajo():
        nonlocal id_trabajo, es_edicion
        global ID
        
        try:
            # Validar campos obligatorios
            cliente = entry_cliente.get().strip()
            nombre = entry_nombre.get().strip()
            aplicacion = combo_aplicacion.get()
            
            if not cliente:
                raise ValueError("El campo Cliente es obligatorio")
            if not nombre:
                raise ValueError("El campo Nombre del trabajo es obligatorio")
            if not aplicacion:
                raise ValueError("Debe seleccionar una aplicación")
            if not cortes_marco:
                raise ValueError("Debe agregar al menos un corte de marco")
            
            # Obtener configuración de la aplicación
            config = CONFIG_APLICACIONES.get(aplicacion)
            if not config:
                raise ValueError(f"Aplicación '{aplicacion}' no encontrada en configuraciones")
            
            # Verificar perfiles necesarios
            codigo_perfil_marco = config["perfil_marco"]
            perfil_marco = perfiles.get(codigo_perfil_marco)
            if not perfil_marco:
                raise ValueError(f"Perfil de marco '{codigo_perfil_marco}' no encontrado")
            
            # Verificar hoja si corresponde
            if config["perfil_hoja"]:
                codigo_perfil_hoja = config["perfil_hoja"]
                perfil_hoja = perfiles.get(codigo_perfil_hoja)
                if not perfil_hoja:
                    raise ValueError(f"Perfil de hoja '{codigo_perfil_hoja}' no encontrado")
            else:
                codigo_perfil_hoja = None
                perfil_hoja = None
            
            # Verificar travesaño si corresponde
            if config["perfil_travesano"]:
                codigo_perfil_travesano = config["perfil_travesano"]
                perfil_travesano = perfiles.get(codigo_perfil_travesano)
                if not perfil_travesano:
                    raise ValueError(f"Perfil de travesaño '{codigo_perfil_travesano}' no encontrado")
            else:
                codigo_perfil_travesano = None
                perfil_travesano = None
            
            # Crear el trabajo
            trabajo = (
                id_trabajo,
                cliente,
                nombre,
                codigo_perfil_marco,
                perfil_marco["descripcion"],
                int(perfil_marco["long_barra"]),
                cortes_marco.copy(),
                cortes_hoja.copy(),
                cortes_travesano.copy() if config["perfil_travesano"] else [],
                aplicacion,
                perfil_marco["descripcion"],
                perfil_hoja["descripcion"] if perfil_hoja else "",
                perfil_travesano["descripcion"] if perfil_travesano else ""
            )
            
            # Agregar o actualizar el trabajo
            if es_edicion:
                for i, t in enumerate(trabajos):
                    if t[0] == id_trabajo:
                        trabajos[i] = trabajo
                        break
            else:
                trabajos.append(trabajo)
                ID += 1
            
            actualizar_lista_trabajos()
            actualizar_color_boton_calcular()
            ventana_nueva.destroy()
            
        except ValueError as e:
            messagebox.showerror("Error al guardar", str(e))
        except Exception as e:
            messagebox.showerror("Error inesperado", f"Error al guardar el trabajo: {str(e)}")

    def actualizar_configuracion_angulos(event=None):
        """Actualiza la configuración de ángulos según la aplicación seleccionada"""
        aplicacion = combo_aplicacion.get()
        if not aplicacion:
            return
            
        config = CONFIG_APLICACIONES[aplicacion]
        
        # Aplicar configuraciones de ángulos
        if "angulos" in config:
            angulos = config["angulos"]
            
            # Configurar ángulos para marco
            if "marco" in angulos:
                marco_conf = angulos["marco"]
                # Configurar checkboxes
                chk_superior_var.set("superior" in marco_conf["aplicar"])
                chk_inferior_var.set("inferior" in marco_conf["aplicar"])
                chk_izquierda_var.set("izquierda" in marco_conf["aplicar"])
                chk_derecha_var.set("derecha" in marco_conf["aplicar"])
                
                # Configurar ángulos horizontales (ancho)
                combo_angulo_sup_izq.set(str(marco_conf["ancho"][0]))
                combo_angulo_sup_der.set(str(marco_conf["ancho"][1]))
                combo_angulo_inf_izq.set(str(marco_conf["ancho"][0]))
                combo_angulo_inf_der.set(str(marco_conf["ancho"][1]))
                
                # Configurar ángulos verticales (largo)
                combo_angulo_izq_sup.set(str(marco_conf["largo"][0]))
                combo_angulo_izq_inf.set(str(marco_conf["largo"][1]))
                combo_angulo_der_sup.set(str(marco_conf["largo"][0]))
                combo_angulo_der_inf.set(str(marco_conf["largo"][1]))
        
        # Actualizar visualización
        dibujar_aplicacion_dinamica()
    # Crear ventana para agregar/editar trabajo
    ventana_nueva = tk.Toplevel(ventana)
    ventana_nueva.title("Editar Trabajo" if es_edicion else "Agregar Trabajo")
    ventana_nueva.transient(ventana)
    ventana_nueva.grab_set()

    # Variables para checkboxes y ángulos
    aplicar_superior = False
    aplicar_inferior = False
    aplicar_izquierda = False
    aplicar_derecha = False
    angulo_sup_izq = 90
    angulo_sup_der = 90
    angulo_inf_izq = 90
    angulo_inf_der = 90
    angulo_izq_sup = 90
    angulo_izq_inf = 90
    angulo_der_sup = 90
    angulo_der_inf = 90

    # Campos principales
    tk.Label(ventana_nueva, text="Cliente:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    entry_cliente = tk.Entry(ventana_nueva, width=30)
    entry_cliente.grid(row=0, column=1, padx=10, pady=5, sticky="w")
    entry_cliente.insert(0, cliente_actual)

    tk.Label(ventana_nueva, text="Nombre del trabajo:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    entry_nombre = tk.Entry(ventana_nueva, width=30)
    entry_nombre.grid(row=1, column=1, padx=10, pady=5, sticky="w")
    entry_nombre.insert(0, nombre_actual)

    # Selector de aplicación
    tk.Label(ventana_nueva, text="Aplicación:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    combo_aplicacion = ttk.Combobox(ventana_nueva, values=list(CONFIG_APLICACIONES.keys()), state="readonly")
    combo_aplicacion.grid(row=2, column=1, padx=10, pady=5, sticky="w")
    if es_edicion:
        combo_aplicacion.set(aplicacion_actual)

    # Marco para mostrar la imagen y controles de medidas
    frame_imagen = ttk.LabelFrame(ventana_nueva, text="Vista previa y medidas")
    frame_imagen.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

    # Controles para medidas horizontales (arriba de la imagen)
    frame_horizontal = ttk.Frame(frame_imagen)
    frame_horizontal.grid(row=0, column=1, sticky="ew", pady=5)

    tk.Label(frame_horizontal, text="Medida Horizontal:").pack(side=tk.LEFT)
    entry_medida_horizontal = tk.Entry(frame_horizontal, width=8)
    entry_medida_horizontal.pack(side=tk.LEFT, padx=2)

    chk_superior_var = tk.BooleanVar(value=True)
    chk_superior = tk.Checkbutton(frame_horizontal, text="Superior", variable=chk_superior_var)
    chk_superior.pack(side=tk.LEFT, padx=2)

    chk_inferior_var = tk.BooleanVar(value=False)
    chk_inferior = tk.Checkbutton(frame_horizontal, text="Inferior", variable=chk_inferior_var)
    chk_inferior.pack(side=tk.LEFT, padx=2)

    tk.Label(frame_horizontal, text="Ángulo Izq:").pack(side=tk.LEFT)
    combo_angulo_sup_izq = ttk.Combobox(frame_horizontal, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_sup_izq.pack(side=tk.LEFT, padx=2)
    combo_angulo_sup_izq.set("90")
    
    tk.Label(frame_horizontal, text="Ángulo Der:").pack(side=tk.LEFT)
    combo_angulo_sup_der = ttk.Combobox(frame_horizontal, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_sup_der.pack(side=tk.LEFT, padx=2)
    combo_angulo_sup_der.set("90")

    # Controles para medidas verticales (izquierda de la imagen)
    frame_vertical = ttk.Frame(frame_imagen)
    frame_vertical.grid(row=1, column=0, sticky="ns", padx=5)

    tk.Label(frame_vertical, text="Medida Vertical:").pack()
    entry_medida_vertical = tk.Entry(frame_vertical, width=8)
    entry_medida_vertical.pack(pady=2)

    chk_izquierda_var = tk.BooleanVar(value=True)
    chk_izquierda = tk.Checkbutton(frame_vertical, text="Izquierda", variable=chk_izquierda_var)
    chk_izquierda.pack()

    chk_derecha_var = tk.BooleanVar(value=False)
    chk_derecha = tk.Checkbutton(frame_vertical, text="Derecha", variable=chk_derecha_var)
    chk_derecha.pack()

    tk.Label(frame_vertical, text="Ángulo Sup:").pack()
    combo_angulo_izq_sup = ttk.Combobox(frame_vertical, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_izq_sup.pack(pady=2)
    combo_angulo_izq_sup.set("90")
    
    tk.Label(frame_vertical, text="Ángulo Inf:").pack()
    combo_angulo_izq_inf = ttk.Combobox(frame_vertical, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_izq_inf.pack(pady=2)
    combo_angulo_izq_inf.set("90")

    # Ángulos para el lado inferior (debajo de la imagen)
    frame_angulos_inferior = ttk.Frame(frame_imagen)
    frame_angulos_inferior.grid(row=2, column=1, sticky="ew", pady=5)

    tk.Label(frame_angulos_inferior, text="Ángulo Izq:").pack(side=tk.LEFT)
    combo_angulo_inf_izq = ttk.Combobox(frame_angulos_inferior, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_inf_izq.pack(side=tk.LEFT, padx=2)
    combo_angulo_inf_izq.set("90")
    
    tk.Label(frame_angulos_inferior, text="Ángulo Der:").pack(side=tk.LEFT)
    combo_angulo_inf_der = ttk.Combobox(frame_angulos_inferior, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_inf_der.pack(side=tk.LEFT, padx=2)
    combo_angulo_inf_der.set("90")

    # Ángulos para el lado derecho (derecha de la imagen)
    frame_angulos_derecha = ttk.Frame(frame_imagen)
    frame_angulos_derecha.grid(row=1, column=2, sticky="ns", padx=5)

    tk.Label(frame_angulos_derecha, text="Ángulo Sup:").pack()
    combo_angulo_der_sup = ttk.Combobox(frame_angulos_derecha, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_der_sup.pack(pady=2)
    combo_angulo_der_sup.set("90")
    
    tk.Label(frame_angulos_derecha, text="Ángulo Inf:").pack()
    combo_angulo_der_inf = ttk.Combobox(frame_angulos_derecha, values=["45", "90", "135"], width=3, state="readonly")
    combo_angulo_der_inf.pack(pady=2)
    combo_angulo_der_inf.set("90")

    # Sección de cortes agregados
    frame_cortes = ttk.LabelFrame(ventana_nueva, text="Cortes Agregados")
    frame_cortes.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

    listbox_cortes = tk.Listbox(frame_cortes, width=50, height=5)
    listbox_cortes.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # Botones para agregar/eliminar cortes
    frame_botones = ttk.Frame(frame_cortes)
    frame_botones.pack(fill=tk.X, pady=5)

    btn_agregar_corte = tk.Button(frame_botones, text="+ Agregar Cortes", command=agregar_corte)
    btn_agregar_corte.pack(side=tk.LEFT, padx=5)

    btn_eliminar_corte = tk.Button(frame_botones, text="- Eliminar Corte", command=eliminar_corte)
    btn_eliminar_corte.pack(side=tk.LEFT, padx=5)

    # Crear canvas para visualización dinámica
    frame_visualizacion = ttk.Frame(ventana_nueva)
    frame_visualizacion.grid(row=3, column=2, rowspan=3, padx=10, pady=5, sticky="nsew")
    
    canvas_visualizacion = Canvas(
        frame_visualizacion,
        bg="white",
        width=500,
        height=500
    )
    canvas_visualizacion.pack(fill=tk.BOTH, expand=True)
    
    # Dibujar cuadrado inicial 500x500
    canvas_visualizacion.create_rectangle(
        50, 50, 300, 300,
        outline="black", width=2, tags="marco"
    )

    # Si es edición, cargar cortes existentes y actualizar visualización
    if es_edicion:
        for corte in cortes_marco:
            listbox_cortes.insert(tk.END, f"{corte[0]} mm - {corte[1]} (Marco) - {corte[2]}°/{corte[3]}°")
        for corte in cortes_hoja:
            listbox_cortes.insert(tk.END, f"{corte[0]} mm - {corte[1]} (Hoja) - {corte[2]}°/{corte[3]}°")
        for corte in cortes_travesano:
            listbox_cortes.insert(tk.END, f"{corte[0]} mm - {corte[1]} (Travesaño) - {corte[2]}°/{corte[3]}°")
        
        dibujar_aplicacion_dinamica()

    # Botón final para guardar el trabajo
    tk.Button(ventana_nueva, text="Guardar Trabajo", command=guardar_trabajo).grid(
        row=5, column=0, columnspan=2, pady=10, sticky="ew"
    )

    # Vincular eventos para actualización dinámica
    combo_aplicacion.bind("<<ComboboxSelected>>", actualizar_configuracion_angulos)
    entry_medida_horizontal.bind("<KeyRelease>", lambda e: dibujar_aplicacion_dinamica())
    entry_medida_vertical.bind("<KeyRelease>", lambda e: dibujar_aplicacion_dinamica())

    # Si es edición, actualizar imagen y configuración
    if es_edicion and aplicacion_actual:
        actualizar_configuracion_angulos()

# [Resto de las funciones del programa (actualizar_lista_trabajos, optimizar_cortes, etc.)...]

# ... (el resto del código permanece igual)
def actualizar_lista_trabajos():
    lista_trabajos.delete(*lista_trabajos.get_children())
    for trabajo in trabajos:
        # Contar cortes por tipo
        num_cortes_marco = len(trabajo[6])
        num_cortes_hoja = len(trabajo[7])
        num_cortes_travesano = len(trabajo[8]) if len(trabajo) > 8 else 0
        
        # Insertar en la lista
        lista_trabajos.insert("", "end", values=(
            trabajo[0],  # ID
            trabajo[1],  # Cliente
            trabajo[2],  # Nombre
       #     trabajo[4],  # Descripción perfil marco
            f"{num_cortes_marco} marco, {num_cortes_hoja} hoja, {num_cortes_travesano} travesaño",
            trabajo[5],  # Longitud barra
            trabajo[9] if len(trabajo) > 9 else "",  # Aplicación
            trabajo[10] if len(trabajo) > 10 else "",  # Marco
            trabajo[11] if len(trabajo) > 11 else "",  # Hoja
            trabajo[12] if len(trabajo) > 12 else ""   # Travesaño
        ))

def optimizar_cortes():
    if not trabajos:
        messagebox.showerror("Error", "No hay trabajos en la lista.")
        return

    global proveedor_global
    if not proveedor_global:
        messagebox.showerror("Error", "El nombre del proveedor no ha sido ingresado.")
        return

    # Organizar cortes por tipo de perfil
    cortes_por_perfil = {}
    
    for trabajo in trabajos:
        id_trabajo = trabajo[0]
        codigo_marco = trabajo[3]
        codigo_hoja = CONFIG_APLICACIONES.get(trabajo[9], {}).get("perfil_hoja") if len(trabajo) > 9 else None
        codigo_travesano = CONFIG_APLICACIONES.get(trabajo[9], {}).get("perfil_travesano") if len(trabajo) > 9 else None
        
        # Procesar cortes de marco
        if codigo_marco not in cortes_por_perfil:
            cortes_por_perfil[codigo_marco] = {
                "long_barra": int(trabajo[5]),
                "cortes": [],
                "descripcion": trabajo[4]
            }
        
        for corte in trabajo[6]:
            cortes_por_perfil[codigo_marco]["cortes"].append((corte, id_trabajo, "MARCO"))
        
        # Procesar cortes de hoja si existen
        if codigo_hoja:
            if codigo_hoja not in cortes_por_perfil:
                cortes_por_perfil[codigo_hoja] = {
                    "long_barra": int(perfiles[codigo_hoja]["long_barra"]),
                    "cortes": [],
                    "descripcion": perfiles[codigo_hoja]["descripcion"]
                }
            
            for corte in trabajo[7]:
                cortes_por_perfil[codigo_hoja]["cortes"].append((corte, id_trabajo, "HOJA"))
        
        # Procesar cortes de travesaño si existen
        if codigo_travesano:
            if codigo_travesano not in cortes_por_perfil:
                cortes_por_perfil[codigo_travesano] = {
                    "long_barra": int(perfiles[codigo_travesano]["long_barra"]),
                    "cortes": [],
                    "descripcion": perfiles[codigo_travesano]["descripcion"]
                }
            
            for corte in trabajo[8]:
                cortes_por_perfil[codigo_travesano]["cortes"].append((corte, id_trabajo, "TRAVESAÑO"))

    # Optimizar cada tipo de perfil y contar barras globalmente
    resultados_optimizacion = {}
    contador_barras_global = 1  # Iniciar conteo global de barras
    
    for codigo_perfil, datos in cortes_por_perfil.items():
        # Ordenar cortes de mayor a menor
        cortes_ordenados = sorted(datos["cortes"], key=lambda x: x[0][0], reverse=True)
        barras = []
        long_barra = datos["long_barra"]
        
        for corte_data in cortes_ordenados:
            corte, id_trabajo, tipo = corte_data
            longitud, orientacion, angulo_izq, angulo_der = corte
            added = False
            
            # Aplicar kerf según posición y ángulos
            if not barras or len(barras[-1]["cortes"]) > 0:  # No es el primer corte de la primera barra
                kerf = 8  # Kerf completo
            else:  # Es el primer corte de la barra
                kerf = 4 if angulo_izq == 90 else 8  # Kerf reducido si ángulo_izq = 90
            
            longitud_con_kerf = longitud + kerf
            
            # Intentar colocar en barras existentes
            for barra in barras:
                espacio_ocupado = sum(c[0][0] + c[1] for c in barra["cortes"])
                if espacio_ocupado + longitud_con_kerf <= long_barra:
                    barra["cortes"].append((corte, kerf, id_trabajo, tipo))
                    added = True
                    break
            
            # Si no cabe, crear nueva barra
            if not added:
                barras.append({
                    "cortes": [(corte, kerf, id_trabajo, tipo)],
                    "sobrante": long_barra - longitud_con_kerf,
                    "numero_barra_global": contador_barras_global  # Asignar número global
                })
                contador_barras_global += 1  # Incrementar contador global
        
        resultados_optimizacion[codigo_perfil] = {
            "descripcion": datos["descripcion"],
            "long_barra": long_barra,
            "barras": barras
        }

    # Generar reporte CSV con numeración global
    generar_reporte_csv(resultados_optimizacion)
    
def generar_reporte_csv(resultados_optimizacion):
    tabla_detallada = []
    id_general = 1
    
    for codigo_perfil, datos in resultados_optimizacion.items():
        perfil_data = perfiles[codigo_perfil]
        
        for barra in datos["barras"]:
            for num_corte, (corte_data, kerf, id_trabajo, tipo) in enumerate(barra["cortes"], 1):
                corte, orientacion, angulo_izq, angulo_der = corte_data
                
                # Obtener datos del trabajo
                trabajo = next((t for t in trabajos if t[0] == id_trabajo), None)
                if not trabajo:
                    continue
                
                # Calcular fecha/hora en formato Excel
                fecha_hora_excel = (datetime.now() - datetime(1899, 12, 30)).total_seconds() / 86400
                
                # Ajustar longitud según kerf (restar 4mm si es primer corte y ángulo izq = 90°)
                long_corte_final = corte + (0 if num_corte == 1 and angulo_izq == 90 else 4)
                
                tabla_detallada.append([
                    id_general,
                    barra["numero_barra_global"],  # Usar número global de barra
                    datos["long_barra"],
                    num_corte,
                    long_corte_final,
                    angulo_izq,
                    angulo_der,
                    codigo_perfil,
                    datos["descripcion"],
                    perfil_data["ancho"],
                    perfil_data["alto"],
                    1,  # carril
                    1,  # almacen
                    orientacion,
                    "870651",  # codigo_ds
                    perfil_data["refuerzo"],
                    id_general,  # no_pedido
                    proveedor_global,
                    fecha_hora_excel,
                    trabajo[1],  # cliente
                    "1",  # color_goma
                    "BLANCO",  # color
                    trabajo[2],  # descripcion2 (nombre trabajo)
                    None,  # ruta_imagen
                    1,  # stock
                    proveedor_global,  # cliente_vertical
                    id_trabajo,
                    trabajo[9] if len(trabajo) > 9 else "",  # aplicacion
                    trabajo[10] if len(trabajo) > 10 else "",  # marco
                    trabajo[11] if len(trabajo) > 11 else "",  # hoja
                    trabajo[12] if len(trabajo) > 12 else "",  # travesaño
                    tipo  # tipo de perfil (MARCO/HOJA/TRAVESAÑO)
                ])
                id_general += 1

    archivo_csv = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv")],
        title="Optimización_detallada"
    )

    if archivo_csv:
        with open(archivo_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow([
                "id", "num_barra", "long_barra", "num_corte",
                "long_corte", "angulo_inicial_corte", "angulo_final_corte",
                "codigo_perfil", "descripcion", "ancho_perfil", "altura_perfil",
                "carril", "almacen", "orientacion", "codigo_ds", "Poz", "no_pedido",
                "proveedor", "fecha_hora", "cliente", "color_goma", "color",
                "descripcion2", "ruta_imagen", "stock", "cliente_vertical", "id_trabajo",
                "aplicacion", "marco", "hoja", "travesaño", "tipo_perfil"
            ])
            writer.writerows(tabla_detallada)

        messagebox.showinfo("Éxito", "Reporte generado correctamente")
        mostrar_resultados_desde_csv(archivo_csv)
        
def mostrar_resultados_desde_csv(archivo_csv):
    try:
        with open(archivo_csv, newline="", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            next(reader)  # Saltar cabecera
            resultados = list(reader)
    except Exception as e:
        messagebox.showerror("Error", f"Error al leer CSV: {str(e)}")
        return

    ventana_resultados = tk.Toplevel()
    ventana_resultados.title("Cortes Optimizados")

    canvas = Canvas(ventana_resultados, width=800, height=600, bg="white")
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(ventana_resultados, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.configure(yscrollcommand=scrollbar.set)
    frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    y_position = 20
    barra_height = 20
    color_corte_marco = "#4CAF50"  # Verde
    color_corte_hoja = "#2196F3"    # Azul
    color_corte_travesano = "#9C27B0"  # Morado
    color_sobrante = "#FF5252"      # Rojo

    # Organizar resultados por número global de barra
    barras_globales = {}
    for row in resultados:
        if len(row) < 27:
            continue

        num_barra_global = int(row[1])
        codigo_perfil = row[7]
        descripcion = row[8]
        tipo_perfil = f"{codigo_perfil} - {descripcion}"
        id_trabajo = int(row[26])
        tipo_corte = row[31] if len(row) > 31 else "MARCO"
        long_corte = int(row[4])
        orientacion = row[13]
        angulo_izq = int(row[5])
        angulo_der = int(row[6])
        long_barra = int(row[2])
        
        if num_barra_global not in barras_globales:
            barras_globales[num_barra_global] = {
                "tipo_perfil": tipo_perfil,
                "long_barra": long_barra,
                "cortes": [],
                "sobrante": long_barra
            }
        
        barras_globales[num_barra_global]["cortes"].append((
            long_corte, id_trabajo, orientacion, angulo_izq, angulo_der, tipo_corte
        ))
        barras_globales[num_barra_global]["sobrante"] -= long_corte

    # Mostrar barras ordenadas por número global
    for num_barra_global in sorted(barras_globales.keys()):
        barra_data = barras_globales[num_barra_global]
        tipo_perfil = barra_data["tipo_perfil"]
        longitud_barra = barra_data["long_barra"]
        cortes = barra_data["cortes"]
        sobrante = barra_data["sobrante"]

        # Mostrar información de la barra
        canvas.create_text(50, y_position, 
                         text=f"Barra {num_barra_global} | Perfil: {tipo_perfil} | Longitud: {longitud_barra} mm", 
                         font=("Arial", 11, "bold"), anchor="w", fill="#333")
        y_position += 30

        canvas.create_rectangle(50, y_position, 750, y_position + barra_height, outline="#333", width=2, fill="#F0F0F0")

        x_position = 50
        for corte in cortes:
            long_corte, id_trabajo, orientacion, angulo_izq, angulo_der, tipo_corte = corte
            
            # Determinar color según tipo de corte
            if tipo_corte == "MARCO":
                color = color_corte_marco
            elif tipo_corte == "HOJA":
                color = color_corte_hoja
            else:  # TRAVESAÑO
                color = color_corte_travesano

            ancho = (long_corte / longitud_barra) * 700
            canvas.create_rectangle(x_position, y_position, x_position + ancho, y_position + barra_height, fill=color, outline="#333")

            # Dibujar ángulos
            angle_length = 20
            if angulo_izq == 45 and angulo_der == 135:
                canvas.create_line(x_position, y_position, x_position + angle_length, y_position + barra_height, fill="black", width=2)
                canvas.create_line(x_position + ancho, y_position, x_position + ancho - angle_length, y_position + barra_height, fill="black", width=2)
            elif angulo_izq == 135 and angulo_der == 45:
                canvas.create_line(x_position, y_position + barra_height, x_position + angle_length, y_position, fill="black", width=2)
                canvas.create_line(x_position + ancho, y_position + barra_height, x_position + ancho - angle_length, y_position, fill="black", width=2)
            else:
                if angulo_izq == 45:
                    canvas.create_line(x_position, y_position, x_position + angle_length, y_position + barra_height, fill="black", width=2)
                elif angulo_izq == 135:
                    canvas.create_line(x_position, y_position + barra_height, x_position + angle_length, y_position, fill="black", width=2)

                if angulo_der == 45:
                    canvas.create_line(x_position + ancho, y_position, x_position + ancho - angle_length, y_position + barra_height, fill="black", width=2)
                elif angulo_der == 135:
                    canvas.create_line(x_position + ancho, y_position + barra_height, x_position + ancho - angle_length, y_position, fill="black", width=2)

            # Etiqueta del corte
            trabajo = next((t for t in trabajos if t[0] == id_trabajo), None)
            nombre_trabajo = trabajo[2] if trabajo else f"ID:{id_trabajo}"
            
            canvas.create_text(x_position + (ancho / 2), y_position + 45, 
                             text=f"{long_corte} mm\n({tipo_corte}, {nombre_trabajo})\n{angulo_izq}°/{angulo_der}°", 
                             fill="black", font=("Arial", 8, "bold","italic"), anchor="s")

            x_position += ancho

        # Dibujar sobrante si existe
        if sobrante > 0:
            ancho_sobrante = (sobrante / longitud_barra) * 700
            canvas.create_rectangle(x_position, y_position, x_position + ancho_sobrante, y_position + barra_height, fill=color_sobrante, outline="#333")
            canvas.create_text(x_position + (ancho_sobrante / 2), y_position + (barra_height / 2) + 20, 
                             text=f"S: {sobrante} mm", fill="black", font=("Arial", 8, "bold"))

        y_position += barra_height + 50

    frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))
    btn_imprimir = tk.Button(ventana_resultados, text="🖨️ Imprimir Etiquetas", command=lambda: generar_pdf_etiquetas(barras_globales))
    btn_imprimir.pack(side=tk.TOP, pady=10)

    btn_ver_csv = tk.Button(ventana_resultados, text="📄 Ver CSV", command=lambda: visualizar_csv(archivo_csv))
    btn_ver_csv.pack(side=tk.TOP, pady=10)

def visualizar_csv(archivo_csv):
    try:
        with open(archivo_csv, newline="", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            contenido = list(reader)
    except Exception as e:
        messagebox.showerror("Error", f"Error al leer CSV: {str(e)}")
        return

    ventana_csv = tk.Toplevel()
    ventana_csv.title("Contenido del CSV")

    tree = ttk.Treeview(ventana_csv, columns=contenido[0], show="headings")
    for col in contenido[0]:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for fila in contenido[1:]:
        tree.insert("", "end", values=fila)

    scrollbar = ttk.Scrollbar(ventana_csv, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

def limpiar_tabla():
    global ID, trabajos
    trabajos = []
    ID = 1
    actualizar_lista_trabajos()
    actualizar_color_boton_calcular()
    messagebox.showinfo("Éxito", "Tabla de trabajos limpiada correctamente.")

def eliminar_trabajo():
    seleccionado = lista_trabajos.selection()
    if seleccionado:
        item = lista_trabajos.item(seleccionado)
        id_trabajo = item["values"][0]

        global trabajos
        trabajos = [t for t in trabajos if t[0] != id_trabajo]

        actualizar_lista_trabajos()
        actualizar_color_boton_calcular()
        messagebox.showinfo("Éxito", "Trabajo eliminado correctamente.")
    else:
        messagebox.showwarning("Advertencia", "Seleccione un trabajo para eliminar.")

def actualizar_color_boton_calcular():
    if trabajos:
        btn_calcular.config(bg="green", fg="white")
    else:
        btn_calcular.config(bg="SystemButtonFace", fg="black")

def modificar_trabajo():
    seleccionado = lista_trabajos.selection()
    if not seleccionado:
        messagebox.showwarning("Advertencia", "Seleccione un trabajo para modificar")
        return
    
    item = lista_trabajos.item(seleccionado)
    id_trabajo = item["values"][0]

    for trabajo in trabajos:
        if trabajo[0] == id_trabajo:
            agregar_trabajo(job_para_editar=trabajo)
            break

def actualizar_botones_seleccion(event=None):
    seleccionado = lista_trabajos.selection()
    if seleccionado:
        btn_eliminar.config(state=tk.NORMAL, bg="red", fg="white")
        btn_modificar.config(state=tk.NORMAL)
    else:
        btn_eliminar.config(state=tk.DISABLED, bg="SystemButtonFace", fg="black")
        btn_modificar.config(state=tk.DISABLED)

def solicitar_proveedor():
    def guardar_proveedor():
        global proveedor_global
        proveedor_global = entry_proveedor.get().strip()
        if not proveedor_global:
            messagebox.showerror("Error", "El nombre del proveedor es obligatorio.")
        else:
            ventana_proveedor.destroy()
            ventana.deiconify()

    ventana.withdraw()
    ventana_proveedor = tk.Toplevel()
    ventana_proveedor.title("Inicio Sesión")
    ventana_proveedor.geometry("300x120")

    tk.Label(ventana_proveedor, text="Nombre del Proveedor:").pack(pady=10)
    entry_proveedor = tk.Entry(ventana_proveedor, width=30)
    entry_proveedor.pack(pady=5)

    tk.Button(ventana_proveedor, text="Inicio", command=guardar_proveedor).pack(pady=10)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Optimización de Cortes de PVC")
ventana.geometry("1000x600")

# Configuración de la interfaz
frame_controles = ttk.Frame(ventana)
frame_controles.pack(pady=10, fill=tk.X)

btn_agregar = tk.Button(frame_controles, text="Agregar Trabajo", command=agregar_trabajo)
btn_agregar.grid(row=0, column=0, padx=5)

btn_calcular = tk.Button(frame_controles, text="Calcular", command=optimizar_cortes)
btn_calcular.grid(row=0, column=1, padx=5)

btn_limpiar = tk.Button(frame_controles, text="Limpiar Tabla", command=limpiar_tabla)
btn_limpiar.grid(row=0, column=2, padx=5)

btn_eliminar = tk.Button(frame_controles, text="Eliminar Trabajo", command=eliminar_trabajo, state=tk.DISABLED)
btn_eliminar.grid(row=0, column=3, padx=5)

btn_modificar = tk.Button(frame_controles, text="Modificar Trabajo", command=modificar_trabajo, state=tk.DISABLED)
btn_modificar.grid(row=0, column=4, padx=5)

# Lista de trabajos
columns = ("ID", "Cliente", "Trabajo", "Cortes", "Barra (mm)", "Aplicación", "Marco", "Hoja", "Travesaño")
lista_trabajos = ttk.Treeview(ventana, columns=columns, show="headings", height=15)

for col in columns:
    lista_trabajos.heading(col, text=col)
    lista_trabajos.column(col, width=100, anchor=tk.W)

scrollbar = ttk.Scrollbar(ventana, orient=tk.VERTICAL, command=lista_trabajos.yview)
lista_trabajos.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
lista_trabajos.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Vincular eventos
lista_trabajos.bind("<<TreeviewSelect>>", actualizar_botones_seleccion)

# Actualizar el color del botón calcular al inicio
actualizar_color_boton_calcular()

# Llamar a la función para solicitar el proveedor antes de iniciar la ventana principal
solicitar_proveedor()

# Iniciar la aplicación
ventana.mainloop()