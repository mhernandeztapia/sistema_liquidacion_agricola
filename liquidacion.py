"""
Sistema de Liquidación de Sueldos

"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime, date
from generador_pdf import generar_pdf_liquidacion
from generador_excel import generar_excel_liquidacion
from calculos import calcular_liquidacion

LABORES = {
    "Paletizador/a":          610_000,
    "Seleccionadora de Fruta": 570_000,
    "Armador/a de Caja":       590_000,
    "Embaladora de Caja":      560_000,
}

MESES = [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
]

COLORES = {
    "verde_oscuro": "#1B4332",
    "verde_medio":  "#2D6A4F",
    "verde_claro":  "#52B788",
    "fondo":        "#F8F9FA",
    "blanco":       "#FFFFFF",
    "gris_claro":   "#E9ECEF",
    "gris_texto":   "#6C757D",
    "negro":        "#212529",
    "rojo":         "#DC3545",
    "amarillo":     "#FFC107",
}

ARCHIVO_TRABAJADORES = os.path.join(os.path.dirname(__file__), "trabajadores.json")


# ─────────────────────────────────────────────
#  PERSISTENCIA DE TRABAJADORES
# ─────────────────────────────────────────────
def cargar_trabajadores():
    if os.path.exists(ARCHIVO_TRABAJADORES):
        with open(ARCHIVO_TRABAJADORES, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def guardar_trabajadores(datos):
    with open(ARCHIVO_TRABAJADORES, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────
#  VENTANA PRINCIPAL
# ─────────────────────────────────────────────
class AppLiquidacion(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Liquidación de Sueldos")
        self.geometry("1100x750")
        self.minsize(900, 650)
        self.configure(bg=COLORES["fondo"])
        self.resizable(True, True)

        self.trabajadores = cargar_trabajadores()
        self._construir_ui()
        self._actualizar_lista_trabajadores()

    # ----------- Layout principal 
    def _construir_ui(self):
        # Barra superior
        header = tk.Frame(self, bg=COLORES["verde_oscuro"], height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header, text="Liquidación de Sueldo",
            font=("Helvetica", 22, "bold"),
            fg=COLORES["blanco"], bg=COLORES["verde_oscuro"]
        ).pack(side="left", padx=20, pady=14)

        tk.Label(
            header, text="Sistema de Liquidación de Sueldos",
            font=("Helvetica", 11),
            fg=COLORES["verde_claro"], bg=COLORES["verde_oscuro"]
        ).pack(side="left", pady=14)

        # Notebook (pestañas)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook",            background=COLORES["fondo"])
        style.configure("TNotebook.Tab",        background=COLORES["gris_claro"],
                        foreground=COLORES["negro"], padding=[14, 6],
                        font=("Helvetica", 10, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", COLORES["verde_medio"])],
                  foreground=[("selected", COLORES["blanco"])])

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=10)

        self.tab_liquidar  = tk.Frame(self.notebook, bg=COLORES["fondo"])
        self.tab_trabajadores = tk.Frame(self.notebook, bg=COLORES["fondo"])

        self.notebook.add(self.tab_liquidar,     text="  📄  Nueva Liquidación  ")
        self.notebook.add(self.tab_trabajadores, text="  👷  Trabajadores  ")

        self._construir_tab_liquidar()
        self._construir_tab_trabajadores()

    def _construir_tab_liquidar(self):
        tab = self.tab_liquidar
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)

        #-----------Panel izquierdo: Datos del trabajador
        izq = self._tarjeta(tab, "👷  Datos del Trabajador")
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=4)

        campos_izq = [
            ("Nombre completo *",   "nombre"),
            ("RUT (ej: 12.345.678-9)", "rut"),
            ("Nombre empresa/faena", "empresa"),
        ]
        self.vars_liq = {}
        for i, (label, key) in enumerate(campos_izq):
            self._campo_texto(izq, label, key, i)

        #------------------Labor
        tk.Label(izq, text="Labor / Cargo *",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"], bg=COLORES["blanco"]).grid(
            row=3, column=0, sticky="w", padx=10, pady=(8, 2))

        self.vars_liq["labor"] = tk.StringVar()
        cb_labor = ttk.Combobox(izq, textvariable=self.vars_liq["labor"],
                                values=list(LABORES.keys()), state="readonly",
                                font=("Helvetica", 10), width=30)
        cb_labor.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 4))
        cb_labor.bind("<<ComboboxSelected>>", self._al_cambiar_labor)

        tk.Label(izq, text="Sueldo base ($)",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"], bg=COLORES["blanco"]).grid(
            row=5, column=0, sticky="w", padx=10, pady=(4, 2))

        self.vars_liq["sueldo_base"] = tk.StringVar(value="0")
        entrada_sueldo = tk.Entry(izq, textvariable=self.vars_liq["sueldo_base"],
                                  font=("Helvetica", 10), width=32,
                                  relief="solid", bd=1)
        entrada_sueldo.grid(row=6, column=0, sticky="ew", padx=10, pady=(0, 4))

        #------------------Período
        tk.Label(izq, text="Período de liquidación",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"], bg=COLORES["blanco"]).grid(
            row=7, column=0, sticky="w", padx=10, pady=(8, 2))

        periodo_frame = tk.Frame(izq, bg=COLORES["blanco"])
        periodo_frame.grid(row=8, column=0, sticky="ew", padx=10, pady=(0, 4))

        hoy = date.today()
        self.vars_liq["mes"] = tk.StringVar(value=MESES[hoy.month - 1])
        self.vars_liq["anio"] = tk.StringVar(value=str(hoy.year))

        cb_mes = ttk.Combobox(periodo_frame, textvariable=self.vars_liq["mes"],
                              values=MESES, state="readonly",
                              font=("Helvetica", 10), width=14)
        cb_mes.pack(side="left", padx=(0, 6))

        ttk.Combobox(periodo_frame, textvariable=self.vars_liq["anio"],
                     values=[str(y) for y in range(2023, 2031)],
                     state="readonly", font=("Helvetica", 10), width=8).pack(side="left")

        # ── Panel derecho: Haberes y descuentos ──
        der = self._tarjeta(tab, "💰  Haberes y Descuentos")
        der.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=4)
        der.columnconfigure(0, weight=1)
        der.columnconfigure(1, weight=1)

        # Días trabajados
        tk.Label(der, text="Días trabajados *",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"], bg=COLORES["blanco"]).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 2))
        self.vars_liq["dias_trabajados"] = tk.StringVar(value="30")
        tk.Entry(der, textvariable=self.vars_liq["dias_trabajados"],
                 font=("Helvetica", 10), width=10,
                 relief="solid", bd=1).grid(row=1, column=0, sticky="w", padx=10)

        # Horas extra
        tk.Label(der, text="Horas extras (+ $5.000 c/u)",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"], bg=COLORES["blanco"]).grid(
            row=0, column=1, sticky="w", padx=10, pady=(8, 2))
        self.vars_liq["horas_extra"] = tk.StringVar(value="0")
        tk.Entry(der, textvariable=self.vars_liq["horas_extra"],
                 font=("Helvetica", 10), width=10,
                 relief="solid", bd=1).grid(row=1, column=1, sticky="w", padx=10)

        # Separador
        ttk.Separator(der, orient="horizontal").grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

        # Checkboxes de haberes adicionales
        self.vars_liq["bono_colacion"]     = tk.BooleanVar(value=True)
        self.vars_liq["bono_movilizacion"] = tk.BooleanVar(value=True)
        self.vars_liq["bono_asistencia"]   = tk.BooleanVar(value=False)
        self.vars_liq["asig_familiar"]     = tk.BooleanVar(value=False)

        haberes = [
            ("bono_colacion",     "Bono Colación ($47.318)"),
            ("bono_movilizacion", "Bono Movilización ($30.500)"),
            ("bono_asistencia",   "Bono Asistencia Perfecta ($25.000)"),
            ("asig_familiar",     "Asignación Familiar"),
        ]
        tk.Label(der, text="Haberes adicionales",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"], bg=COLORES["blanco"]).grid(
            row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 4))

        for j, (key, texto) in enumerate(haberes):
            tk.Checkbutton(der, text=texto, variable=self.vars_liq[key],
                           font=("Helvetica", 10), bg=COLORES["blanco"],
                           activebackground=COLORES["blanco"],
                           fg=COLORES["negro"]).grid(
                row=4 + j, column=0, columnspan=2, sticky="w", padx=20)

        # Asignación familiar: cargas
        self.vars_liq["cargas_familiares"] = tk.StringVar(value="0")
        f_cargas = tk.Frame(der, bg=COLORES["blanco"])
        f_cargas.grid(row=8, column=0, columnspan=2, sticky="w", padx=30)
        tk.Label(f_cargas, text="N° de cargas:", font=("Helvetica", 9),
                 bg=COLORES["blanco"]).pack(side="left")
        tk.Entry(f_cargas, textvariable=self.vars_liq["cargas_familiares"],
                 width=5, font=("Helvetica", 9), relief="solid", bd=1).pack(side="left", padx=4)

        ttk.Separator(der, orient="horizontal").grid(
            row=9, column=0, columnspan=2, sticky="ew", pady=10, padx=10)

        # Descuentos manuales
        tk.Label(der, text="Descuentos adicionales",
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["rojo"], bg=COLORES["blanco"]).grid(
            row=10, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 4))

        descuentos = [
            ("desc_prestamo",  "Préstamo empresa ($)"),
            ("desc_adelanto",  "Adelanto de sueldo ($)"),
            ("desc_otros",     "Otros descuentos ($)"),
        ]
        for j, (key, texto) in enumerate(descuentos):
            tk.Label(der, text=texto, font=("Helvetica", 9),
                     bg=COLORES["blanco"], fg=COLORES["gris_texto"]).grid(
                row=11 + j, column=0, sticky="w", padx=20, pady=2)
            self.vars_liq[key] = tk.StringVar(value="0")
            tk.Entry(der, textvariable=self.vars_liq[key],
                     font=("Helvetica", 10), width=14,
                     relief="solid", bd=1).grid(row=11 + j, column=1, sticky="w", padx=10)

        # ── Botones de acción ──────────────────
        btn_frame = tk.Frame(tab, bg=COLORES["fondo"])
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)

        self._boton(btn_frame, "🔍  Vista Previa",
                    COLORES["verde_medio"], self._vista_previa).pack(side="left", padx=6)
        self._boton(btn_frame, "📄  Exportar PDF",
                    COLORES["verde_oscuro"], self._exportar_pdf).pack(side="left", padx=6)
        self._boton(btn_frame, "📊  Exportar Excel",
                    "#166534", self._exportar_excel).pack(side="left", padx=6)
        self._boton(btn_frame, "💾  Guardar Trabajador",
                    COLORES["amarillo"], self._guardar_trabajador,
                    fg=COLORES["negro"]).pack(side="left", padx=6)
        self._boton(btn_frame, "🗑  Limpiar",
                    COLORES["gris_texto"], self._limpiar_formulario).pack(side="left", padx=6)

    # ─────────────────────────────────────────
    #  PESTAÑA: TRABAJADORES GUARDADOS
    # ─────────────────────────────────────────
    def _construir_tab_trabajadores(self):
        tab = self.tab_trabajadores

        # Barra de búsqueda
        top = tk.Frame(tab, bg=COLORES["fondo"])
        top.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(top, text="Buscar:", font=("Helvetica", 10),
                 bg=COLORES["fondo"]).pack(side="left")
        self.busqueda_var = tk.StringVar()
        self.busqueda_var.trace("w", lambda *_: self._filtrar_lista())
        tk.Entry(top, textvariable=self.busqueda_var,
                 font=("Helvetica", 10), width=30,
                 relief="solid", bd=1).pack(side="left", padx=6)

        # Tabla
        cols = ("Nombre", "RUT", "Labor", "Sueldo Base")
        self.tree = ttk.Treeview(tab, columns=cols, show="headings", height=18)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=200, anchor="center")

        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica", 10), rowheight=28)
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))

        sb = ttk.Scrollbar(tab, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=4)
        sb.pack(side="left", fill="y", pady=4)

        self.tree.bind("<Double-1>", self._cargar_trabajador_seleccionado)

        btn2 = tk.Frame(tab, bg=COLORES["fondo"])
        btn2.pack(fill="x", padx=12, pady=6)
        self._boton(btn2, "📝  Cargar en Formulario",
                    COLORES["verde_medio"],
                    self._cargar_trabajador_seleccionado).pack(side="left", padx=4)
        self._boton(btn2, "🗑  Eliminar Trabajador",
                    COLORES["rojo"],
                    self._eliminar_trabajador).pack(side="left", padx=4)

    # ─────────────────────────────────────────
    #  HELPERS DE CONSTRUCCIÓN UI
    # ─────────────────────────────────────────
    def _tarjeta(self, parent, titulo):
        frame = tk.LabelFrame(parent, text=titulo,
                              font=("Helvetica", 11, "bold"),
                              fg=COLORES["verde_oscuro"],
                              bg=COLORES["blanco"],
                              relief="solid", bd=1,
                              labelanchor="nw", padx=4, pady=4)
        frame.columnconfigure(0, weight=1)
        return frame

    def _campo_texto(self, parent, label, key, row):
        tk.Label(parent, text=label,
                 font=("Helvetica", 10, "bold"),
                 fg=COLORES["verde_oscuro"],
                 bg=COLORES["blanco"]).grid(
            row=row * 2, column=0, sticky="w", padx=10, pady=(8, 2))
        self.vars_liq[key] = tk.StringVar()
        tk.Entry(parent, textvariable=self.vars_liq[key],
                 font=("Helvetica", 10), width=32,
                 relief="solid", bd=1).grid(
            row=row * 2 + 1, column=0, sticky="ew", padx=10, pady=(0, 4))

    def _boton(self, parent, texto, bg, comando, fg="#FFFFFF"):
        return tk.Button(parent, text=texto, font=("Helvetica", 10, "bold"),
                         bg=bg, fg=fg, activebackground=bg,
                         relief="flat", padx=14, pady=8,
                         cursor="hand2", command=comando)

    # ─────────────────────────────────────────
    #  LÓGICA DE INTERFAZ
    # ─────────────────────────────────────────
    def _al_cambiar_labor(self, _event=None):
        labor = self.vars_liq["labor"].get()
        if labor in LABORES:
            self.vars_liq["sueldo_base"].set(str(LABORES[labor]))

    def _recolectar_datos(self):
        """Recolecta y valida los campos del formulario."""
        datos = {}
        for k, v in self.vars_liq.items():
            datos[k] = v.get()

        # Validación básica
        if not datos["nombre"].strip():
            messagebox.showerror("Campo requerido", "Ingresa el nombre del trabajador.")
            return None
        if not datos["labor"]:
            messagebox.showerror("Campo requerido", "Selecciona una labor.")
            return None
        try:
            datos["sueldo_base"]       = int(datos["sueldo_base"])
            datos["dias_trabajados"]   = int(datos["dias_trabajados"])
            datos["horas_extra"]       = int(datos["horas_extra"])
            datos["cargas_familiares"] = int(datos["cargas_familiares"])
            datos["desc_prestamo"]     = int(datos["desc_prestamo"])
            datos["desc_adelanto"]     = int(datos["desc_adelanto"])
            datos["desc_otros"]        = int(datos["desc_otros"])
        except ValueError:
            messagebox.showerror("Error de formato",
                                 "Revisa que días, horas y montos sean números enteros.")
            return None
        if datos["dias_trabajados"] < 0 or datos["dias_trabajados"] > 31:
            messagebox.showerror("Valor inválido", "Los días trabajados deben ser entre 0 y 31.")
            return None
        return datos

    def _vista_previa(self):
        datos = self._recolectar_datos()
        if not datos:
            return
        resumen = calcular_liquidacion(datos)
        VentanaResumen(self, datos, resumen)

    def _exportar_pdf(self):
        datos = self._recolectar_datos()
        if not datos:
            return
        resumen = calcular_liquidacion(datos)
        ruta = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"Liquidacion_{datos['nombre'].replace(' ','_')}_{datos['mes']}_{datos['anio']}.pdf"
        )
        if ruta:
            generar_pdf_liquidacion(datos, resumen, ruta)
            messagebox.showinfo("✅ PDF generado",
                                f"Liquidación exportada correctamente:\n{ruta}")

    def _exportar_excel(self):
        datos = self._recolectar_datos()
        if not datos:
            return
        resumen = calcular_liquidacion(datos)
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"Liquidacion_{datos['nombre'].replace(' ','_')}_{datos['mes']}_{datos['anio']}.xlsx"
        )
        if ruta:
            generar_excel_liquidacion(datos, resumen, ruta)
            messagebox.showinfo("✅ Excel generado",
                                f"Planilla exportada correctamente:\n{ruta}")

    def _guardar_trabajador(self):
        datos = self._recolectar_datos()
        if not datos:
            return
        clave = datos["rut"].strip() or datos["nombre"].strip()
        self.trabajadores[clave] = {
            "nombre": datos["nombre"],
            "rut":    datos["rut"],
            "labor":  datos["labor"],
            "sueldo_base": datos["sueldo_base"],
            "empresa": datos["empresa"],
        }
        guardar_trabajadores(self.trabajadores)
        self._actualizar_lista_trabajadores()
        messagebox.showinfo("Guardado", f"Trabajador '{datos['nombre']}' guardado correctamente.")

    def _limpiar_formulario(self):
        for k, v in self.vars_liq.items():
            if isinstance(v, tk.BooleanVar):
                v.set(False)
            elif k in ("dias_trabajados",):
                v.set("30")
            elif k in ("horas_extra","cargas_familiares",
                        "desc_prestamo","desc_adelanto","desc_otros"):
                v.set("0")
            elif k == "sueldo_base":
                v.set("0")
            else:
                v.set("")
        self.vars_liq["bono_colacion"].set(True)
        self.vars_liq["bono_movilizacion"].set(True)

    def _actualizar_lista_trabajadores(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for t in self.trabajadores.values():
            self.tree.insert("", "end", values=(
                t.get("nombre",""),
                t.get("rut",""),
                t.get("labor",""),
                f"${t.get('sueldo_base',0):,}".replace(",",".")
            ))

    def _filtrar_lista(self):
        texto = self.busqueda_var.get().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for t in self.trabajadores.values():
            if texto in t.get("nombre","").lower() or texto in t.get("rut","").lower():
                self.tree.insert("", "end", values=(
                    t.get("nombre",""),
                    t.get("rut",""),
                    t.get("labor",""),
                    f"${t.get('sueldo_base',0):,}".replace(",",".")
                ))

    def _cargar_trabajador_seleccionado(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selección", "Selecciona un trabajador de la lista.")
            return
        vals = self.tree.item(sel[0])["values"]
        nombre = vals[0]
        for t in self.trabajadores.values():
            if t.get("nombre") == nombre:
                self.vars_liq["nombre"].set(t.get("nombre",""))
                self.vars_liq["rut"].set(t.get("rut",""))
                self.vars_liq["empresa"].set(t.get("empresa",""))
                self.vars_liq["labor"].set(t.get("labor",""))
                self.vars_liq["sueldo_base"].set(str(t.get("sueldo_base",0)))
                self.notebook.select(0)
                break

    def _eliminar_trabajador(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selección", "Selecciona un trabajador para eliminar.")
            return
        vals = self.tree.item(sel[0])["values"]
        nombre = vals[0]
        if messagebox.askyesno("Confirmar", f"¿Eliminar a '{nombre}' de la lista?"):
            claves = [k for k, v in self.trabajadores.items()
                      if v.get("nombre") == nombre]
            for c in claves:
                del self.trabajadores[c]
            guardar_trabajadores(self.trabajadores)
            self._actualizar_lista_trabajadores()


# ─────────────────────────────────────────────
#  VENTANA DE RESUMEN / VISTA PREVIA
# ─────────────────────────────────────────────
class VentanaResumen(tk.Toplevel):
    def __init__(self, parent, datos, resumen):
        super().__init__(parent)
        self.title("Vista Previa – Liquidación")
        self.geometry("520x640")
        self.configure(bg=COLORES["blanco"])
        self.resizable(False, False)
        self.grab_set()

        self._construir(datos, resumen)

    def _construir(self, d, r):
        # Encabezado
        hdr = tk.Frame(self, bg=COLORES["verde_oscuro"])
        hdr.pack(fill="x")
        tk.Label(hdr, text="LIQUIDACIÓN DE SUELDO",
                 font=("Helvetica", 16, "bold"),
                 fg=COLORES["blanco"], bg=COLORES["verde_oscuro"],
                 pady=14).pack()
        tk.Label(hdr, text=f"{d['mes']} {d['anio']}",
                 font=("Helvetica", 11),
                 fg=COLORES["verde_claro"], bg=COLORES["verde_oscuro"],
                 pady=6).pack()

        # Datos personales
        info = tk.Frame(self, bg=COLORES["gris_claro"])
        info.pack(fill="x", padx=20, pady=12)
        filas_info = [
            ("Trabajador", d["nombre"]),
            ("RUT",        d["rut"] or "—"),
            ("Empresa",    d["empresa"] or "—"),
            ("Labor",      d["labor"]),
            ("Días trabajados", str(d["dias_trabajados"])),
        ]
        for lbl, val in filas_info:
            f = tk.Frame(info, bg=COLORES["gris_claro"])
            f.pack(fill="x", padx=10, pady=2)
            tk.Label(f, text=lbl + ":", width=18, anchor="w",
                     font=("Helvetica", 9, "bold"),
                     bg=COLORES["gris_claro"]).pack(side="left")
            tk.Label(f, text=val, anchor="w",
                     font=("Helvetica", 9),
                     bg=COLORES["gris_claro"]).pack(side="left")

        def seccion(titulo, items, color_total=COLORES["negro"]):
            tk.Label(self, text=titulo,
                     font=("Helvetica", 10, "bold"),
                     fg=COLORES["verde_oscuro"], bg=COLORES["blanco"],
                     anchor="w").pack(fill="x", padx=20, pady=(8, 2))
            ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20)
            for nombre, monto in items:
                f = tk.Frame(self, bg=COLORES["blanco"])
                f.pack(fill="x", padx=30, pady=1)
                tk.Label(f, text=nombre, font=("Helvetica", 9),
                         bg=COLORES["blanco"], anchor="w").pack(side="left")
                tk.Label(f, text=f"${monto:,.0f}".replace(",","."),
                         font=("Helvetica", 9, "bold"),
                         fg=color_total,
                         bg=COLORES["blanco"], anchor="e").pack(side="right")

        # Haberes
        haberes_items = [("Sueldo base", r["sueldo_base_proporcional"])]
        if r["bono_colacion"]:     haberes_items.append(("Bono colación", r["bono_colacion"]))
        if r["bono_movilizacion"]: haberes_items.append(("Bono movilización", r["bono_movilizacion"]))
        if r["bono_asistencia"]:   haberes_items.append(("Bono asistencia perfecta", r["bono_asistencia"]))
        if r["asig_familiar"]:     haberes_items.append(("Asignación familiar", r["asig_familiar"]))
        if r["valor_horas_extra"]: haberes_items.append(("Horas extras", r["valor_horas_extra"]))
        seccion("HABERES", haberes_items)

        # Descuentos legales
        desc_legales = [
            ("AFP (10.00%)",       r["desc_afp"]),
            ("Salud (7.00%)",      r["desc_salud"]),
            ("Seg. Cesantía (0.6%)", r["desc_cesantia"]),
        ]
        seccion("DESCUENTOS LEGALES", desc_legales, COLORES["rojo"])

        # Descuentos adicionales
        desc_adicionales = []
        if r["desc_prestamo"]:  desc_adicionales.append(("Préstamo empresa",   r["desc_prestamo"]))
        if r["desc_adelanto"]:  desc_adicionales.append(("Adelanto de sueldo", r["desc_adelanto"]))
        if r["desc_otros"]:     desc_adicionales.append(("Otros descuentos",   r["desc_otros"]))
        if desc_adicionales:
            seccion("OTROS DESCUENTOS", desc_adicionales, COLORES["rojo"])

        # Totales
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=20, pady=6)
        totales = [
            ("Total haberes",      r["total_haberes"]),
            ("Total descuentos",   r["total_descuentos"]),
        ]
        for nombre, monto in totales:
            f = tk.Frame(self, bg=COLORES["blanco"])
            f.pack(fill="x", padx=20, pady=1)
            tk.Label(f, text=nombre, font=("Helvetica", 10, "bold"),
                     bg=COLORES["blanco"], anchor="w").pack(side="left")
            tk.Label(f, text=f"${monto:,.0f}".replace(",","."),
                     font=("Helvetica", 10, "bold"),
                     bg=COLORES["blanco"], anchor="e").pack(side="right")

        # Líquido a pagar
        liquido_frame = tk.Frame(self, bg=COLORES["verde_oscuro"])
        liquido_frame.pack(fill="x", padx=20, pady=8)
        tk.Label(liquido_frame, text="LÍQUIDO A PAGAR",
                 font=("Helvetica", 12, "bold"),
                 fg=COLORES["blanco"], bg=COLORES["verde_oscuro"],
                 padx=14, pady=8).pack(side="left")
        tk.Label(liquido_frame,
                 text=f"${r['liquido_a_pagar']:,.0f}".replace(",","."),
                 font=("Helvetica", 14, "bold"),
                 fg=COLORES["amarillo"], bg=COLORES["verde_oscuro"],
                 padx=14, pady=8).pack(side="right")

        tk.Button(self, text="✖  Cerrar", font=("Helvetica", 10, "bold"),
                  bg=COLORES["gris_texto"], fg=COLORES["blanco"],
                  relief="flat", padx=20, pady=6,
                  cursor="hand2", command=self.destroy).pack(pady=10)


# ─────────────────────────────────────────────
#  PUNTO DE ENTRADA
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = AppLiquidacion()
    app.mainloop()
