from __future__ import annotations

import queue
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from modules.afip_consolidator.controller import ConsolidatorController
from modules.afip_consolidator.importers import EXTENSIONES
from modules.afip_consolidator.models import ModoSalida


class ConsolidatorApp(ctk.CTk):
    """Interfaz: un trabajador y una cola; todas las actualizaciones ocurren en Tk."""

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        self.title("Consolidador de reportes Excel y CSV")
        self.geometry("940x690")
        self.minsize(760, 560)
        self.controller = ConsolidatorController()
        self.eventos = queue.Queue()
        self.ocupado = False
        self.modo = ctk.StringVar(value=ModoSalida.SABANA.value)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(self, text="Consolidador de reportes", font=ctk.CTkFont(size=26, weight="bold")).grid(
            row=0, column=0, padx=24, pady=(24, 4), sticky="w")
        ctk.CTkLabel(self, text="Consolidador de Excel y CSV",
                     anchor="w").grid(row=1, column=0, padx=24, pady=(0, 16), sticky="ew")
        panel = ctk.CTkFrame(self)
        panel.grid(row=2, column=0, padx=24, sticky="ew")
        panel.grid_columnconfigure((0, 1, 2), weight=1)
        self.agregar = ctk.CTkButton(panel, text="Agregar archivos", command=self.seleccionar)
        self.carpeta = ctk.CTkButton(panel, text="Agregar carpeta", command=self.seleccionar_carpeta)
        self.limpiar = ctk.CTkButton(panel, text="Limpiar selección", command=lambda: self._seleccion([]))
        for columna, boton in enumerate((self.agregar, self.carpeta, self.limpiar)):
            boton.grid(row=0, column=columna, padx=10, pady=12, sticky="ew")
        opciones = ctk.CTkFrame(self, fg_color="transparent")
        opciones.grid(row=3, column=0, padx=24, pady=12, sticky="ew")
        opciones.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(opciones, text="Formato de salida:").grid(row=0, column=0, padx=(0, 12))
        self.selector = ctk.CTkSegmentedButton(opciones, values=[m.value for m in ModoSalida], variable=self.modo)
        self.selector.grid(row=0, column=1, sticky="ew")
        periodo = ctk.CTkFrame(opciones, fg_color="transparent")
        periodo.grid(row=1, column=0, columnspan=2, pady=(12, 0), sticky="ew")
        self.validar_periodo = ctk.BooleanVar(value=False)
        self.desde = ctk.StringVar()
        self.hasta = ctk.StringVar()
        self.check_periodo = ctk.CTkCheckBox(periodo, text="Comprobar meses faltantes", variable=self.validar_periodo)
        self.check_periodo.grid(row=0, column=0, padx=(0, 16))
        self.entrada_desde = ctk.CTkEntry(periodo, textvariable=self.desde, placeholder_text="Desde MM/AAAA", width=150)
        self.entrada_hasta = ctk.CTkEntry(periodo, textvariable=self.hasta, placeholder_text="Hasta MM/AAAA", width=150)
        self.entrada_desde.grid(row=0, column=1, padx=6)
        self.entrada_hasta.grid(row=0, column=2, padx=6)
        self.vista = ctk.CTkTextbox(self, wrap="word")
        self.vista.grid(row=4, column=0, padx=24, pady=(0, 12), sticky="nsew")
        self.estado = ctk.CTkLabel(self, text="Seleccioná archivos para comenzar.", anchor="w")
        self.estado.grid(row=5, column=0, padx=24, sticky="ew")
        self.progreso = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progreso.grid(row=6, column=0, padx=24, pady=8, sticky="ew")
        self.progreso.set(0)
        acciones = ctk.CTkFrame(self, fg_color="transparent")
        acciones.grid(row=7, column=0, padx=24, pady=(4, 24), sticky="ew")
        acciones.grid_columnconfigure((0, 1), weight=1)
        self.preparar = ctk.CTkButton(acciones, text="1. Preparar y revisar", height=42, command=self.preparar_seleccion)
        self.guardar = ctk.CTkButton(acciones, text="2. Guardar consolidado", height=42, command=self.guardar_resultado)
        self.preparar.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        self.guardar.grid(row=0, column=1, padx=(6, 0), sticky="ew")
        self.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._seleccion([])
        for variable in (self.validar_periodo, self.desde, self.hasta):
            variable.trace_add("write", self._cambio_periodo)
        self.after(100, self._recibir)

    def _texto(self, texto):
        self.vista.configure(state="normal")
        self.vista.delete("1.0", "end")
        self.vista.insert("1.0", texto)
        self.vista.configure(state="disabled")

    def _controles(self):
        estado = "disabled" if self.ocupado else "normal"
        for control in (self.agregar, self.carpeta, self.limpiar, self.selector,
                        self.check_periodo, self.entrada_desde, self.entrada_hasta):
            control.configure(state=estado)
        self.preparar.configure(state="normal" if not self.ocupado and self.controller.rutas else "disabled")
        self.guardar.configure(state="normal" if not self.ocupado and self.controller.resultado is not None else "disabled")

    def _seleccion(self, rutas):
        self.controller.seleccionar(rutas)
        self._texto("ARCHIVOS SELECCIONADOS\n\n" + ("\n".join(str(r) for r in self.controller.rutas) or
                    "Agregá reportes con columnas Fecha y Tipo / Tipo de Comprobante.\n"
                    "Cada Excel debe tener una sola hoja con datos."))
        self.estado.configure(text=f"{len(self.controller.rutas)} archivo(s). Prepará la selección antes de exportar.")
        self._controles()

    def seleccionar(self):
        rutas = filedialog.askopenfilenames(parent=self, title="Seleccionar reportes",
                    filetypes=[("Reportes Excel y CSV", "*.xlsx *.xlsm *.xls *.csv")])
        if rutas:
            self._seleccion([*self.controller.rutas, *rutas])

    def seleccionar_carpeta(self):
        ruta = filedialog.askdirectory(parent=self, title="Carpeta de entrada (sin subcarpetas)")
        if not ruta:
            return
        try:
            archivos = sorted(p for p in Path(ruta).iterdir() if p.is_file()
                              and p.suffix.lower() in EXTENSIONES and not p.name.startswith(("~$", ".")))
            if not archivos:
                messagebox.showinfo("Carpeta vacía", "No hay archivos compatibles en esta carpeta.", parent=self)
                return
            self._seleccion([*self.controller.rutas, *archivos])
        except OSError as exc:
            messagebox.showerror("No se pudo abrir la carpeta", str(exc), parent=self)

    def _ejecutar(self, tipo, funcion):
        if self.ocupado:
            return
        self.ocupado = True
        self._controles()
        self.progreso.start()
        self.estado.configure(text="Preparando archivos…" if tipo == "preparado" else "Guardando consolidado…")
        def trabajo():
            try:
                self.eventos.put((tipo, funcion()))
            except Exception as exc:
                self.eventos.put(("error", str(exc)))
        threading.Thread(target=trabajo, daemon=True).start()

    def preparar_seleccion(self):
        self.controller.resultado = None
        self._texto("Procesando la selección actual…")
        periodo = (self.desde.get(), self.hasta.get()) if self.validar_periodo.get() else None
        self._ejecutar("preparado", lambda: self.controller.preparar(periodo=periodo))

    def _cambio_periodo(self, *_):
        self.controller.resultado = None
        self.estado.configure(text="Cambió el período. Volvé a preparar para actualizar la revisión.")
        self._controles()

    def guardar_resultado(self):
        ruta = filedialog.asksaveasfilename(parent=self, title="Guardar consolidado",
                     defaultextension=".xlsx", initialfile="Consolidado.xlsx",
                     filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        reemplazar = Path(ruta).exists()
        if reemplazar and not messagebox.askyesno("Reemplazar archivo", "El destino existe. ¿Querés reemplazarlo?", parent=self):
            return
        modo = ModoSalida(self.modo.get())
        self._ejecutar("guardado", lambda: self.controller.exportar(ruta, modo, reemplazar=reemplazar))

    def _recibir(self):
        try:
            tipo, valor = self.eventos.get_nowait()
        except queue.Empty:
            pass
        else:
            self.ocupado = False
            self.progreso.stop()
            self.progreso.set(0)
            self._controles()
            if tipo == "error":
                self.estado.configure(text="No se completó la operación. Revisá el mensaje.")
                messagebox.showerror("No se pudo completar", valor, parent=self)
            elif tipo == "preparado":
                lineas = [f"{a.ruta.name} — {a.filas} filas" for a in valor.archivos]
                self._texto(f"LISTO PARA EXPORTAR\n\n{len(valor.archivos)} archivos · {len(valor.comprobantes)} filas\n\n"
                            + "\n".join(lineas) + "\n\nAVISOS\n\n"
                            + ("\n\n".join(valor.avisos) or "Sin advertencias técnicas.")
                            + "\n\nSe conservan las filas repetidas. Las revisiones quedan solo en esta pantalla.")
                self.estado.configure(text="Revisá los avisos y guardá el consolidado.")
            else:
                self.estado.configure(text="Consolidado guardado correctamente.")
                messagebox.showinfo("Archivo guardado", f"Archivo generado:\n{valor}", parent=self)
        self.after(100, self._recibir)

    def cerrar(self):
        if self.ocupado:
            messagebox.showinfo("Operación en curso", "Esperá a que finalice la operación antes de cerrar.", parent=self)
        else:
            self.destroy()
