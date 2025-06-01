import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
import datetime
import os
import math  # Para cálculo de columnas de leyenda

class App:
    def __init__(self, root):
        self.root = root
        root.title("Control de Bombas - Intervalo de 5 minutos con Fecha Completa")

        # Rutas dinámicas
        self.filepath = None
        self.output_pdf_path = r"C:\Users\HZUBZU0T\Desktop\puntoexe\ultimo_intervalo.pdf"  # PDF fijo
        self.output_xlsx_path = r"C:\Users\HZUBZU0T\Desktop\puntoexe\acumulado_intervalos.xlsx"  # XLSX de salida

        # Botón para seleccionar archivo TXT
        select_btn = ttk.Button(root, text="Seleccionar TXT...", command=self.seleccionar_txt)
        select_btn.pack(padx=10, pady=5, anchor='w')

        # Frame para fecha y hora de apagado de bomba (separado)
        fm = ttk.LabelFrame(root, text="Fecha y hora de apagado de bomba")
        fm.pack(padx=10, pady=5, fill='x')
        # Día, Mes, Año
        ttk.Label(fm, text="Día:").grid(row=0, column=0, padx=3, pady=3)
        self.day_entry = ttk.Entry(fm, width=5)
        self.day_entry.grid(row=0, column=1, padx=3, pady=3)
        ttk.Label(fm, text="Mes:").grid(row=0, column=2, padx=3, pady=3)
        self.month_entry = ttk.Entry(fm, width=5)
        self.month_entry.grid(row=0, column=3, padx=3, pady=3)
        ttk.Label(fm, text="Año:").grid(row=0, column=4, padx=3, pady=3)
        self.year_entry = ttk.Entry(fm, width=7)
        self.year_entry.grid(row=0, column=5, padx=3, pady=3)
        # Hora, Minuto
        ttk.Label(fm, text="Hora (0-23):").grid(row=1, column=0, padx=3, pady=3)
        self.hora_entry = ttk.Entry(fm, width=5)
        self.hora_entry.grid(row=1, column=1, padx=3, pady=3)
        ttk.Label(fm, text="Minuto (0-59):").grid(row=1, column=2, padx=3, pady=3)
        self.min_entry = ttk.Entry(fm, width=5)
        self.min_entry.grid(row=1, column=3, padx=3, pady=3)

        # Botón para procesar (leer, graficar, generar PDF y XLSX)
        btn = ttk.Button(root, text="Graficar 5 minutos", command=self.procesar_intervalo)
        btn.pack(padx=10, pady=10)

    def seleccionar_txt(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo TXT",
            filetypes=[("Archivos de texto", "*.txt"), ("Todos", "*.*")]
        )
        if path:
            self.filepath = path
            messagebox.showinfo("TXT seleccionado", f"Archivo: {os.path.basename(path)}")

    def procesar_intervalo(self):
        # Verificar que se haya seleccionado archivo
        if not self.filepath or not os.path.isfile(self.filepath):
            messagebox.showerror("Error", "Debes seleccionar un archivo TXT primero.")
            return
        df = self.leer_datos_txt(self.filepath)
        if df is None or df.empty:
            messagebox.showerror("Error", "No se pudieron cargar los datos del archivo.")
            return

        # Leer y validar fecha/hora desde campos separados
        d = self.day_entry.get().strip()
        mo = self.month_entry.get().strip()
        y = self.year_entry.get().strip()
        h = self.hora_entry.get().strip()
        m = self.min_entry.get().strip()
        if not all([d.isdigit(), mo.isdigit(), y.isdigit(), h.isdigit(), m.isdigit()]):
            messagebox.showerror("Entrada inválida", "Introduce valores numéricos para día, mes, año, hora y minuto.")
            return
        day, month, year = int(d), int(mo), int(y)
        hour, minute = int(h), int(m)
        try:
            start_dt = datetime.datetime(year, month, day, hour, minute)
        except ValueError:
            messagebox.showerror("Fecha inválida", "La fecha y hora introducidas no son válidas.")
            return

        df_sorted = df.sort_index()
        mask = df_sorted.index >= start_dt
        if not mask.any():
            messagebox.showinfo("Sin datos", f"No hay registros desde {start_dt.strftime('%Y-%m-%d %H:%M')} en adelante.")
            return
        start_idx = df_sorted.index[mask][0]
        end_dt = start_idx + datetime.timedelta(minutes=5)
        data = df_sorted.loc[start_idx:end_dt]
        if data.empty:
            messagebox.showinfo("Sin datos", "No hay datos en el intervalo seleccionado.")
            return

        # Mostrar en GUI, generar PDF y guardar en XLSX
        self.mostrar_grafico(data, start_idx, end_dt)
        self.generar_pdf(data, start_idx, end_dt)
        self.guardar_xlsx(data)

    def leer_datos_txt(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = [l for l in f if l.strip()]
            header = lines[0].strip().split('\t')
            n = len(header)
            rows = []
            for l in lines[1:]:
                parts = l.strip().split('\t')
                parts += ['0'] * (n - len(parts))
                rows.append(parts[:n])
            df = pd.DataFrame(rows, columns=header)
            for c in header:
                if c not in ['Date','Time']:
                    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            df['DateTime'] = pd.to_datetime(
                df['Date'] + ' ' + df['Time'],
                format='%m/%d/%Y %I:%M:%S %p',
                errors='coerce'
            )
            df.set_index('DateTime', inplace=True)
            return df.drop(columns=['Date','Time'])
        except Exception as e:
            messagebox.showerror("Error al leer", str(e))
            return None

    def mostrar_grafico(self, data, start_dt, end_dt):
        top = tk.Toplevel(self.root)
        top.title(f"Datos {start_dt.strftime('%Y-%m-%d %H:%M')} a {end_dt.strftime('%H:%M')}")

        # Tabla de valores
        vf = ttk.LabelFrame(top, text="Valores")
        vf.pack(fill='x', padx=10, pady=5)
        for j in range(5): vf.columnconfigure(j, weight=1)
        hdrs = ["Parámetro","Inicio","Fin","Diferencia","Estado"]
        for j,h in enumerate(hdrs):
            ttk.Label(vf, text=h, anchor='center').grid(row=0, column=j, sticky='ew', padx=5)
        fr, lr = data.iloc[0], data.iloc[-1]
        for i, c in enumerate(data.columns, start=1):
            i0, i1 = fr[c], lr[c]
            d = i1 - i0
            s = '✗' if abs(d)>50 else '✓'
            clr = 'red' if abs(d)>50 else 'green'
            vals = [c, f"{i0:.2f}", f"{i1:.2f}", f"{d:.2f}", s]
            for j,v in enumerate(vals):
                ttk.Label(vf, text=v, foreground=clr if j==4 else None, anchor='center').grid(row=i, column=j, sticky='ew', padx=5)

        # Gráfico relativo
        rel = data.subtract(fr)
        fig = Figure(figsize=(8,5), dpi=100)
        ax = fig.add_subplot(111)
        for c in rel.columns: ax.plot(rel.index, rel[c], label=c)
        ax.set_xlabel('DateTime'); ax.set_ylabel('Cambio respecto al inicio')
        ax.set_title('Intervalo de 5 minutos (inicio = 0)')
        ncol = math.ceil(len(rel.columns)/2)
        ax.legend(
            loc='lower center', bbox_to_anchor=(0.5,-0.3), ncol=ncol,
            fontsize='x-small', frameon=False, markerscale=0.5,
            handlelength=1, handletextpad=0.2, columnspacing=0.2, labelspacing=0.1
        )
        fig.subplots_adjust(bottom=0.35)
        fig.autofmt_xdate(rotation=45)
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.get_tk_widget().pack(fill='both', expand=True)
        canvas.draw()

    def generar_pdf(self, data, start_dt, end_dt):
        try:
            with PdfPages(self.output_pdf_path) as pdf:
                fig = Figure(figsize=(8,10), dpi=100)
                fig.suptitle(f"Datos {start_dt.strftime('%Y-%m-%d %H:%M')} a {end_dt.strftime('%H:%M')}", y=0.95)
                gs = fig.add_gridspec(2, height_ratios=[1,2], top=0.90, bottom=0.05)
                # tabla
                ax_t = fig.add_subplot(gs[0]); ax_t.axis('off')
                rows=[]; cols=['Parámetro','Inicio','Fin','Diferencia','Estado']
                fr, lr = data.iloc[0], data.iloc[-1]
                for c in data.columns:
                    i0, i1 = fr[c], lr[c]; d=i1-i0; s='✗' if abs(d)>50 else '✓'
                    rows.append([c,f"{i0:.2f}",f"{i1:.2f}",f"{d:.2f}",s])
                tbl=ax_t.table(cellText=rows,colLabels=cols,loc='center')
                tbl.auto_set_font_size(False); tbl.set_fontsize(8); tbl.scale(1,1.5)
                # gráfico
                ax_c = fig.add_subplot(gs[1])
                rel = data.subtract(fr)
                for c in rel.columns: ax_c.plot(rel.index,rel[c],label=c)
                ax_c.set_xlabel('DateTime'); ax_c.set_ylabel('Cambio respecto al inicio')
                ax_c.set_title('Intervalo de 5 minutos (inicio = 0)')
                ncol2 = math.ceil(len(rel.columns)/2)
                ax_c.legend(loc='lower center', bbox_to_anchor=(0.5,-0.25), ncol=ncol2, fontsize='x-small', frameon=False)
                fig.subplots_adjust(hspace=0.3); fig.autofmt_xdate(rotation=45)
                pdf.savefig(fig,bbox_inches='tight')
            messagebox.showinfo('Éxito', f'PDF guardado en: {self.output_pdf_path}')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo guardar el PDF: {e}')

    def guardar_xlsx(self, data):
        # Guarda o añade las filas actuales en un XLSX acumulativo con formato de 12 señales y registro de nombres
        df5 = data.reset_index().copy()  # Tener DateTime como columna
        nombre_ciclo = os.path.splitext(os.path.basename(self.filepath))[0]
        tiempo_carga = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        orig_cols = list(data.columns)

        rows = []
        for _, row in df5.iterrows():
            base = {
                'nombre_ciclo': nombre_ciclo,
                'tiempo': row['DateTime'],
                'tiempo_carga': tiempo_carga,
                # nombre_registro es lista de columnas originales separadas por comas
                'nombre_registro': ",".join(orig_cols)
            }
            # Asignar valores a señal_1..señal_12
            for i in range(12):
                key = f'señal_{i+1}'
                if i < len(orig_cols):
                    base[key] = row[orig_cols[i]]
                else:
                    base[key] = 0
            rows.append(base)

        df_formateado = pd.DataFrame(rows)
        # Definir orden de columnas
        cols_order = ['nombre_ciclo', 'nombre_registro', 'tiempo', 'tiempo_carga'] + [f'señal_{i+1}' for i in range(12)]
        df_formateado = df_formateado[cols_order]

        if os.path.exists(self.output_xlsx_path):
            try:
                existing = pd.read_excel(self.output_xlsx_path)
                # Asegurarse de que existing tiene mismas columnas en mismo orden
                existing = existing.reindex(columns=cols_order, fill_value=0)
                combined = pd.concat([existing, df_formateado], ignore_index=True)
                combined.to_excel(self.output_xlsx_path, index=False)
            except Exception as e:
                messagebox.showerror('Error', f'No se pudo actualizar XLSX: {e}')
        else:
            try:
                df_formateado.to_excel(self.output_xlsx_path, index=False)
            except Exception as e:
                messagebox.showerror('Error', f'No se pudo crear XLSX: {e}')

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
