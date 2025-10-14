import os
import warnings
import pandas as pd
from docxtpl import DocxTemplate

# Silenciar warnings de docxcompose
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")

excel_file = "datos.xlsx"              # cambia al nombre real de tu Excel
template_file = "AAAA-MEMORIA TIPO SEP ZIER R0.docx"    # cambia al nombre real de tu Word plantilla
output_folder = "salidas_word"

try:
    os.makedirs(output_folder, exist_ok=True)

    # Leer Excel
    df = pd.read_excel(excel_file)
    print(f"📖 Excel leído con {len(df)} filas")
    print("👉 Columnas detectadas en Excel:", list(df.columns))

    # Nombre base de la plantilla
    template_name, ext = os.path.splitext(os.path.basename(template_file))

    for index, row in df.iterrows():
        print(f"\n➡️ Procesando fila {index+1}...")

        # Convertir fila en diccionario
        context = {}
        for k, v in row.items():
            if isinstance(v, float) and v.is_integer():
                context[k] = int(v)
            else:
                context[k] = v

        print("Contexto:", context)

        # Verificar que código exista
        if "codigo_proyecto" not in context:
            raise KeyError("⚠️ La columna 'codigo_proyecto' no existe en el Excel")

        codigo = str(context["codigo_proyecto"])
        print(f"Código proyecto: {codigo}")

        # Abrir plantilla
        if not os.path.exists(template_file):
            raise FileNotFoundError(f"⚠️ No se encuentra la plantilla: {template_file}")
        doc = DocxTemplate(template_file)

        # Rellenar
        doc.render(context)

        # Nombre del nuevo archivo
        new_name = template_name.replace("AAAA", codigo, 1) + ext
        filename = os.path.join(output_folder, new_name)

        doc.save(filename)
        print(f"✅ Documento generado: {filename}")

except Exception as e:
    print("❌ Error crítico:", e)
