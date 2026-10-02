from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURACIÓN
# ============================================================

# Aquí colocas los Excel generados por el algoritmo anterior
CARPETA_ENTRADA = Path("input")

# Aquí se generará el Excel final
CARPETA_SALIDA = Path("output")

ARCHIVO_SALIDA = CARPETA_SALIDA / "porcentajes_notas.xlsx"


# Las áreas/cursos que vienen en los Excel anteriores
MATERIAS = [
    "Desarrollo Personal, Ciudadanía y Cívica",
    "Ciencias Sociales",
    "Educación para el Trabajo",
    "Educación Física",
    "Comunicación",
    "Arte y Cultura",
    "Castellano como segunda lengua",
    "Inglés",
    "Matemática",
    "Ciencia y Tecnología",
    "Educación Religiosa",
    "Comp. Trans.",
]

NOTAS = ["AD", "A", "B", "C"]


# ============================================================
# BUSCAR EXCEL
# ============================================================

def buscar_excels():
    archivos = sorted(
        CARPETA_ENTRADA.glob("*.xlsx")
    )

    # Ignorar archivos temporales de Excel
    archivos = [
        archivo
        for archivo in archivos
        if not archivo.name.startswith("~$")
    ]

    if not archivos:
        raise FileNotFoundError(
            "No se encontraron archivos .xlsx en la carpeta input."
        )

    return archivos


# ============================================================
# LEER LOS EXCEL DEL ALGORITMO ANTERIOR
# ============================================================

def leer_archivo(ruta_excel):
    """
    Lee la hoja Promedios de un Excel generado por el
    algoritmo anterior.

    Cada Excel representa una sección.
    Ejemplo:
        1A
        1B
        1C
    """

    try:
        df = pd.read_excel(
            ruta_excel,
            sheet_name="Promedios",
        )
    except Exception as error:
        raise RuntimeError(
            f"No se pudo leer {ruta_excel.name}: {error}"
        )

    columnas_necesarias = [
        "Grado",
        "Sección",
        "Estudiante",
    ] + MATERIAS

    faltantes = [
        columna
        for columna in columnas_necesarias
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"{ruta_excel.name} no tiene estas columnas: "
            + ", ".join(faltantes)
        )

    return df


# ============================================================
# NORMALIZAR NOTA
# ============================================================

def normalizar_nota(valor):
    if pd.isna(valor):
        return ""

    return str(valor).strip().upper()


# ============================================================
# PROCESAR TODOS LOS EXCEL
# ============================================================

def procesar_archivos(archivos):
    """
    Lee todos los Excel por separado y junta sus registros
    únicamente para poder hacer el cálculo por grado.

    NO modifica los Excel originales.
    """

    registros = []

    for numero, ruta_excel in enumerate(
        archivos,
        start=1,
    ):

        print()
        print("-" * 70)
        print(
            f"LEYENDO EXCEL {numero}/{len(archivos)}"
        )
        print(
            f"Archivo: {ruta_excel.name}"
        )
        print("-" * 70)

        df = leer_archivo(ruta_excel)

        print(
            f"Estudiantes: {len(df)}"
        )

        for _, fila in df.iterrows():

            try:
                grado = int(float(fila["Grado"]))
            except (ValueError, TypeError):
                continue

            seccion = str(
                fila["Sección"]
            ).strip().upper()

            registro = {
                "Grado": grado,
                "Sección": seccion,
                "Estudiante": fila["Estudiante"],
            }

            for materia in MATERIAS:
                registro[materia] = normalizar_nota(
                    fila[materia]
                )

            registros.append(registro)

    return pd.DataFrame(registros)


# ============================================================
# CALCULAR PORCENTAJES
# ============================================================

def calcular_porcentajes(df):
    """
    Para cada grado y cada curso:

        1° = 1A + 1B + 1C
        2° = 2A + 2B + 2C
        etc.

    Ejemplo para un curso:

        AD = 3
        A  = 47
        B  = 28
        C  = 4

        TOTAL = 82

        % AD = 3 / 82 * 100
        % A  = 47 / 82 * 100
        % B  = 28 / 82 * 100
        % C  = 4 / 82 * 100

    Los vacíos NO se cuentan.
    """

    resultados = []

    for grado in sorted(
        df["Grado"].dropna().unique()
    ):

        datos_grado = df[
            df["Grado"] == grado
        ]

        secciones = sorted(
            datos_grado["Sección"]
            .dropna()
            .astype(str)
            .unique()
        )

        for materia in MATERIAS:

            notas = (
                datos_grado[materia]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            cantidad_ad = int(
                (notas == "AD").sum()
            )

            cantidad_a = int(
                (notas == "A").sum()
            )

            cantidad_b = int(
                (notas == "B").sum()
            )

            cantidad_c = int(
                (notas == "C").sum()
            )

            total = (
                cantidad_ad
                + cantidad_a
                + cantidad_b
                + cantidad_c
            )

            if total > 0:

                porcentaje_ad = round(
                    cantidad_ad / total * 100,
                    2,
                )

                porcentaje_a = round(
                    cantidad_a / total * 100,
                    2,
                )

                porcentaje_b = round(
                    cantidad_b / total * 100,
                    2,
                )

                porcentaje_c = round(
                    cantidad_c / total * 100,
                    2,
                )

            else:

                porcentaje_ad = 0
                porcentaje_a = 0
                porcentaje_b = 0
                porcentaje_c = 0

            resultados.append({
                "Grado": f"{grado}°",
                "Secciones": ", ".join(secciones),
                "Área / Curso": materia,

                "AD": cantidad_ad,
                "% AD": porcentaje_ad,

                "A": cantidad_a,
                "% A": porcentaje_a,

                "B": cantidad_b,
                "% B": porcentaje_b,

                "C": cantidad_c,
                "% C": porcentaje_c,

                "Total": total,
            })

    return pd.DataFrame(resultados)


# ============================================================
# GENERAR EXCEL SIMPLE
# ============================================================

def crear_salida_pegable(df_resultado):
    """
    Crea el formato exactamente pensado para copiar/pegar
    en el formato oficial:

        Nivel | Grado | Área / Curso | Resultado

    Cada curso ocupa 4 filas:
        AD: xx.xx%
        A:  xx.xx%
        B:  xx.xx%
        C:  xx.xx%

    Los campos Nivel, Grado y Área se repiten para facilitar
    el pegado directo en Excel.
    """

    filas = []

    for _, fila in df_resultado.iterrows():

        grado_texto = str(fila["Grado"]).strip()

        # Para este proyecto los grados 1° a 5° corresponden
        # a Secundaria.
        nivel = "Secundaria"

        materia = fila["Área / Curso"]

        filas.append({
            "Nivel": nivel,
            "Grado": grado_texto,
            "Área / Curso": materia,
            "Resultado": f"AD:{round(float(fila['% AD']))}%",
        })
        
        filas.append({
            "Nivel": nivel,
            "Grado": grado_texto,
            "Área / Curso": materia,
            "Resultado": f"A:{round(float(fila['% A']))}%",
        })
        
        filas.append({
            "Nivel": nivel,
            "Grado": grado_texto,
            "Área / Curso": materia,
            "Resultado": f"B:{round(float(fila['% B']))}%",
        })
        
        filas.append({
            "Nivel": nivel,
            "Grado": grado_texto,
            "Área / Curso": materia,
            "Resultado": f"C:{round(float(fila['% C']))}%",
        })

    return pd.DataFrame(filas)


def generar_excel(df_resultado):

    CARPETA_SALIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    df_pegable = crear_salida_pegable(
        df_resultado
    )

    with pd.ExcelWriter(
        ARCHIVO_SALIDA,
        engine="openpyxl",
    ) as writer:

        df_pegable.to_excel(
            writer,
            index=False,
            sheet_name="Para pegar",
        )

    aplicar_formato_excel(
        ARCHIVO_SALIDA
    )


# ============================================================
# FORMATO
# ============================================================

def aplicar_formato_excel(ruta):

    libro = load_workbook(ruta)

    hoja = libro["Para pegar"]

    # El formato se deja simple para copiar y pegar.
    hoja.freeze_panes = "A2"

    hoja.sheet_view.showGridLines = True

    # Ocultar los encabezados para que al seleccionar/copiar
    # se pueda usar directamente el contenido de las 4 columnas.
    hoja.auto_filter.ref = None

    for celda in hoja[1]:

        celda.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )

        celda.font = Font(
            bold=True,
            color="FFFFFF",
        )

        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    hoja.row_dimensions[1].height = 25

    hoja.column_dimensions["A"].width = 18
    hoja.column_dimensions["B"].width = 14
    hoja.column_dimensions["C"].width = 38
    hoja.column_dimensions["D"].width = 18

    for fila in hoja.iter_rows(
        min_row=2,
    ):

        for celda in fila:

            celda.alignment = Alignment(
                vertical="center",
                wrap_text=True,
            )

            celda.border = Border(
                bottom=Side(
                    style="thin",
                    color="D9E1F2",
                )
            )

    # Mantener juntas visualmente las cuatro filas de cada curso.
    for fila_inicio in range(
        2,
        hoja.max_row + 1,
        4,
    ):

        fila_fin = min(
            fila_inicio + 3,
            hoja.max_row,
        )

        for columna in range(1, 5):

            for fila in range(
                fila_inicio,
                fila_fin + 1,
            ):

                celda = hoja.cell(
                    fila,
                    columna,
                )

                celda.alignment = Alignment(
                    vertical="center",
                    horizontal="left"
                    if columna != 4
                    else "left",
                    wrap_text=True,
                )

    libro.save(ruta)


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print("   PORCENTAJES DE NOTAS AD - A - B - C")
    print("=" * 70)

    CARPETA_ENTRADA.mkdir(
        parents=True,
        exist_ok=True,
    )

    CARPETA_SALIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    archivos = buscar_excels()

    print()
    print(
        f"EXCEL ENCONTRADOS EN INPUT: {len(archivos)}"
    )

    for numero, archivo in enumerate(
        archivos,
        start=1,
    ):

        print(
            f"  {numero:02d}. {archivo.name}"
        )

    df = procesar_archivos(
        archivos
    )

    if df.empty:
        print()
        print(
            "ERROR: no se encontraron estudiantes."
        )
        return

    print()
    print("=" * 70)
    print("CALCULANDO")
    print("=" * 70)

    resultado = calcular_porcentajes(
        df
    )

    generar_excel(
        resultado
    )

    print()
    print("=" * 70)
    print("PROCESO FINALIZADO")
    print("=" * 70)

    print(
        f"Excel generados leídos: {len(archivos)}"
    )

    print(
        f"Registros de estudiantes: {len(df)}"
    )

    print(
        f"Resultado para pegar: {ARCHIVO_SALIDA}"
    )


if __name__ == "__main__":
    main()
