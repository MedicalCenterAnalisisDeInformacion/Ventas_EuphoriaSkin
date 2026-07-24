import json
import calendar
import unicodedata
from pathlib import Path

import pandas as pd
import numpy as np
from datetime import date, timedelta

# Ajustar rutas y el párametro
EXCEL_PATH  = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_Euphoria/07_Julio/24-07-2026/Dataset.xlsx"
OUTPUT_PATH = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_Euphoria/07_Julio/24-07-2026/index.html"
BOL_EXCLUIR = ["BOLEUCH", "BOLEUGDE", "BOLEUMIN"]
FECHA_BASE  = date(2026, 7, 24)

# Lógica de procesamiento
MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]
DIAS_ES  = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]

MES_A_NUM = {
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
    "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12,
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12
}


def _margen_seguro(utilidad, ventas):
    """Calcula utilidad/ventas evitando Infinity (ventas=0 con utilidad!=0) y NaN (0/0)."""
    m = (utilidad / ventas).replace([np.inf, -np.inf], 0).fillna(0)
    return m.round(4)


def _mapear_mes(serie_mes, origen: str):
    """Mapea la columna 'Mes' (texto) a número de mes, avisando si hay valores no reconocidos."""
    serie_norm = serie_mes.str.strip().str.lower()
    mes_num = serie_norm.map(MES_A_NUM)
    no_reconocidos = serie_norm[mes_num.isna()]
    if len(no_reconocidos):
        valores = sorted(no_reconocidos.unique().tolist())
        print(f"⚠️  [{origen}] {len(no_reconocidos)} fila(s) con valor de 'Mes' no reconocido "
              f"(se agruparán como '?'): {valores}")
    return mes_num


def _formatear_periodo(mes_num_serie, anio_serie):
    """Da formato 'Mmm-aa' a un período (mes abreviado + año de 2 dígitos),
    p. ej. Julio 2026 -> 'Jul-26'."""
    mes_abr  = mes_num_serie.apply(lambda m: MESES_ES[int(m)-1].capitalize()[:3] if m and m > 0 else "?")
    anio_abr = anio_serie.astype(int).astype(str).str[-2:]
    return mes_abr + "-" + anio_abr


def _normalizar_texto(s) -> str:
    """Normaliza nombres de sucursal para poder cruzarlos de forma robusta:
    quita espacios extra, acentos y diferencias de mayúsculas/minúsculas.
    Así 'Euphoria  Polanco', 'EUPHORIA POLANCO' y 'Euphoria Pólanco' cruzan igual."""
    s = " ".join(str(s).strip().split())
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return s.casefold()


def _detectar_columna(df, patron: str):
    """Busca, sin importar mayúsculas/acentos, la primera columna cuyo nombre contenga 'patron'."""
    patron_norm = _normalizar_texto(patron)
    for c in df.columns:
        if patron_norm in _normalizar_texto(c):
            return c
    return None


def formatear_fechas(base: date):
    ayer = base - timedelta(days=1)
    fecha_reporte = f"{base.day} de {MESES_ES[base.month-1].capitalize()} de {base.year}"
    dia_semana    = DIAS_ES[ayer.weekday()].capitalize()
    fecha_info    = f"{dia_semana}, {ayer.day} de {MESES_ES[ayer.month-1]} de {ayer.year}"
    mes_header    = MESES_ES[base.month-1].capitalize()
    return fecha_reporte, fecha_info, mes_header


def procesar_mes_curso(vmc, suc, bol_list):
    vmc["ClaveSucursal"] = pd.to_numeric(vmc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vmc = vmc.merge(suc, on="ClaveSucursal", how="left")
    vmc["FechaStr"] = pd.to_datetime(vmc["Fechas de emisión"]).dt.strftime("%Y-%m-%d")

    vmc_v = vmc[~vmc["Artículo"].isin(bol_list)]

    agg_v = vmc_v.groupby(["FechaStr","NombreSucursal"]).agg(
        unidades  =("Unidades","sum"),
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
    ).reset_index()

    agg_t = vmc.groupby(["FechaStr","NombreSucursal"]).agg(
        tickets=("Movimiento","nunique")
    ).reset_index()

    agg = agg_v.merge(agg_t, on=["FechaStr","NombreSucursal"], how="left")
    agg["tickets"] = agg["tickets"].fillna(0).astype(int)
    agg["margen"]  = _margen_seguro(agg["utilidad"], agg["ventas"])
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg


def procesar_lineas(vm, art, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]

    art_clean = art.drop_duplicates(subset=["Artículo"])
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])

    vm_merged = vm.merge(art_clean, on="Artículo", how="left").merge(suc_clean, on="ClaveSucursal", how="left")
    vm_merged["Línea"]          = vm_merged["Línea"].fillna("NO ASIGNADO")
    vm_merged["NombreSucursal"] = vm_merged["NombreSucursal"].fillna("OTRO")

    vm_merged["MesNum"] = _mapear_mes(vm_merged["Mes"], "VentasMensuales (Resumen por Línea)").fillna(0).astype(int)
    vm_merged["Año"]    = pd.to_numeric(vm_merged["Año"], errors="coerce").fillna(0).astype(int)
    vm_merged["PeriodoLabel"] = _formatear_periodo(vm_merged["MesNum"], vm_merged["Año"])

    linea_agg = vm_merged.groupby(["NombreSucursal","Línea","PeriodoLabel"]).agg(
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
        costo     =("Costo","sum"),
        unidades  =("Unidades","sum"),
    ).reset_index()

    for c in ["ventas","utilidad","costo"]:
        linea_agg[c] = linea_agg[c].round(2)
    linea_agg["margen"] = _margen_seguro(linea_agg["utilidad"], linea_agg["ventas"])
    return linea_agg


def procesar_historico(vm, tkt, suc, bol_list):
    vm["ClaveSucursal"]  = pd.to_numeric(vm["ClaveSucursal"],  errors="coerce").fillna(0).astype(int)
    tkt["ClaveSucursal"] = pd.to_numeric(tkt["ClaveSucursal"], errors="coerce").fillna(0).astype(int)

    vm = vm[~vm["Artículo"].isin(bol_list)].copy()
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales")

    agg_v = vm.groupby(["Año","MesNum","ClaveSucursal"]).agg(
        unidades  =("Unidades","sum"),
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
    ).reset_index()

    tkt["MesNum"] = _mapear_mes(tkt["Mes"], "TicketsMensuales")
    agg_t = tkt.groupby(["Año","MesNum","ClaveSucursal"]).agg(
        tickets=("Tickets","sum")
    ).reset_index()

    agg = agg_v.merge(agg_t, on=["Año","MesNum","ClaveSucursal"], how="outer")
    for c in ["unidades","ventas","utilidad"]:
        agg[c] = agg[c].fillna(0).round(2)
    agg["tickets"] = agg["tickets"].fillna(0).astype(int)
    agg["Año"]    = agg["Año"].fillna(0).astype(int)
    agg["MesNum"] = agg["MesNum"].fillna(0).astype(int)

    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    agg = agg.merge(suc_clean, on="ClaveSucursal", how="left")
    agg["NombreSucursal"] = agg["NombreSucursal"].fillna("OTRO")

    agg["margen"] = _margen_seguro(agg["utilidad"], agg["ventas"])
    agg["PeriodoLabel"] = _formatear_periodo(agg["MesNum"], agg["Año"])
    agg = agg.sort_values(["Año","MesNum"]).reset_index(drop=True)
    return agg


def procesar_top_articulos(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]

    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")

    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Descripción","Fabricante"]], on="Artículo", how="left")
    vm["Descripción"] = vm["Descripción"].fillna("SIN DESCRIPCIÓN")
    vm["Fabricante"]  = vm["Fabricante"].fillna("SIN FABRICANTE")

    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Top Artículos)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])

    agg = vm.groupby(["Artículo","Descripción","Fabricante","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()

    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg


def procesar_lineas_categoria(vm, art_dim, suc, bol_list):
    """
    Agrega ventas/utilidad/unidades a nivel Línea > Categoría, cruzado con
    Sucursal y Período (Mes/Año), para poder filtrarse en el dashboard con
    los mismos segmentadores de Sucursales y Meses que el resto de pestañas.
    """
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]

    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")

    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Línea","Categoría"]], on="Artículo", how="left")
    vm["Línea"]     = vm["Línea"].fillna("NO ASIGNADO")
    vm["Categoría"] = vm["Categoría"].fillna("SIN CATEGORÍA")

    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Líneas y Categorías)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])

    agg = vm.groupby(["Línea","Categoría","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()

    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg


def procesar_fabricantes(vm, art_dim, suc, bol_list):
    """
    Agrega ventas/utilidad/unidades a nivel Fabricante, cruzado con
    Sucursal y Período (Mes/Año), para poder filtrarse en el dashboard con
    los mismos segmentadores de Sucursales y Meses que el resto de pestañas.
    """
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]

    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")

    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Fabricante"]], on="Artículo", how="left")
    vm["Fabricante"] = vm["Fabricante"].fillna("SIN FABRICANTE")

    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Fabricantes)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])

    agg = vm.groupby(["Fabricante","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()

    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg


def procesar_presupuesto(agg, objetivos, suc, fecha_base):
    """
    Calcula, para cada sucursal CON presupuesto asignado, un pronóstico de
    cierre de mes a partir del ritmo de venta observado en 'agg' (ventas del
    mes en curso, ya sin BOL), junto con sus métricas operativas del mes.

    Lógica del pronóstico: (ventas acumuladas / días transcurridos) * días
    que la sucursal operará en el mes.

    El "inicio efectivo" de cada sucursal para este cálculo se toma de la
    columna 'FechaApertura' de la hoja 'ObjetivosVentas' (formato dd/mm/aaaa):
      - Si FechaApertura cae DENTRO del mes en curso, se usa esa fecha como
        inicio (sucursal nueva) → tanto los "días transcurridos" como los
        "días que operará en el mes" se cuentan desde ahí, no desde el día 1.
      - Si FechaApertura es anterior al mes en curso (o no se proporciona),
        se asume sucursal ya establecida y se usa el mes completo (día 1).
      - Si para alguna sucursal no hay FechaApertura registrada, se usa como
        respaldo el método anterior: inferir el inicio a partir de la PRIMERA
        fecha con venta registrada ese mes.

    El otro extremo ("ayer") se fija siempre en FECHA_BASE − 1 día, sin
    importar si hubo o no venta ese día en los datos, para que un día sin
    ventas no acorte el conteo de días transcurridos.

    Las sucursales SIN presupuesto asignado (celda vacía, en 0, o con texto
    explícito como "Sin presupuesto asignado") se EXCLUYEN por completo del
    resultado: no tiene sentido comparar contra una meta inexistente.
    """
    dias_mes = calendar.monthrange(fecha_base.year, fecha_base.month)[1]

    # "Ayer" fijo: no depende de si hubo venta ese día en los datos (blindaje
    # para que un día sin ventas de toda la cadena no acorte el cálculo).
    fecha_max_global = pd.Timestamp(fecha_base) - pd.Timedelta(days=1)
    primer_dia_mes = pd.Timestamp(year=fecha_base.year, month=fecha_base.month, day=1)
    fin_mes = pd.Timestamp(year=fecha_base.year, month=fecha_base.month, day=dias_mes)

    resumen_ventas = agg.groupby("NombreSucursal").agg(
        unidadesActual=("unidades", "sum"),
        ventasActual=("ventas", "sum"),
        utilidadActual=("utilidad", "sum"),
        fechaMin=("FechaStr", "min"),
    ).reset_index()
    resumen_ventas["fechaMinDt"] = pd.to_datetime(resumen_ventas["fechaMin"])

    # Base: TODAS las sucursales de la dimensión (para incluir también las que
    # aún no tienen venta este mes, con pronóstico 0). Se conserva ClaveSucursal
    # porque 'ObjetivosVentas' puede traer la clave numérica en vez del nombre,
    # y porque se usará para ordenar la tabla final por número de sucursal.
    suc_clean = suc.drop_duplicates(subset=["NombreSucursal"])[["ClaveSucursal", "NombreSucursal"]].copy()
    suc_clean["NombreSucursal"] = suc_clean["NombreSucursal"].astype(str).str.strip()
    resumen_ventas["NombreSucursal"] = resumen_ventas["NombreSucursal"].astype(str).str.strip()

    resumen = suc_clean.merge(resumen_ventas, on="NombreSucursal", how="left")
    for c in ["unidadesActual", "ventasActual", "utilidadActual"]:
        resumen[c] = resumen[c].fillna(0)
    resumen["margen"] = _margen_seguro(resumen["utilidadActual"], resumen["ventasActual"])

    # ── Presupuesto / Objetivo de ventas / Fecha de apertura ──
    # Detecta las columnas de forma flexible (sin depender de que se llamen
    # EXACTAMENTE "Sucursal", "Presupuesto" y "FechaApertura"; acepta variantes
    # con mayúsculas, acentos o texto adicional, p. ej. "Nombre Sucursal").
    col_suc = _detectar_columna(objetivos, "sucursal")
    col_pre = _detectar_columna(objetivos, "presupuesto")
    col_ape = _detectar_columna(objetivos, "apertura")

    resumen["FechaApertura"] = pd.NaT

    if col_suc is None or col_pre is None:
        print("⚠️  [Presupuesto] No se encontraron las columnas 'Sucursal' y/o 'Presupuesto' "
              f"en la hoja 'ObjetivosVentas' (columnas disponibles: {list(objetivos.columns)}). "
              "No se generará la tabla de pronóstico vs. presupuesto.")
        resumen["presupuesto"] = 0.0
    else:
        cols_leer = [col_suc, col_pre] + ([col_ape] if col_ape else [])
        rename_map = {col_suc: "SucursalObjetivo", col_pre: "presupuesto"}
        if col_ape:
            rename_map[col_ape] = "FechaApertura"

        obj_clean = objetivos[cols_leer].rename(columns=rename_map).dropna(subset=["SucursalObjetivo"]).copy()

        # Cualquier valor no numérico en 'Presupuesto' (celda vacía, texto como
        # "Sin presupuesto asignado", "N/A", etc.) se interpreta como SIN
        # presupuesto y queda en 0 → esa sucursal se excluirá más abajo.
        obj_clean["presupuesto"] = pd.to_numeric(obj_clean["presupuesto"], errors="coerce").fillna(0)

        if col_ape:
            obj_clean["FechaApertura"] = pd.to_datetime(obj_clean["FechaApertura"], dayfirst=True, errors="coerce")
        else:
            obj_clean["FechaApertura"] = pd.NaT
            print("⚠️  [Presupuesto] No se encontró columna 'FechaApertura' en 'ObjetivosVentas'; "
                  "se usará el método de inferencia anterior (primera venta del mes) para todas las sucursales.")

        # ¿La columna 'Sucursal' de ObjetivosVentas trae CLAVES numéricas
        # (p. ej. 30601) o NOMBRES de texto? Se decide con la mayoría de los
        # valores y se cruza por la llave correspondiente.
        claves_num = pd.to_numeric(obj_clean["SucursalObjetivo"], errors="coerce")
        usa_clave = claves_num.notna().mean() >= 0.5 if len(obj_clean) else False

        if usa_clave:
            obj_clean["ClaveSucursal"] = claves_num.fillna(0).astype(int)
            obj_clean = obj_clean.drop_duplicates(subset=["ClaveSucursal"])
            resumen = resumen.drop(columns=["FechaApertura"]).merge(
                obj_clean[["ClaveSucursal", "presupuesto", "FechaApertura"]], on="ClaveSucursal", how="left"
            )
            keys_dim  = set(resumen["ClaveSucursal"])
            keys_obj  = set(obj_clean["ClaveSucursal"])
            huerfanos = sorted(keys_obj - keys_dim)
            modo = "ClaveSucursal (código numérico)"
        else:
            resumen["_key"]     = resumen["NombreSucursal"].apply(_normalizar_texto)
            obj_clean["_key"]   = obj_clean["SucursalObjetivo"].apply(_normalizar_texto)
            obj_clean = obj_clean.drop_duplicates(subset=["_key"])
            resumen = resumen.drop(columns=["FechaApertura"]).merge(
                obj_clean[["_key", "presupuesto", "FechaApertura"]], on="_key", how="left"
            )
            keys_dim = set(resumen["_key"])
            keys_obj = set(obj_clean["_key"])
            huerfanos = sorted(
                obj_clean.loc[obj_clean["_key"].isin(keys_obj - keys_dim), "SucursalObjetivo"].unique().tolist()
            )
            resumen = resumen.drop(columns=["_key"])
            modo = "nombre de sucursal (texto normalizado)"

        if huerfanos:
            print(f"⚠️  [Presupuesto] Cruce por {modo}. {len(huerfanos)} valor(es) en 'ObjetivosVentas' NO "
                  f"encontrados en la dimensión de sucursales: {huerfanos}")

        resumen["presupuesto"] = resumen["presupuesto"].fillna(0)

    # Sólo se incluyen sucursales CON presupuesto asignado (> 0). El resto
    # (sin fila en 'ObjetivosVentas', en 0, o marcadas explícitamente como
    # sin presupuesto) se excluye por completo de la tabla.
    antes = len(resumen)
    sin_presupuesto = sorted(resumen.loc[resumen["presupuesto"] <= 0, "NombreSucursal"].tolist())
    resumen = resumen[resumen["presupuesto"] > 0].copy()
    if sin_presupuesto:
        print(f"ℹ️  [Presupuesto] {len(sin_presupuesto)} de {antes} sucursal(es) se excluyeron de la tabla "
              f"por no tener presupuesto asignado: {sin_presupuesto}")

    # ── Inicio efectivo, días operativos del mes y días transcurridos ──
    usa_inferencia = sorted(
        resumen.loc[resumen["FechaApertura"].isna(), "NombreSucursal"].tolist()
    )
    if usa_inferencia:
        print(f"ℹ️  [Presupuesto] {len(usa_inferencia)} sucursal(es) sin 'FechaApertura' registrada: se usó "
              f"el método de inferencia anterior (primera venta del mes) para calcular sus días: {usa_inferencia}")

    def _calcular_dias(row):
        apertura = row["FechaApertura"]
        if pd.isna(apertura):
            # Respaldo: inferir a partir de la primera venta del mes (método anterior).
            inicio = row["fechaMinDt"] if pd.notna(row["fechaMinDt"]) else primer_dia_mes
        else:
            inicio = apertura
        if inicio < primer_dia_mes:
            inicio = primer_dia_mes
        dias_operativos = max(0, (fin_mes - inicio).days + 1)
        dias_transcurridos = max(0, (fecha_max_global - inicio).days + 1)
        return pd.Series({"diasOperativosMes": dias_operativos, "diasTranscurridos": dias_transcurridos})

    resumen[["diasOperativosMes", "diasTranscurridos"]] = resumen.apply(_calcular_dias, axis=1)

    resumen["pronostico"] = np.where(
        resumen["diasTranscurridos"] > 0,
        resumen["ventasActual"] / resumen["diasTranscurridos"] * resumen["diasOperativosMes"],
        0.0,
    )
    resumen["cumplimiento"] = _margen_seguro(resumen["pronostico"], resumen["presupuesto"])

    for c in ["unidadesActual", "ventasActual", "utilidadActual", "pronostico", "presupuesto"]:
        resumen[c] = resumen[c].round(2)

    resumen = resumen.sort_values("ClaveSucursal").reset_index(drop=True)
    return resumen[["ClaveSucursal", "NombreSucursal", "unidadesActual", "ventasActual", "utilidadActual",
                     "margen", "presupuesto", "cumplimiento", "pronostico"]]


def generar_html(agg, linea_agg, historico_agg, top_art_agg, lineas_cat_agg, fabricantes_agg,
                 presupuesto_agg, lista_sucursales, fecha_reporte, fecha_info, mes_header):
    data_json         = json.dumps(agg.to_dict("records"),           ensure_ascii=False)
    linea_json        = json.dumps(linea_agg.to_dict("records"),     ensure_ascii=False)
    historico_json    = json.dumps(historico_agg.to_dict("records"), ensure_ascii=False)
    top_art_json      = json.dumps(top_art_agg.to_dict("records"),   ensure_ascii=False)
    lineas_cat_json   = json.dumps(lineas_cat_agg.to_dict("records"),ensure_ascii=False)
    fabricantes_json  = json.dumps(fabricantes_agg.to_dict("records"),ensure_ascii=False)
    presupuesto_json  = json.dumps(presupuesto_agg.to_dict("records"),ensure_ascii=False)
    sucursales_json   = json.dumps(lista_sucursales,                  ensure_ascii=False)

    periodos_unicos = (
        historico_agg[["Año","MesNum","PeriodoLabel"]]
        .drop_duplicates()
        .sort_values(["Año","MesNum"])["PeriodoLabel"]
        .tolist()
    )
    periodos_json = json.dumps(periodos_unicos, ensure_ascii=False)

    html_template = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Euphoria Skin · Análisis de Ventas</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-chart-treemap@3/dist/chartjs-chart-treemap.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
/* ── TEXTO FLUIDO EN PANTALLAS PEQUEÑAS ──
   Casi todo el CSS usa 'rem', que se calcula a partir de este tamaño base.
   En vez de reducir cada font-size individualmente, se escala el tamaño
   raíz de forma fluida: entre 320px y 400px de ancho el texto (títulos,
   tablas, botones, KPIs) se reduce proporcionalmente; arriba de ~400px
   se mantiene el tamaño normal (16px). */
html{font-size:clamp(13px, 4vw, 16px)}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#fcfbfe;color:#222126;min-height:100vh}
header{background:#7030A0;padding:1.2rem 2rem;display:flex;align-items:center;justify-content:space-between;box-shadow:0 3px 20px rgba(112,48,160,0.2)}
.header-left h1{font-size:1.15rem;font-weight:700;color:#fff;letter-spacing:.03em}
.header-left p{font-size:.72rem;color:#f3ecfa;margin-top:4px;letter-spacing:0.02em}
.header-right{display:flex;flex-direction:column;align-items:flex-end;gap:.35rem}
.hbadge{background:#ffffff1f;border:1px solid #ffffff40;color:#fff;font-size:.68rem;font-weight:600;padding:4px 14px;border-radius:20px;letter-spacing:.05em}
.hdate{font-size:.72rem;color:#f3ecfa;font-weight:500}
/* ── SEGMENTADORES ── */
.filter-bar{background:#fff;border-bottom:1px solid #edeaf2;padding:.75rem 2rem;display:flex;gap:.8rem;flex-wrap:wrap;align-items:center;justify-content:center;box-shadow:0 2px 4px rgba(112,48,160,0.02)}
.filter-bar + .filter-bar{border-top:1px solid #f3eff7;box-shadow:none}
.suc-label{font-size:.67rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#7030A0;margin-right:.25rem;white-space:nowrap}
.suc-btns{display:flex;gap:6px;flex-wrap:wrap;justify-content:center}
.suc-btn{display:inline-flex;align-items:center;padding:6px 13px;border-radius:20px;border:1px solid #e2daeb;background:#f9f6fc;color:#6d6575;font-size:.72rem;font-weight:600;cursor:pointer;transition:all .2s;font-family:inherit;white-space:nowrap}
.suc-btn.active{background:#7030A0 !important;border-color:#7030A0 !important;color:#fff !important;font-weight:700;box-shadow:0 3px 10px rgba(112,48,160,0.2)}
.mes-btn{display:inline-flex;align-items:center;padding:5px 11px;border-radius:20px;border:1px solid #e2daeb;background:#f9f6fc;color:#6d6575;font-size:.69rem;font-weight:600;cursor:pointer;transition:all .2s;font-family:inherit;white-space:nowrap}
.mes-btn.active{background:#494350 !important;border-color:#494350 !important;color:#fff !important;font-weight:700}
.suc-sep{width:1px;height:20px;background:#e3ddeb;margin:0 .1rem;flex-shrink:0}
.btn-all{padding:6px 13px;border-radius:20px;border:1px solid #7030A0;background:#fff;color:#7030A0;font-size:.7rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .2s;white-space:nowrap}
.btn-all.dark{border-color:#494350;color:#494350}
/* ── TABS ── */
.tabs-nav{background:#fff;border-bottom:2px solid #e9e5ed;padding:0 2rem;display:flex;justify-content:center;gap:0;flex-wrap:wrap}
.tab-nav-btn{padding:.85rem 1.6rem;font-size:.82rem;font-weight:700;color:#85808c;cursor:pointer;border:none;background:none;font-family:inherit;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .2s;letter-spacing:.02em}
.tab-nav-btn.active{color:#7030A0;border-bottom-color:#7030A0}
.tab-nav-btn:hover:not(.active){color:#494350;border-bottom-color:#e2daeb}
.tab-content{display:none}
.tab-content.active{display:block}
/* ── LAYOUT ── */
.main{padding:1.5rem 2rem 3rem;max-width:1450px;margin:0 auto}
.kpi-grid{display:flex;flex-wrap:wrap;justify-content:center;gap:.9rem;margin-bottom:1.5rem}
.kpi{background:#fff;border-radius:12px;padding:1.1rem 1.2rem;border:1px solid #e9e5ed;position:relative;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.015);flex:1 1 170px;max-width:215px;min-width:0;text-align:center}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--ac,#7030A0)}
.kpi-label{font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#85808c;margin-bottom:6px}
.kpi-value{font-size:1.4rem;font-weight:700;color:#222126;line-height:1}
.kpi-value.purple{color:#7030A0}
.kpi-value.gold{color:#aa7300}
.kpi-sub{font-size:.67rem;color:#85808c;margin-top:5px}
.kpi-note{font-size:.68rem;color:#85808c;margin:-.6rem 0 1rem;font-style:italic}
.tc{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e9e5ed;margin-bottom:.9rem;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:1.1rem;gap:1rem;flex-wrap:wrap}
.card-title{font-size:.88rem;font-weight:700;color:#222126;letter-spacing:0.01em}
.card-sub{font-size:.7rem;color:#85808c;margin-top:2px}
.note-bol{font-size:.67rem;background:#fbf9fc;color:#7030A0;border:1px solid #e3ddeb;padding:4px 10px;border-radius:6px;white-space:nowrap;font-weight:500}
.metric-tabs{display:flex;gap:4px;background:#f3eff5;padding:3px;border-radius:8px;border:1px solid #e6e1eb}
.tab-btn{background:none;border:none;padding:5px 12px;font-size:.72rem;font-weight:600;color:#6f6a75;cursor:pointer;border-radius:6px;font-family:inherit;transition:all .15s}
.tab-btn.active{background:#fff;color:#7030A0;box-shadow:0 2px 5px rgba(112,48,160,0.1);font-weight:700}
table{width:100%;border-collapse:collapse;font-size:.78rem;min-width:100%}
thead tr{background:#fdfcfd;border-bottom:2px solid #e9e5ed}
thead th{text-align:left;padding:9px 11px;font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#7030A0;white-space:nowrap}
thead th.r{text-align:right}
tbody tr{border-bottom:1px solid #f6f4f8;transition:background .15s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:#fdfbff}
tbody tr.total-row{background:linear-gradient(90deg,#fcfbfe,#fff);border-top:2px solid #beadd1}
tbody tr.total-row td{font-weight:700;color:#7030A0}
td{padding:9px 11px;color:#4f4c54}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.date{font-weight:600;color:#222126;white-space:nowrap}
td.dayname{font-size:.68rem;color:#85808c;white-space:nowrap}
td.art-code{font-family:'Consolas',monospace;font-size:.72rem;color:#7030A0;font-weight:700;white-space:nowrap}
/* CAMBIO: descripción completa sin corte, wrappea en múltiples líneas */
td.art-desc{font-size:.75rem;color:#4f4c54;line-height:1.4;word-break:break-word;min-width:200px}
td.art-fab{font-size:.72rem;color:#6d6575;word-break:break-word;min-width:120px}
td.rank{font-size:.78rem;font-weight:800;color:#85808c;width:32px;text-align:center}
td.rank.top3{color:#7030A0}
.pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:.65rem;font-weight:600;letter-spacing:0.02em}
.pill.hi{background:#e3f7ed;color:#176440}
.pill.mi{background:#fff3db;color:#805200}
.pill.lo{background:#fbe4e4;color:#a12727}
.charts-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:.9rem}
.cc{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e9e5ed;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.lines-container{display:grid;grid-template-columns:minmax(0,3fr) minmax(0,2fr);gap:.9rem;margin-top:1.5rem;align-items:start}
.lines-table-box,.lines-chart-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e9e5ed;box-shadow:0 2px 6px rgba(0,0,0,0.015);display:flex;flex-direction:column;min-width:0}
.hist-card{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e9e5ed;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.cw{position:relative;height:265px}
/* La altura de .bars-horizontal ahora la fija JS dinámicamente según el
   número de líneas mostradas (mismo patrón que fab-chart-wrap / lc-chart-wrap),
   en vez de estirarse por CSS para igualar la altura de la tabla vecina. */
.cw.bars-horizontal{min-height:220px}
.cw.hist-chart{position:relative;height:360px;width:100%}
/* CAMBIO: burbujas más altas para mejor proporción con ancho completo */
.cw.bubble-chart{position:relative;height:560px;width:100%}
.empty{text-align:center;padding:2.5rem;color:#85808c;font-size:.82rem}
.section-divider{border:none;border-top:2px solid #ede9f3;margin:2rem 0}
/* ── TOP ARTÍCULOS ── */
.top-header-card{background:#fff;border-radius:12px;padding:1rem 1.6rem;border:1px solid #e9e5ed;box-shadow:0 2px 6px rgba(0,0,0,0.015);margin-bottom:1rem}
.top-layout{display:block}
.top-table-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e9e5ed;box-shadow:0 2px 6px rgba(0,0,0,0.015);margin-bottom:1rem}
.top-bubble-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e9e5ed;box-shadow:0 2px 6px rgba(0,0,0,0.015);display:flex;flex-direction:column;min-height:0}
.badge-rank{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-size:.68rem;font-weight:800;background:#f3eff7;color:#85808c}
.badge-rank.r1{background:#7030A0;color:#fff}
.badge-rank.r2{background:#9052B8;color:#fff}
.badge-rank.r3{background:#A97DD1;color:#fff}
.top-kpi-row{display:grid;grid-template-columns:repeat(3,1fr);gap:.75rem;margin-bottom:1rem}
.top-kpi{background:#fff;border-radius:10px;padding:.9rem 1.1rem;border:1px solid #e9e5ed;position:relative;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.015);text-align:center}
.top-kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--ac,#7030A0)}
.top-kpi .kpi-label{font-size:.6rem}
.top-kpi .kpi-value{font-size:1.1rem}
/* ── LÍNEAS Y CATEGORÍAS (treemap) ── */
.legend-gradient{display:inline-flex;align-items:center;gap:6px;font-size:.67rem;color:#6d6575;white-space:nowrap}
.legend-bar{width:70px;height:8px;border-radius:4px;background:linear-gradient(90deg,#e2d3f0,#4c1a7a)}
tr.lc-parent{cursor:pointer}
tr.lc-parent:hover{background:#fdfbff}
tr.lc-child td{background:#fbfaFc;color:#6d6575}
tr.lc-child td:nth-child(2){padding-left:26px}
.lc-icon{display:inline-block;width:14px;color:#7030A0;font-weight:700}
/* ── FABRICANTES ── */
.fq-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.65rem;font-weight:700;letter-spacing:.02em;white-space:nowrap}
tr.fq-parent{cursor:pointer}
tr.fq-parent:hover td{filter:brightness(0.97)}
tr.fq-parent td{padding-top:10px;padding-bottom:10px}
.fab-bar-cell{position:relative}
.fab-bar-bg{position:absolute;right:0;top:3px;bottom:3px;border-radius:4px;z-index:0}
.fab-bar-cell b{position:relative;z-index:1}
.fab-rank-num{font-size:.68rem;color:#85808c;font-weight:700}
/* ── TABLAS RESPONSIVAS: auto-escala para caber en el ancho de pantalla ──
   En vez de reformatear filas en tarjetas verticales (lo cual alarga la
   página), la tabla completa se encoge proporcionalmente (JS calcula el
   factor de escala) manteniendo su formato tabular normal. */
.table-scale-wrap{
  width:100%;
  overflow:hidden;
}
.table-scale-wrap table{
  transform-origin: top left;
}
/* ── RESPONSIVE: .charts-row ya se resuelve solo vía auto-fit/minmax arriba.
   .lines-container mantiene proporción asimétrica (3fr/2fr) en pantallas
   anchas, así que sigue necesitando un punto de quiebre explícito para
   apilarse en pantallas angostas. */
@media (max-width: 768px){
  .lines-container{grid-template-columns:1fr}
}
/* ── KPIs EN PANTALLAS PEQUEÑAS: 3 POR FILA ──
   Antes se apilaban de una en una, lo que alargaba mucho la página (mucho
   scroll) y desperdiciaba el ancho disponible. Ahora van de tres en tres.
   Los tamaños de fuente usan clamp() con unidades vw para que el texto
   escale de forma continua con el ancho real de la pantalla (a 320px se ve
   más pequeño que a 600px), en vez de quedar fijo y desbordar la tarjeta.
   Las reglas :nth-child evitan que la última fila quede coja: si sobra una
   tarjeta ocupa el ancho completo, y si sobran dos se reparten a la mitad. */
@media (max-width: 640px){
  .main{padding:1.2rem 1rem 2.5rem}
  .kpi-grid{gap:.4rem}
  .kpi{flex:1 1 calc(33.333% - .4rem);max-width:none;padding:.7rem .4rem;min-width:0}
  .kpi:last-child:nth-child(3n+1){flex-basis:100%}
  .kpi:nth-last-child(2):nth-child(3n+1),
  .kpi:last-child:nth-child(3n+2){flex-basis:calc(50% - .4rem)}
  .kpi-value{font-size:clamp(.72rem, 4.1vw, 1.15rem);overflow-wrap:anywhere}
  .kpi-label{font-size:clamp(.42rem, 2.2vw, .58rem);letter-spacing:.03em;line-height:1.2;margin-bottom:4px;overflow-wrap:break-word;hyphens:auto}
  .kpi-sub{font-size:clamp(.4rem, 2vw, .58rem);line-height:1.2;margin-top:3px;overflow-wrap:break-word;hyphens:auto}
  /* Mismo criterio para las tarjetas de las demás pestañas.
     Se usa !important porque algunas instancias fijan sus columnas por
     estilo inline (ej. "Ventas vs. Objetivos" usa repeat(4,1fr) inline),
     y un inline style sólo puede sobreescribirse con !important. */
  .top-kpi-row{grid-template-columns:repeat(3,1fr) !important;gap:.4rem}
  .top-kpi:last-child:nth-child(3n+1){grid-column:1 / -1}
  .top-kpi:nth-last-child(2):nth-child(3n+1){grid-column:span 2}
  .top-kpi{padding:.65rem .4rem;min-width:0}
  .top-kpi .kpi-value{font-size:clamp(.68rem, 3.4vw, .95rem);overflow-wrap:anywhere}
  .top-kpi .kpi-label{font-size:clamp(.4rem, 2.1vw, .55rem);line-height:1.2}
  /* Las tarjetas "Línea Líder" y "Categoría Líder" fijan su tamaño por
     estilo inline, así que necesitan !important para reducirse también. */
  .top-kpi .kpi-value.kv-txt{font-size:clamp(.6rem, 2.9vw, .85rem) !important}
  /* ── SEGMENTADORES COMPACTOS ──
     .suc-btns es el contenedor tanto de los botones de Sucursales como de
     Meses (comparten la misma clase). Se reducen padding, gap y tamaño de
     fuente al mínimo legible para que cuantos más botones sea posible
     quepan por fila, sin ocultar ninguno ni usar scroll interno: todos
     siguen visibles, sólo más compactos. */
  .filter-bar{padding:.5rem .75rem;gap:.4rem}
  .suc-label{font-size:.6rem;margin-right:0}
  .suc-btns{gap:4px}
  .suc-btn,.mes-btn{padding:3px 8px;font-size:.6rem}
  .btn-all{padding:4px 10px;font-size:.6rem}
  .suc-sep{height:14px}
  .tab-nav-btn{padding:.7rem .9rem;font-size:.74rem}
}
</style>
</head>
<body>

<header>
  <div class="header-left">
    <h1>Euphoria Skin — Análisis de Ventas</h1>
    <p><strong>Período __MES_HEADER__</strong></p>
    <p>Elaborado con información al día __FECHA_INFO__</p>
  </div>
  <div class="header-right">
    <span class="hdate">Reporte Generado: __FECHA_VALOR__</span>
    <div class="hbadge">DASHBOARD v1.2</div>
  </div>
</header>

<div class="filter-bar" id="filter-bar-suc">
  <span class="suc-label">Sucursales</span>
  <div class="suc-btns" id="suc-btns"></div>
  <div class="suc-sep"></div>
  <button class="btn-all" onclick="toggleAll()">Todas / Ninguna</button>
</div>

<div class="filter-bar" id="filter-bar-mes">
  <span class="suc-label" style="color:#494350">Meses · Histórico</span>
  <div class="suc-btns" id="mes-btns"></div>
  <div class="suc-sep"></div>
  <button class="btn-all dark" onclick="toggleAllMeses()">Todos / Ninguno</button>
</div>

<div class="tabs-nav">
  <button class="tab-nav-btn active" id="tabnav-objetivos"    onclick="switchTab('objetivos')">Ventas vs. Objetivos</button>
  <button class="tab-nav-btn"        id="tabnav-resumen"      onclick="switchTab('resumen')">Resumen</button>
  <button class="tab-nav-btn"        id="tabnav-toparticulos" onclick="switchTab('toparticulos')">Top 10 Artículos</button>
  <button class="tab-nav-btn"        id="tabnav-lineascategoria" onclick="switchTab('lineascategoria')">Líneas y Categorías</button>
  <button class="tab-nav-btn"        id="tabnav-fabricantes" onclick="switchTab('fabricantes')">Fabricantes</button>
</div>

<div id="tab-objetivos" class="tab-content active">
<div class="main">

  <div class="tc" style="margin-bottom:.9rem">
    <div class="card-head">
      <div><div class="card-title">Ventas vs. Objetivos</div><div class="card-sub">Ventas $ mes en curso</div></div>
      <span class="note-bol">Pronóstico de Alcance = Ventas estimadas al cierre de mes</span>
    </div>
    <div class="top-kpi-row" style="grid-template-columns:repeat(4,1fr)">
      <div class="top-kpi" style="--ac:#494350"><div class="kpi-label">Ventas $</div><div class="kpi-value" id="p-ventas">—</div><div class="kpi-sub">Acumulado mes en curso</div></div>
      <div class="top-kpi" style="--ac:#7030A0"><div class="kpi-label">Pronóstico de cierre</div><div class="kpi-value purple" id="p-pronostico">—</div><div class="kpi-sub">Suma de pronósticos</div></div>
      <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Presupuesto</div><div class="kpi-value gold" id="p-presupuesto">—</div><div class="kpi-sub">Acumulado</div></div>
      <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Alcance actual</div><div class="kpi-value" id="p-cumplimiento" style="color:#1A7A4A">—</div><div class="kpi-sub">Ventas actuales ÷ Presupuesto</div></div>
    </div>
    <div class="table-scale-wrap">
      <table>
        <thead><tr><th>Sucursal</th><th class="r">Unidades</th><th class="r">Venta $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Presupuesto</th><th class="r">Alcance actual</th><th class="r">Pronóstico de Alcance</th></tr></thead>
        <tbody id="tabla-presupuesto"></tbody>
      </table>
    </div>
  </div>

  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Ventas vs. Objetivo</div><div class="card-sub">Comparativo por sucursal</div></div>
    </div>
    <div class="cw" id="obj-chart-wrap" style="height:420px"><canvas id="chart-objetivos"></canvas></div>
  </div>

</div>
</div>

<div id="tab-resumen" class="tab-content">
<div class="main">

  <p class="kpi-note">* Las tarjetas que se muestran se calculan con datos de las Ventas $ del mes en curso.</p>

  <div class="kpi-grid">
    <div class="kpi" style="--ac:#8A62AD"><div class="kpi-label">Unidades vendidas</div><div class="kpi-value" id="k-uni">—</div><div class="kpi-sub">Ventas mes en curso</div></div>
    <div class="kpi" style="--ac:#7030A0"><div class="kpi-label">Ventas $</div><div class="kpi-value purple" id="k-ventas">—</div><div class="kpi-sub" id="k-ventas-s">—</div></div>
    <div class="kpi" style="--ac:#541e82"><div class="kpi-label">Utilidad</div><div class="kpi-value purple" id="k-util">—</div><div class="kpi-sub" id="k-util-sub">Margen: —</div></div>
    <div class="kpi" style="--ac:#494350"><div class="kpi-label">Tickets</div><div class="kpi-value" id="k-tkt">—</div><div class="kpi-sub">Volumen de ventas</div></div>
    <div class="kpi" style="--ac:#937ca8"><div class="kpi-label">Ventas $ promedio por ticket</div><div class="kpi-value" id="k-vtkt">—</div><div class="kpi-sub">Ventas $ ÷ Tickets</div></div>
    <div class="kpi" style="--ac:#baa3d4"><div class="kpi-label">Unidades promedio por Ticket</div><div class="kpi-value" id="k-utkt">—</div><div class="kpi-sub">Unidades ÷ Tickets</div></div>
  </div>

  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Resumen de Ventas · Mes en curso</div><div class="card-sub">Acumulado interactivo según selección</div></div>
      <span class="note-bol">Desempeño de ventas por día</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Fecha</th><th>Día</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-body"></tbody>
    </table>
    </div>
  </div>

  <div class="charts-row">
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">Ventas $</div><div class="card-sub">Volumen diario · Sucursales seleccionadas</div></div></div><div class="cw"><canvas id="chart-ventas"></canvas></div></div>
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">No. de Tickets</div><div class="card-sub">Volumen diario · Sucursales seleccionadas</div></div></div><div class="cw"><canvas id="chart-tickets"></canvas></div></div>
  </div>

  <hr class="section-divider">

  <div class="tc" style="margin-bottom:.9rem">
    <div class="card-head">
      <div><div class="card-title">Ventas Históricas Mensuales</div><div class="card-sub">Acumulado por mes · Sucursales y meses seleccionados</div></div>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Período</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-historico"></tbody>
    </table>
    </div>
  </div>

  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Tendencia Histórica Mensual</div><div class="card-sub">Ventas mensuales y Margen · Sucursales y meses seleccionados</div></div>
      <div class="metric-tabs">
        <button class="tab-btn active" id="btn-h-ventas"   onclick="changeHistMetric('ventas')">Ventas c/Desc</button>
        <button class="tab-btn"        id="btn-h-utilidad" onclick="changeHistMetric('utilidad')">Utilidad</button>
        <button class="tab-btn"        id="btn-h-unidades" onclick="changeHistMetric('unidades')">Unidades</button>
        <button class="tab-btn"        id="btn-h-tickets"  onclick="changeHistMetric('tickets')">Tickets</button>
      </div>
    </div>
    <div class="cw hist-chart"><canvas id="chart-historico"></canvas></div>
  </div>

  <hr class="section-divider">

  <div class="lines-container">
    <div class="lines-table-box">
      <div class="card-head">
        <div><div class="card-title">Resumen de Ventas Acumuladas por Línea</div><div class="card-sub">Acumulado de ventas según selección</div></div>
      </div>
      <div class="table-scale-wrap">
        <table>
          <thead><tr><th>Línea</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th></tr></thead>
          <tbody id="tabla-lineas"></tbody>
        </table>
      </div>
    </div>
    <div class="lines-chart-box">
      <div class="card-head" style="margin-bottom:.7rem">
        <div><div class="card-title">Porcentaje de Participación (%)</div><div class="card-sub">Proporción porcentual de acuerdo a Ventas, Unidades o Utilidad</div></div>
        <div class="metric-tabs">
          <button class="tab-btn active" id="btn-m-ventas"   onclick="changeLineMetric('ventas')">Ventas c/Desc</button>
          <button class="tab-btn"        id="btn-m-unidades" onclick="changeLineMetric('unidades')">Unidades</button>
          <button class="tab-btn"        id="btn-m-utilidad" onclick="changeLineMetric('utilidad')">Utilidad</button>
        </div>
      </div>
      <div class="cw bars-horizontal" id="lineas-chart-wrap"><canvas id="chart-lineas"></canvas></div>
    </div>
  </div>

</div>
</div>

<div id="tab-toparticulos" class="tab-content">
<div class="main">

  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#2E5B88"><div class="kpi-label">Unidades vendidas Top 10</div><div class="kpi-value" id="tk-uni">—</div><div class="kpi-sub">Suma de top 10</div></div>
    <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Ventas $ Top 10</div><div class="kpi-value gold" id="tk-ventas">—</div><div class="kpi-sub">Suma de top 10</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">% participación sobre ventas $</div><div class="kpi-value" id="tk-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Ventas $ Top 10 ÷ Ventas $ acumuladas</div></div>
  </div>

  <div class="top-layout">
    <div class="top-table-box">
      <div class="card-head">
        <div><div class="card-title">Ranking de Artículos</div><div class="card-sub">Top 10 por Unidades Vendidas</div></div>
        <span class="note-bol">Acorde a selección</span>
      </div>
      <div class="table-scale-wrap">
      <table>
        <thead>
          <tr>
            <th style="width:36px;text-align:center">#</th>
            <th>Artículo</th>
            <th>Descripción</th>
            <th>Fabricante</th>
            <th class="r">Unidades</th>
            <th class="r">Venta $</th>
            <th class="r">% Venta</th>
          </tr>
        </thead>
        <tbody id="tabla-top"></tbody>
      </table>
      </div>
    </div>

    <div class="top-bubble-box">
      <div class="card-head" style="margin-bottom:.5rem">
        <div>
          <div class="card-title">Mapa de Desempeño · Top 10</div>
          <div class="card-sub">Eje X: Ventas $ · Eje Y: Margen (%) · Tamaño burbuja: Unidades vendidas</div>
        </div>
      </div>
      <div class="cw bubble-chart"><canvas id="chart-bubble"></canvas></div>
    </div>
  </div>

</div>
</div>

<div id="tab-lineascategoria" class="tab-content">
<div class="main">

  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#7030A0"><div class="kpi-label">Línea Líder</div><div class="kpi-value purple kv-txt" id="tk2-linea-nombre" style="font-size:1.05rem">—</div><div class="kpi-sub" id="tk2-linea-sub">—</div></div>
    <div class="top-kpi" style="--ac:#5C2485"><div class="kpi-label">Categoría Líder</div><div class="kpi-value kv-txt" id="tk2-cat-nombre" style="color:#5C2485;font-size:1.05rem">—</div><div class="kpi-sub" id="tk2-cat-sub">—</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Margen · Categoría Líder</div><div class="kpi-value" id="tk2-cat-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Utilidad ÷ Ventas de la categoría líder</div></div>
  </div>

  <div class="top-table-box">
    <div class="card-head">
      <div><div class="card-title">Detalle por Línea y Categoría</div><div class="card-sub">Clic en una línea para desplegar sus categorías a detalle</div></div>
      <span class="note-bol">Acorde a sucursales y meses seleccionados</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead>
        <tr>
          <th style="width:26px"></th>
          <th>Línea / Categoría</th>
          <th class="r">Unidades</th>
          <th class="r">Ventas $</th>
          <th class="r">Utilidad</th>
          <th class="r">Margen</th>
        </tr>
      </thead>
      <tbody id="tabla-lineascategoria"></tbody>
    </table>
    </div>
  </div>

  <div class="hist-card" style="margin-top:1rem">
    <div class="card-head">
      <div><div class="card-title">Mapa de Línea &gt; Categoría</div><div class="card-sub">Top 3 categorías por línea </div></div>
      <span class="legend-gradient">Menor venta<span class="legend-bar"></span>Mayor venta</span>
    </div>
    <div class="cw" id="lc-chart-wrap" style="height:360px"><canvas id="chart-treemap"></canvas></div>
  </div>

</div>
</div>

<div id="tab-fabricantes" class="tab-content">
<div class="main">

  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#7030A0"><div class="kpi-label">Ventas $ · Fabricantes al 50%</div><div class="kpi-value purple" id="tk3-ventas">—</div><div class="kpi-sub">Suma de fabricantes hasta cubrir el 50%</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Margen · Fabricantes al 50%</div><div class="kpi-value" id="tk3-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Utilidad ÷ Ventas del grupo</div></div>
    <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Fabricantes incluidos</div><div class="kpi-value gold" id="tk3-pct">—</div><div class="kpi-sub" id="tk3-pct-sub">Necesarios para alcanzar el 50% de ventas</div></div>
  </div>

  <div class="hist-card" style="margin-bottom:1rem">
    <div class="card-head">
      <div><div class="card-title">Fabricantes que concentran el 50% de las Ventas</div><div class="card-sub">% de participación en Ventas $</div></div>
    </div>
    <div class="cw" id="fab-chart-wrap" style="height:420px"><canvas id="chart-fabricantes"></canvas></div>
  </div>

  <div class="top-table-box">
    <div class="card-head">
      <div><div class="card-title">Detalle Completo por Fabricante</div><div class="card-sub">Todos los fabricantes de la selección, agrupados en cuartiles según su venta acumulada · clic en un cuartil para desplegar el detalle</div></div>
      <span class="note-bol">Acorde a sucursales y meses seleccionados</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead>
        <tr>
          <th style="width:26px"></th>
          <th>Fabricante</th>
          <th class="r">Ventas $</th>
          <th class="r">Margen</th>
          <th class="r">% Participación</th>
        </tr>
      </thead>
      <tbody id="tabla-fabricantes"></tbody>
    </table>
    </div>
  </div>

</div>
</div>

<footer style="text-align:center;padding:1.4rem 2rem 2rem;font-size:.7rem;color:#a39cad;border-top:1px solid #edeaf2;margin-top:1rem">
  Elaborado por el Equipo de Planeación y Análisis de la información.
</footer>

<script>
document.addEventListener("DOMContentLoaded", function() {

    const RAW         = __DATA_JSON__;
    const LINEAS      = __LINEA_JSON__;
    const HISTORICO   = __HISTORICO_JSON__;
    const TOP_ART     = __TOP_ART_JSON__;
    const LINEAS_CAT  = __LINEAS_CAT_JSON__;
    const FABRICANTES = __FABRICANTES_JSON__;
    const PRESUPUESTO = __PRESUPUESTO_JSON__;
    const SUCS        = __SUCURSALES_JSON__;
    const PERIODOS    = __PERIODOS_JSON__;

    const MESES_ABR = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

    const PALETTE = [
        '#7030A0','#2196A8','#E05C2A','#3A7D44','#C4922A',
        '#1A559E','#B03060','#4E8C6E','#6B4226','#5C5FA8',
        '#8E3A59','#2E7D60','#A04010','#3D6B99','#7A6F1E',
        '#9C3D6B','#1E6B7A','#5A3A8E','#3D7A4A','#8E5A1E',
    ];

    const PALETTE_LINEAS = [
        '#5C2485','#7030A0','#8340B8','#9652CC','#A866E0',
        '#685A75','#827491','#9D90AD','#3B1A60','#B07FD4'
    ];

    const PALETTE_BUBBLE = [
        '#7030A0','#2196A8','#E05C2A','#3A7D44','#C4922A',
        '#1A559E','#B03060','#4E8C6E','#6B4226','#5C5FA8'
    ];

    function buildColorMap(list, palette) {
        const map = {};
        list.forEach((name, i) => { map[name] = palette[i % palette.length]; });
        return map;
    }

    const ALL_LINEAS = [...new Set(LINEAS.map(r => r.Línea))].sort();
    const LC = buildColorMap(ALL_LINEAS, PALETTE_LINEAS);
    let SC = {};

    const DAYS = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];

    let active      = new Set(SUCS);
    let activeMeses = new Set(PERIODOS);
    let charts      = {};
    let currentLineMetric = 'ventas';
    let currentHistMetric = 'ventas';
    // La pestaña inicial es la de Objetivos (es la primera de la navegación).
    let currentTab = 'objetivos';

    const fM  = v => '$'+(v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':Math.round(v));
    const fF  = v => '$'+Math.round(v).toLocaleString('es-MX');
    const fP  = v => (v*100).toFixed(2)+'%';
    const fN  = v => Math.round(v).toLocaleString('es-MX');

    const PREMIUM_TOOLTIP_OPTS = {
        backgroundColor:'#ffffff',titleColor:'#222126',bodyColor:'#4f4c54',
        borderColor:'#e9e5ed',borderWidth:1,padding:10,cornerRadius:8,
        boxPadding:6,usePointStyle:true,
        titleFont:{family:"'Segoe UI', sans-serif",weight:'bold',size:12},
        bodyFont:{family:"'Segoe UI', sans-serif",size:12}
    };

    Chart.register(ChartDataLabels);

    // ── TABLAS RESPONSIVAS: encoge cada tabla (sin scroll lateral) para que
    // quepa siempre en el ancho disponible de su contenedor. Se recalcula
    // cada vez que una tabla se reconstruye (ver llamadas a fitTables() al
    // final de cada función update*) y también al redimensionar la ventana.
    function fitTables(){
        document.querySelectorAll('.table-scale-wrap').forEach(wrap => {
            const table = wrap.querySelector('table');
            if(!table) return;
            table.style.transform = 'none';
            wrap.style.height = 'auto';
            const availWidth  = wrap.clientWidth;
            const neededWidth = table.scrollWidth;
            if(availWidth > 0 && neededWidth > availWidth){
                const scale = availWidth / neededWidth;
                table.style.transform = `scale(${scale})`;
                wrap.style.height = (table.offsetHeight * scale) + 'px';
            }
        });
    }
    window.addEventListener('resize', () => {
        clearTimeout(window._fitTablesTimer);
        window._fitTablesTimer = setTimeout(fitTables, 150);
    });

    window.switchTab = function(tab) {
        currentTab = tab;
        document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.tab-nav-btn').forEach(el => el.classList.remove('active'));
        document.getElementById('tab-' + tab).classList.add('active');
        document.getElementById('tabnav-' + tab).classList.add('active');

        // La pestaña de Objetivos está desvinculada de los segmentadores, pero se
        // decidió mantener las barras de filtros SIEMPRE visibles. Para volver a
        // ocultarlas ahí, cambiar la línea siguiente por:
        // const ocultarFiltros = (tab === 'objetivos');
        const ocultarFiltros = false;
        document.getElementById('filter-bar-suc').style.display = ocultarFiltros ? 'none' : 'flex';
        document.getElementById('filter-bar-mes').style.display = ocultarFiltros ? 'none' : 'flex';

        if (tab === 'resumen') update();
        if (tab === 'objetivos') updateObjetivos();
        if (tab === 'toparticulos') updateTopArticulos();
        if (tab === 'lineascategoria') updateLineasCategoria();
        if (tab === 'fabricantes') updateFabricantes();
        setTimeout(fitTables, 50);
    };

    function buildButtons(){
        SC = buildColorMap(SUCS, PALETTE);
        const wrap = document.getElementById('suc-btns');
        wrap.innerHTML = '';
        SUCS.forEach(s => {
            const btn = document.createElement('button');
            btn.className = 'suc-btn active';
            btn.innerHTML = `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${SC[s]};margin-right:6px;flex-shrink:0"></span>${s}`;
            btn.addEventListener('click', () => {
                if(active.has(s)){ active.delete(s); btn.classList.remove('active'); }
                else { active.add(s); btn.classList.add('active'); }
                updateAll();
            });
            wrap.appendChild(btn);
        });
    }

    function buildMesButtons(){
        const wrap = document.getElementById('mes-btns');
        wrap.innerHTML = '';
        PERIODOS.forEach(p => {
            const btn = document.createElement('button');
            btn.className = 'mes-btn active';
            btn.textContent = p;
            btn.dataset.periodo = p;
            btn.addEventListener('click', () => {
                if(activeMeses.has(p)){ activeMeses.delete(p); btn.classList.remove('active'); }
                else { activeMeses.add(p); btn.classList.add('active'); }
                if(currentTab === 'resumen') { updateHistorico(); updateLineas(); }
                if(currentTab === 'toparticulos') updateTopArticulos();
                if(currentTab === 'lineascategoria') updateLineasCategoria();
                if(currentTab === 'fabricantes') updateFabricantes();
            });
            wrap.appendChild(btn);
        });
    }

    window.toggleAll = function(){
        if(active.size === SUCS.length){
            active.clear();
            document.querySelectorAll('.suc-btn').forEach(b => b.classList.remove('active'));
        } else {
            SUCS.forEach(s => active.add(s));
            document.querySelectorAll('.suc-btn').forEach(b => b.classList.add('active'));
        }
        updateAll();
    };

    window.toggleAllMeses = function(){
        if(activeMeses.size === PERIODOS.length){
            activeMeses.clear();
            document.querySelectorAll('.mes-btn').forEach(b => b.classList.remove('active'));
        } else {
            PERIODOS.forEach(p => activeMeses.add(p));
            document.querySelectorAll('.mes-btn').forEach(b => b.classList.add('active'));
        }
        if(currentTab === 'resumen') { updateHistorico(); updateLineas(); }
        if(currentTab === 'toparticulos') updateTopArticulos();
        if(currentTab === 'lineascategoria') updateLineasCategoria();
        if(currentTab === 'fabricantes') updateFabricantes();
    };

    // Nota: la pestaña 'objetivos' NO se incluye aquí a propósito: es
    // independiente de los segmentadores, por lo que no necesita recalcularse
    // cuando cambia la selección de sucursales o meses.
    function updateAll(){
        if(currentTab === 'resumen') update();
        if(currentTab === 'toparticulos') updateTopArticulos();
        if(currentTab === 'resumen') updateHistorico();
        if(currentTab === 'lineascategoria') updateLineasCategoria();
        if(currentTab === 'fabricantes') updateFabricantes();
    }

    function getFiltered()          { return RAW.filter(r => active.has(r.NombreSucursal)); }
    function getFilteredLineas()    {
        return LINEAS.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredHist()      {
        return HISTORICO.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredTopArt()    {
        return TOP_ART.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredLineasCat() {
        return LINEAS_CAT.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredFabricantes() {
        return FABRICANTES.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }

    function salesColor(t){
        // Interpola de lavanda claro (venta baja) a morado profundo de marca (venta alta). t en fracción 0-1.
        const clamped = Math.max(0, Math.min(1, t));
        const c1 = [222,208,238], c2 = [76,26,122];
        const r = Math.round(c1[0] + (c2[0]-c1[0])*clamped);
        const g = Math.round(c1[1] + (c2[1]-c1[1])*clamped);
        const b = Math.round(c1[2] + (c2[2]-c1[2])*clamped);
        return `rgba(${r},${g},${b},0.92)`;
    }

    function getDates()      { return [...new Set(RAW.map(r => r.FechaStr))].sort(); }
    function getActiveSucs() { return [...active].sort(); }

    function aggByDate(data){
        const m = {};
        data.forEach(r => {
            if(!m[r.FechaStr]) m[r.FechaStr] = {unidades:0,ventas:0,utilidad:0,tickets:0};
            m[r.FechaStr].unidades  += r.unidades;
            m[r.FechaStr].ventas    += r.ventas;
            m[r.FechaStr].utilidad  += r.utilidad;
            m[r.FechaStr].tickets   += r.tickets;
        });
        return m;
    }

    function dc(id){ if(charts[id]){ charts[id].destroy(); delete charts[id]; } }

    window.changeLineMetric = function(metric){
        currentLineMetric = metric;
        document.querySelectorAll('[id^="btn-m-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-m-' + metric).classList.add('active');
        updateLineas();
    };

    window.changeHistMetric = function(metric){
        currentHistMetric = metric;
        document.querySelectorAll('[id^="btn-h-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-h-' + metric).classList.add('active');
        updateHistorico();
    };

    // ── PESTAÑA VENTAS VS. OBJETIVOS ──
    // Esta pestaña usa SIEMPRE el dataset completo de PRESUPUESTO (todas las
    // sucursales con objetivo asignado) y siempre el mes en curso, sin aplicar
    // los segmentadores de Sucursales ni de Meses.
    function objetivosData(){
        return PRESUPUESTO.slice().sort((a,b) => a.ClaveSucursal - b.ClaveSucursal);
    }

    function updateObjetivos(){
        updatePresupuesto();
        updateChartObjetivos();
        fitTables();
    }

    function updatePresupuesto(){
        const data  = objetivosData();
        const tbody = document.getElementById('tabla-presupuesto');

        if(!data.length){
            tbody.innerHTML = '<tr><td colspan="8" class="empty">No hay sucursales con objetivo de ventas asignado en la hoja \'ObjetivosVentas\'.</td></tr>';
            ['p-ventas','p-pronostico','p-presupuesto','p-cumplimiento'].forEach(id => { document.getElementById(id).textContent = '—'; });
            return;
        }

        let tU=0, tV=0, tUt=0, tP=0, tB=0;

        const rows = data.map(r => {
            tU += r.unidadesActual; tV += r.ventasActual; tUt += r.utilidadActual;
            tP += r.pronostico; tB += r.presupuesto;
            const cump = r.ventasActual/r.presupuesto;
            const cls   = cump>=0.98 ? 'hi' : cump>=0.90 ? 'mi' : 'lo';
            const mgCls = r.margen>=0.50 ? 'hi' : 'mi';
            return `<tr>
                <td data-label="Sucursal"><b>${r.NombreSucursal}</b></td>
                <td class="r" data-label="Unidades">${fN(r.unidadesActual)}</td>
                <td class="r" data-label="Venta $">${fF(r.ventasActual)}</td>
                <td class="r" data-label="Utilidad">${fF(r.utilidadActual)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mgCls}">${fP(r.margen)}</span></td>
                <td class="r" data-label="Presupuesto">${fF(r.presupuesto)}</td>
                <td class="r" data-label="Cumplimiento"><span class="pill ${cls}">${fP(cump)}</span></td>
                <td class="r" data-label="Pronóstico de Alcance"><b>${fF(r.pronostico)}</b></td>
            </tr>`;
        }).join('');

        const totMg   = tV>0 ? tUt/tV : 0;
        const totCump = tB>0 ? tV/tB : 0;
        const totCls  = totCump>=0.98 ? 'hi' : totCump>=0.90 ? 'mi' : 'lo';

        tbody.innerHTML = rows + `<tr class="total-row">
            <td data-label=""><b>TOTAL GENERAL</b></td>
            <td class="r" data-label="Unidades">${fN(tU)}</td>
            <td class="r" data-label="Venta $">${fF(tV)}</td>
            <td class="r" data-label="Utilidad">${fF(tUt)}</td>
            <td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td>
            <td class="r" data-label="Presupuesto">${fF(tB)}</td>
            <td class="r" data-label="Cumplimiento"><span class="pill ${totCls}">${fP(totCump)}</span></td>
            <td class="r" data-label="Pronóstico de Alcance"><b>${fF(tP)}</b></td>
        </tr>`;

        document.getElementById('p-ventas').textContent       = fF(tV);
        document.getElementById('p-pronostico').textContent   = fF(tP);
        document.getElementById('p-presupuesto').textContent  = fF(tB);
        document.getElementById('p-cumplimiento').textContent = fP(totCump);
    }

    function updateChartObjetivos(){
        dc('objetivos');
        const data = objetivosData();
        if(!data.length) return;

        // Altura dinámica: con muchas sucursales las etiquetas del eje X
        // necesitan más espacio vertical al rotarse.
        document.getElementById('obj-chart-wrap').style.height =
            Math.max(360, Math.min(620, 340 + data.length * 6)) + 'px';

        const maxVal = Math.max(
            ...data.map(r => Math.max(r.ventasActual, r.presupuesto))
        ) || 0;

        charts['objetivos'] = new Chart(document.getElementById('chart-objetivos'), {
            type: 'bar',
            data: {
                labels: data.map(r => r.NombreSucursal),
                datasets: [
                    {
                        label: 'Ventas $ mes en curso',
                        data: data.map(r => r.ventasActual),
                        backgroundColor: '#7030A0cc',
                        borderColor: '#7030A0',
                        borderWidth: 1,
                        borderRadius: 4,
                    },
                    {
                        label: 'Presupuesto',
                        data: data.map(r => r.presupuesto),
                        backgroundColor: '#aa7300aa',
                        borderColor: '#aa7300',
                        borderWidth: 1,
                        borderRadius: 4,
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: {
                        display: true, position: 'top', onClick: null,
                        labels: {
                            boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'circle',
                            padding: 12, color: '#615e66',
                            font: { family: "'Segoe UI', sans-serif", size: 10, weight: '600' }
                        }
                    },
                    datalabels: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${fF(ctx.raw)}`,
                            // Línea extra con el avance real (ventas acumuladas ÷ objetivo),
                            // que es distinto del cumplimiento proyectado de la tabla.
                            footer: items => {
                                if(!items.length) return '';
                                const r = data[items[0].dataIndex];
                                const avance = r.presupuesto > 0 ? r.ventasActual / r.presupuesto : 0;
                                return `Avance real: ${fP(avance)}`;
                            }
                        },
                        footerColor: '#7030A0',
                        footerFont: { family: "'Segoe UI', sans-serif", size: 11, weight: 'bold' }
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 9 }, color: '#85808c', maxRotation: 60, minRotation: 45, autoSkip: false }
                    },
                    y: {
                        grid: { color: '#f6f4f8' },
                        suggestedMax: maxVal * 1.1,
                        ticks: { font: { size: 9 }, color: '#85808c', callback: v => fM(v) }
                    }
                }
            }
        });
    }

    function update(){
        const data       = getFiltered();
        const dates      = getDates();
        const byDate     = aggByDate(data);
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;

        const totV  = data.reduce((a,r) => a+r.ventas,   0);
        const totU  = data.reduce((a,r) => a+r.utilidad, 0);
        const totUn = data.reduce((a,r) => a+r.unidades, 0);
        const totTk = data.reduce((a,r) => a+r.tickets,  0);

        document.getElementById('k-uni').textContent      = fN(totUn);
        document.getElementById('k-ventas').textContent   = fF(totV);
        document.getElementById('k-ventas-s').textContent = activeSucs.length + ' sucursal(es)';
        document.getElementById('k-util').textContent     = fF(totU);
        document.getElementById('k-util-sub').textContent = 'Margen: ' + fP(totV>0 ? totU/totV : 0);
        document.getElementById('k-tkt').textContent      = fN(totTk);
        document.getElementById('k-vtkt').textContent     = totTk>0 ? fF(totV/totTk) : '—';
        document.getElementById('k-utkt').textContent     = totTk>0 ? (totUn/totTk).toFixed(1) : '—';

        updateTabla(dates, byDate);
        updateChart('ventas',  dates, byDate, activeSucs, multi);
        updateChart('tickets', dates, byDate, activeSucs, multi);
        updateHistorico();
        updateLineas();
        fitTables();
    }

    function updateTabla(dates, byDate){
        const tbody = document.getElementById('tabla-body');
        const valid = dates.filter(d => byDate[d]);
        if(!valid.length){
            tbody.innerHTML = '<tr><td colspan="7" class="empty">Selecciona sucursales para mapear la grilla.</td></tr>';
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = valid.map(d => {
            const r  = byDate[d];
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            const pill = mg>=.50 ? 'hi' : 'mi';
            const dn = DAYS[new Date(d+'T12:00:00').getDay()];
            const mAbr = MESES_ABR[parseInt(d.slice(5,7),10)-1];
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td class="date" data-label="Fecha">${d.slice(8)} ${mAbr}</td><td class="dayname" data-label="Día">${dn}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${pill}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label="" colspan="2"><b>TOTAL PERÍODO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
    }

    function updateChart(field, dates, byDate, activeSucs, multi){
        dc(field);
        const labels    = dates.map(d => d.slice(8));
        const isVentas  = field === 'ventas';
        const monoColor = isVentas ? '#7030A0' : '#2196A8';

        const datasets  = multi
            ? activeSucs.map(s => ({
                label: s,
                data: dates.map(d => { const r = RAW.find(x => x.FechaStr===d && x.NombreSucursal===s); return r ? r[field] : 0; }),
                borderColor: SC[s]||'#888', borderWidth:1.6,
                pointRadius:3.5, pointHitRadius:20, pointHoverRadius:6, tension:0.12, fill:false
              }))
            : [{
                label: isVentas ? 'Ventas c/Desc' : 'Cantidad de Tickets',
                data: dates.map(d => byDate[d] ? byDate[d][field] : 0),
                borderColor: monoColor, backgroundColor: monoColor+'08', borderWidth:2.2,
                pointRadius:3.5, pointHitRadius:20, pointHoverRadius:6, tension:0.12, fill:true
              }];

        charts[field] = new Chart(document.getElementById('chart-'+field), {
            type:'line', data:{labels, datasets},
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:multi, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:12,color:'#615e66',font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:PREMIUM_TOOLTIP_OPTS
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:10},color:'#85808c'}},
                    y:{grid:{color:'#f6f4f8'}, ticks:{font:{size:10},color:'#85808c', callback: isVentas ? v=>fM(v) : v=>v}}
                }
            }
        });
    }

    function updateLineas(){
        dc('lineas');
        const tbody    = document.getElementById('tabla-lineas');
        const filtered = getFilteredLineas();
        const lineMap  = {};
        let totalUnidades=0, totalVentas=0, totalUtilidad=0;

        filtered.forEach(r => {
            if(!lineMap[r.Línea]) lineMap[r.Línea] = {Línea:r.Línea, unidades:0, ventas:0, utilidad:0};
            lineMap[r.Línea].unidades += r.unidades;
            lineMap[r.Línea].ventas   += r.ventas;
            lineMap[r.Línea].utilidad += r.utilidad;
            totalUnidades += r.unidades; totalVentas += r.ventas; totalUtilidad += r.utilidad;
        });

        const sorted = Object.values(lineMap).sort((a,b) => b[currentLineMetric]-a[currentLineMetric]);

        if(!sorted.length){
            tbody.innerHTML = '<tr><td colspan="5" class="empty">Selecciona sucursales para desplegar líneas.</td></tr>';
            document.getElementById('lineas-chart-wrap').style.height = '220px';
            fitTables();
            return;
        }

        const rows = sorted.map(l => {
            const mg = l.ventas>0 ? l.utilidad/l.ventas : 0;
            return `<tr><td data-label="Línea"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${LC[l.Línea]||'#888'};margin-right:7px"></span>${l.Línea}</td><td class="r" data-label="Unidades">${fN(l.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(l.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td></tr>`;
        }).join('');

        const totalMargen = totalVentas>0 ? totalUtilidad/totalVentas : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL ACUMULADO</b></td><td class="r" data-label="Unidades">${fN(totalUnidades)}</td><td class="r" data-label="Ventas $">${fF(totalVentas)}</td><td class="r" data-label="Utilidad">${fF(totalUtilidad)}</td><td class="r" data-label="Margen"><span class="pill ${totalMargen>=.50?'hi':'mi'}">${fP(totalMargen)}</span></td></tr>`;

        const maxVal = Math.max(...sorted.map(l => l[currentLineMetric]))||0;

        // Altura dinámica según el número de líneas mostradas (mismo patrón
        // que fab-chart-wrap): crece con cada barra en vez de estirarse por
        // CSS para igualar la altura de la tabla vecina.
        document.getElementById('lineas-chart-wrap').style.height =
            Math.max(220, sorted.length*34 + 70) + 'px';

        charts['lineas'] = new Chart(document.getElementById('chart-lineas'), {
            type:'bar',
            data:{labels:sorted.map(l=>l.Línea), datasets:[{data:sorted.map(l=>l[currentLineMetric]), backgroundColor:sorted.map(l=>LC[l.Línea]||'#888'), borderRadius:5, barThickness:14}]},
            options:{
                indexAxis:'y', responsive:true, maintainAspectRatio:false,
                layout:{padding:{right:28}},
                plugins:{
                    legend:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{label:ctx=>(currentLineMetric==='ventas'||currentLineMetric==='utilidad')?' '+fF(ctx.raw):' '+fN(ctx.raw)+' uds'}},
                    datalabels:{display:true,anchor:'end',align:'end',color:'#4f4c54',font:{weight:'600',size:9.5},
                        formatter:(value,ctx)=>{const t=ctx.chart.data.datasets[ctx.datasetIndex].data.reduce((a,b)=>a+b,0); return fP(t>0?value/t:0);}}
                },
                scales:{
                    x:{grid:{display:false},border:{display:false},suggestedMax:maxVal*1.15,ticks:{font:{size:9},color:'#85808c',callback:currentLineMetric==='unidades'?v=>fN(v):v=>fM(v)}},
                    y:{grid:{display:false},border:{display:false},ticks:{font:{size:10,weight:'600'},color:'#222126'}}
                }
            }
        });
        fitTables();
    }

    function updateHistorico(){
        dc('historico');
        const tbody    = document.getElementById('tabla-historico');
        const filtered = getFilteredHist();

        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales y meses para ver el histórico.</td></tr>';
            fitTables();
            return;
        }

        const periodoMap   = {};
        const periodoOrder = [];
        filtered.forEach(r => {
            if(!periodoMap[r.PeriodoLabel]){
                periodoMap[r.PeriodoLabel] = {PeriodoLabel:r.PeriodoLabel, Año:r.Año, MesNum:r.MesNum,
                    unidades:0, ventas:0, utilidad:0, tickets:0};
                periodoOrder.push(r.PeriodoLabel);
            }
            periodoMap[r.PeriodoLabel].unidades  += r.unidades;
            periodoMap[r.PeriodoLabel].ventas    += r.ventas;
            periodoMap[r.PeriodoLabel].utilidad  += r.utilidad;
            periodoMap[r.PeriodoLabel].tickets   += r.tickets;
        });

        const sorted = periodoOrder
            .map(p => periodoMap[p])
            .sort((a,b) => a.Año!==b.Año ? a.Año-b.Año : a.MesNum-b.MesNum);

        let tUn=0,tV=0,tU=0,tTk=0;
        sorted.forEach(r => { tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets; });
        const totMg = tV>0 ? tU/tV : 0;

        const rows = sorted.map(r => {
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            return `<tr><td class="date" data-label="Período">${r.PeriodoLabel}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');

        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL HISTÓRICO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;

        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        const isMoneda   = currentHistMetric==='ventas' || currentHistMetric==='utilidad';
        const labels     = sorted.map(r => r.PeriodoLabel);
        const margenes   = sorted.map(r => r.ventas>0 ? +(r.utilidad/r.ventas*100).toFixed(2) : 0);

        const barDatasets = multi
            ? activeSucs.map(s => {
                const color = SC[s]||'#888';
                return {
                    type:'bar', label:s, yAxisID:'y',
                    data: sorted.map(p => {
                        return filtered
                            .filter(r => r.NombreSucursal===s && r.PeriodoLabel===p.PeriodoLabel)
                            .reduce((a,r) => a+r[currentHistMetric], 0);
                    }),
                    backgroundColor: color+'cc', borderColor: color,
                    borderWidth: 1, borderRadius: 3, stack: 'stack0',
                };
              })
            : [{
                type:'bar',
                label: currentHistMetric==='ventas'?'Ventas c/Desc':currentHistMetric==='utilidad'?'Utilidad':currentHistMetric==='unidades'?'Unidades':'Tickets',
                data: sorted.map(r => r[currentHistMetric]),
                backgroundColor:'#7030A0cc', borderColor:'#7030A0',
                borderWidth:1, borderRadius:4, yAxisID:'y',
              }];

        const margenLine = {
            type:'line', label:'Margen %', yAxisID:'y2',
            data: margenes,
            borderColor:'#aa7300', backgroundColor:'transparent',
            borderWidth:2, pointRadius:3.5, pointHoverRadius:6, tension:0.2,
            datalabels:{display:false}
        };

        charts['historico'] = new Chart(document.getElementById('chart-historico'), {
            data:{ labels, datasets:[...barDatasets, margenLine] },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#615e66',
                            font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{
                        label: ctx => ctx.dataset.yAxisID==='y2'
                            ? ` Margen: ${ctx.raw.toFixed(2)}%`
                            : isMoneda ? ` ${ctx.dataset.label}: ${fF(ctx.raw)}` : ` ${ctx.dataset.label}: ${fN(ctx.raw)}`
                    }}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#85808c',maxRotation:45}},
                    y:{stacked:true, position:'left', grid:{color:'#f6f4f8'}, ticks:{font:{size:9},color:'#85808c', callback: v => isMoneda?fM(v):fN(v)}},
                    y2:{position:'right', grid:{display:false}, ticks:{font:{size:9},color:'#aa7300', callback: v => v.toFixed(1)+'%'}}
                }
            }
        });
        fitTables();
    }

    // ── PESTAÑA TOP ARTÍCULOS ──
    function updateTopArticulos(){
        dc('bubble');
        const filtered = getFilteredTopArt();

        if(!filtered.length){
            document.getElementById('tabla-top').innerHTML =
                '<tr><td colspan="7" class="empty">Selecciona sucursales y meses para ver el ranking.</td></tr>';
            ['tk-uni','tk-ventas','tk-margen'].forEach(id => {
                document.getElementById(id).textContent = '—';
            });
            fitTables();
            return;
        }

        // Agrupar por artículo
        const artMap = {};
        filtered.forEach(r => {
            if(!artMap[r.Artículo]) artMap[r.Artículo] = {
                Artículo:    r.Artículo,
                Descripción: r.Descripción,
                Fabricante:  r.Fabricante,
                unidades: 0,
                ventas:   0,
                utilidad: 0,
            };
            artMap[r.Artículo].unidades += r.unidades;
            artMap[r.Artículo].ventas   += r.ventas;
            artMap[r.Artículo].utilidad += r.utilidad;
        });

        const top10 = Object.values(artMap)
            .sort((a,b) => b.unidades - a.unidades)
            .slice(0, 10);

        const histFiltrado    = getFilteredHist();
        const totalHistVentas = histFiltrado.reduce((a,r) => a + r.ventas, 0);

        const totUni    = top10.reduce((a,r) => a+r.unidades, 0);
        const totVentas = top10.reduce((a,r) => a+r.ventas,   0);

        document.getElementById('tk-uni').textContent    = fN(totUni);
        document.getElementById('tk-ventas').textContent = fF(totVentas);
        document.getElementById('tk-margen').textContent = fP(totalHistVentas>0 ? totVentas/totalHistVentas : 0);

        // ── Tabla: descripciones y fabricantes COMPLETOS, sin truncar ──
        const tbody = document.getElementById('tabla-top');
        const rows = top10.map((r, i) => {
            const rank = i + 1;
            const rc   = rank===1?'r1':rank===2?'r2':rank===3?'r3':'';
            const pct  = totalHistVentas>0 ? r.ventas/totalHistVentas : 0;
            return `<tr>
              <td data-label="#" style="text-align:right"><span class="badge-rank ${rc}">${rank}</span></td>
              <td class="art-code" data-label="Artículo">${r.Artículo}</td>
              <td class="art-desc" data-label="Descripción">${r.Descripción}</td>
              <td class="art-fab" data-label="Fabricante">${r.Fabricante}</td>
              <td class="r" data-label="Unidades"><b>${fN(r.unidades)}</b></td>
              <td class="r" data-label="Venta $">${fF(r.ventas)}</td>
              <td class="r" data-label="% Venta"><span class="pill hi">${fP(pct)}</span></td>
            </tr>`;
        }).join('');
        tbody.innerHTML = rows;

        // ── Gráfico de burbujas ──
        // Escalado por ÁREA: el ojo percibe tamaño por área, no por radio.
        // Rango minR=10 / maxR=44 con la altura ampliada a 560px da proporciones equilibradas.
        const maxUni = Math.max(...top10.map(r => r.unidades)) || 1;
        const minUni = Math.min(...top10.map(r => r.unidades)) || 0;
        const minR = 10, maxR = 44;
        const rScale = v => {
            const t       = (v - minUni) / (maxUni - minUni || 1);
            const areaMin = minR * minR, areaMax = maxR * maxR;
            return Math.sqrt(areaMin + t * (areaMax - areaMin));
        };

        const bubbleDatasets = top10.map((r, i) => {
            const margen = r.ventas > 0 ? +(r.utilidad / r.ventas * 100).toFixed(2) : 0;
            return {
                label: r.Artículo,
                data: [{ x: r.ventas, y: margen, r: rScale(r.unidades) }],
                backgroundColor: PALETTE_BUBBLE[i % PALETTE_BUBBLE.length] + 'bb',
                borderColor:     PALETTE_BUBBLE[i % PALETTE_BUBBLE.length],
                borderWidth: 1.5,
                // Descripción y fabricante completos en el tooltip
                _meta: { desc: r.Descripción, fab: r.Fabricante, uni: r.unidades },
            };
        });

        charts['bubble'] = new Chart(document.getElementById('chart-bubble'), {
            type: 'bubble',
            data: { datasets: bubbleDatasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'nearest', intersect: true },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        onClick: null,
                        labels: {
                            boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'circle',
                            padding: 12, color: '#615e66',
                            font: { family: "'Segoe UI', sans-serif", size: 9.5, weight: '600' }
                        }
                    },
                    datalabels: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            title: ctx => ctx[0].dataset.label,
                            // Descripción completa en el tooltip (sin truncar)
                            label: ctx => {
                                const ds  = ctx.dataset;
                                const pt  = ctx.raw;
                                return [
                                    ` ${ds._meta.desc}`,
                                    ` Fabricante: ${ds._meta.fab}`,
                                    ` Unidades: ${fN(ds._meta.uni)}`,
                                    ` Ventas: ${fF(pt.x)}`,
                                    ` Margen: ${pt.y.toFixed(2)}%`,
                                ];
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        title: { display: true, text: 'Ventas $ c/Desc', color: '#85808c', font: { size: 10, weight: '600' } },
                        grid: { color: '#f6f4f8' },
                        ticks: { font: { size: 9 }, color: '#85808c', callback: v => fM(v) }
                    },
                    y: {
                        title: { display: true, text: 'Margen %', color: '#85808c', font: { size: 10, weight: '600' } },
                        grid: { color: '#f6f4f8' },
                        ticks: { font: { size: 9 }, color: '#85808c', callback: v => v.toFixed(1)+'%' }
                    }
                }
            }
        });
        fitTables();
    }

    // ── PESTAÑA LÍNEAS Y CATEGORÍAS (treemap) ──
    window.toggleLineaCat = function(rid){
        const rows = document.querySelectorAll('.'+rid+'-child');
        if(!rows.length) return;
        const show = rows[0].style.display === 'none';
        rows.forEach(r => { r.style.display = show ? 'table-row' : 'none'; });
        const icon = document.getElementById(rid+'-icon');
        if(icon) icon.textContent = show ? '▾' : '▸';
        fitTables();
    };

    window.toggleFabQuartil = function(rid){
        const rows = document.querySelectorAll('.'+rid+'-child');
        if(!rows.length) return;
        const show = rows[0].style.display === 'none';
        rows.forEach(r => { r.style.display = show ? 'table-row' : 'none'; });
        const icon = document.getElementById(rid+'-icon');
        if(icon) icon.textContent = show ? '▾' : '▸';
        fitTables();
    };

    function updateLineasCategoria(){
        dc('treemap');
        const tbody    = document.getElementById('tabla-lineascategoria');
        const filtered = getFilteredLineasCat();

        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona sucursales y meses para ver el análisis.</td></tr>';
            document.getElementById('tk2-linea-nombre').textContent = '—';
            document.getElementById('tk2-linea-sub').textContent    = '—';
            document.getElementById('tk2-cat-nombre').textContent   = '—';
            document.getElementById('tk2-cat-sub').textContent      = '—';
            document.getElementById('tk2-cat-margen').textContent   = '—';
            fitTables();
            return;
        }

        // Agrupar por Línea + Categoría dentro de la selección activa — SIN límites:
        // se usa el mismo dataset completo (todas las líneas y categorías) tanto
        // para la tabla como para el treemap.
        const key2 = {};
        const lineaTotales = {};
        filtered.forEach(r => {
            const k = r.Línea + '|||' + r.Categoría;
            if(!key2[k]) key2[k] = {Línea:r.Línea, Categoría:r.Categoría, ventas:0, utilidad:0, unidades:0};
            key2[k].ventas    += r.ventas;
            key2[k].utilidad  += r.utilidad;
            key2[k].unidades  += r.unidades;
            lineaTotales[r.Línea] = (lineaTotales[r.Línea]||0) + r.ventas;
        });

        const allLeaves = Object.values(key2);

        // KPIs — Línea líder y Categoría líder (Top 1 de cada uno), con % de participación
        // sobre el total de ventas de la selección activa.
        const totalSeleccion = filtered.reduce((a,r) => a+r.ventas, 0);

        const lineaLiderEntry = Object.entries(lineaTotales).sort((a,b) => b[1]-a[1])[0];
        if(lineaLiderEntry){
            const [lineaLiderNombre, lineaLiderVentas] = lineaLiderEntry;
            const lineaLiderPct = totalSeleccion>0 ? lineaLiderVentas/totalSeleccion : 0;
            document.getElementById('tk2-linea-nombre').textContent = lineaLiderNombre;
            document.getElementById('tk2-linea-sub').textContent    = `${fF(lineaLiderVentas)} · ${fP(lineaLiderPct)} de participación`;
        }

        const catLiderEntry = allLeaves.slice().sort((a,b) => b.ventas - a.ventas)[0];
        if(catLiderEntry){
            const catLiderPct = totalSeleccion>0 ? catLiderEntry.ventas/totalSeleccion : 0;
            const catLiderMg  = catLiderEntry.ventas>0 ? catLiderEntry.utilidad/catLiderEntry.ventas : 0;
            document.getElementById('tk2-cat-nombre').textContent = catLiderEntry.Categoría;
            document.getElementById('tk2-cat-sub').textContent    = `${catLiderEntry.Línea} · ${fF(catLiderEntry.ventas)} · ${fP(catLiderPct)} de participación`;
            document.getElementById('tk2-cat-margen').textContent = fP(catLiderMg);
        }

        // Tabla expandible: Línea (fila padre) > Categoría (filas hijas ocultas por defecto) — todas, sin agrupar en "Otras"
        const lineaMap = {};
        allLeaves.forEach(r => {
            if(!lineaMap[r.Línea]) lineaMap[r.Línea] = {Línea:r.Línea, ventas:0, utilidad:0, unidades:0, cats:[]};
            lineaMap[r.Línea].ventas   += r.ventas;
            lineaMap[r.Línea].utilidad += r.utilidad;
            lineaMap[r.Línea].unidades += r.unidades;
            lineaMap[r.Línea].cats.push(r);
        });

        const lineasSorted = Object.values(lineaMap).sort((a,b) => b.ventas-a.ventas);
        let html = '';
        lineasSorted.forEach((l, i) => {
            const mg  = l.ventas>0 ? l.utilidad/l.ventas : 0;
            const rid = 'lc'+i;
            html += `<tr class="lc-parent" onclick="toggleLineaCat('${rid}')">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">▸</span></td>
                <td data-label="Línea"><b>${l.Línea}</b></td>
                <td class="r" data-label="Unidades">${fN(l.unidades)}</td>
                <td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td>
                <td class="r" data-label="Utilidad">${fF(l.utilidad)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td>
            </tr>`;
            const catsSorted = l.cats.slice().sort((a,b) => b.ventas-a.ventas);
            catsSorted.forEach(c => {
                const cmg = c.ventas>0 ? c.utilidad/c.ventas : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:none">
                    <td data-label=""></td>
                    <td data-label="Categoría">${c.Categoría}</td>
                    <td class="r" data-label="Unidades">${fN(c.unidades)}</td>
                    <td class="r" data-label="Ventas $">${fF(c.ventas)}</td>
                    <td class="r" data-label="Utilidad">${fF(c.utilidad)}</td>
                    <td class="r" data-label="Margen"><span class="pill ${cmg>=.50?'hi':'mi'}">${fP(cmg)}</span></td>
                </tr>`;
            });
        });
        tbody.innerHTML = html;

        // ── DATASET PARA EL TREEMAP: sólo Top 3 categorías por línea ──
        // La tabla de arriba usa 'allLeaves' (todo, sin límite). El treemap se
        // limita a las 3 categorías de mayor venta de cada línea para que siga
        // siendo legible; el resto no se muestra ahí, pero sigue disponible en
        // la tabla completa.
        const treemapLeaves = [];
        Object.keys(lineaTotales).forEach(lineaName => {
            const topCats = allLeaves
                .filter(r => r.Línea === lineaName)
                .sort((a,b) => b.ventas - a.ventas)
                .slice(0, 3);
            treemapLeaves.push(...topCats);
        });

        // Rangos de Ventas c/Desc para normalizar el color por nivel (Línea vs Categoría),
        // calculados sobre lo que realmente se dibuja en el treemap.
        const treemapLineaTotales = {};
        treemapLeaves.forEach(r => { treemapLineaTotales[r.Línea] = (treemapLineaTotales[r.Línea]||0) + r.ventas; });

        const lineaVals = Object.values(treemapLineaTotales);
        const leafVals  = treemapLeaves.map(r => r.ventas);
        const lineaMin  = Math.min(...lineaVals), lineaMax = Math.max(...lineaVals);
        const leafMin   = Math.min(...leafVals),  leafMax  = Math.max(...leafVals);

        // Mapa de respaldo: valor acumulado (redondeado) de cada línea -> nombre de línea.
        // Sirve como fallback para el tooltip de los recuadros de cabecera (nivel Línea),
        // porque la librería de treemap no siempre expone raw.g de forma confiable ahí.
        const headerValueMap = {};
        Object.keys(treemapLineaTotales).forEach(lineaName => {
            headerValueMap[Math.round(treemapLineaTotales[lineaName])] = lineaName;
        });

        // Altura dinámica según el número de celdas resultante de la selección activa
        // (similar a como se hace con el gráfico de fabricantes): mínimo 360px,
        // creciendo conforme aumenten las combinaciones Línea+Categoría mostradas.
        const numCeldas = treemapLeaves.length;
        const alturaTreemap = Math.max(360, Math.min(900, 360 + Math.max(0, numCeldas - 40) * 3));
        document.getElementById('lc-chart-wrap').style.height = alturaTreemap + 'px';

        // Treemap: jerarquía Línea > Categoría · tamaño=Ventas · color=intensidad de Ventas
        charts['treemap'] = new Chart(document.getElementById('chart-treemap'), {
            type: 'treemap',
            data: {
                datasets: [{
                    label: 'Ventas c/Desc',
                    tree: treemapLeaves,
                    key: 'ventas',
                    groups: ['Línea','Categoría'],
                    spacing: 1,
                    borderWidth: 1.5,
                    borderColor: '#ffffff',
                    displayMode: 'headerBoxes',
                    captions: {
                        display: true,
                        color: '#ffffff',
                        font: { weight: '700', size: 10 },
                        formatter: ctx => ctx.raw.g
                    },
                    labels: {
                        display: ctx => ctx.type === 'data' && ctx.raw.l > 0 && ctx.raw.w > 46 && ctx.raw.h > 18,
                        color: ctx => {
                            if(ctx.type !== 'data' || ctx.raw.l === 0) return '#ffffff';
                            const v = ctx.raw.v || 0;
                            const t = leafMax > leafMin ? (v-leafMin)/(leafMax-leafMin) : 1;
                            return t < 0.45 ? '#4a3a5c' : '#ffffff';
                        },
                        font: { size: 8.5, weight: '600' },
                        overflow: 'hidden',
                        formatter: ctx => {
                            if(ctx.type !== 'data' || ctx.raw.l === 0) return '';
                            return ctx.raw.g;
                        }
                    },
                    backgroundColor: ctx => {
                        if(ctx.type !== 'data') return 'transparent';
                        const v = ctx.raw.v || 0;
                        if(ctx.raw.l === 0){
                            const t = lineaMax > lineaMin ? (v-lineaMin)/(lineaMax-lineaMin) : 1;
                            return salesColor(0.55 + t*0.45); // headers siempre en la mitad oscura del rango
                        }
                        const t = leafMax > leafMin ? (v-leafMin)/(leafMax-leafMin) : 1;
                        return salesColor(t);
                    },
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: { display: false},
                    legend: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            // Nivel de cabecera (Línea): "Línea: Nombre", con headerValueMap
                            // como respaldo cuando raw.g no trae el nombre.
                            // Nivel de categoría (hoja): se usa raw._data — el objeto original
                            // completo que conserva chartjs-chart-treemap — para obtener la
                            // Línea real sin depender de un emparejamiento frágil por valor.
                            title: ctx => {
                                const raw = ctx[0] && ctx[0].raw;
                                if(!raw) return '';
                                if(raw.l === 0){
                                    const nombre = raw.g || headerValueMap[Math.round(raw.v || raw.s || 0)] || '';
                                    return nombre ? `Línea: ${nombre}` : 'Línea';
                                }
                                const orig = raw._data;
                                if(orig && orig.Línea && orig.Categoría) return `${orig.Línea} · ${orig.Categoría}`;
                                return raw.g || 'Categoría';
                            },
                            label: ctx => ' Ventas: ' + fF(ctx.raw.v || 0)
                        }
                    }
                }
            }
        });
        fitTables();
    }

    // ── PESTAÑA FABRICANTES ──
    const QUART_LABELS = {1:'Cuartil 1 · Mayor venta', 2:'Cuartil 2', 3:'Cuartil 3', 4:'Cuartil 4 · Menor venta'};
    const QUART_COLORS = {
        1: {bg:'#7030A0', fg:'#ffffff', tint:'#7030A014'},
        2: {bg:'#9C6CBE', fg:'#ffffff', tint:'#9C6CBE12'},
        3: {bg:'#D9C6EC', fg:'#4a3a5c', tint:'#D9C6EC20'},
        4: {bg:'#F0EAF6', fg:'#6d6575', tint:'#F0EAF640'},
    };

    function updateFabricantes(){
        dc('fabricantes');
        const tbody    = document.getElementById('tabla-fabricantes');
        const filtered = getFilteredFabricantes();

        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="5" class="empty">Selecciona sucursales y meses para ver los fabricantes.</td></tr>';
            ['tk3-ventas','tk3-margen','tk3-pct'].forEach(id => { document.getElementById(id).textContent = '—'; });
            document.getElementById('tk3-pct-sub').textContent = 'Necesarios para alcanzar el 50% de ventas';
            fitTables();
            return;
        }

        // Agrupar por Fabricante dentro de la selección activa
        const fabMap = {};
        let totalSeleccion = 0;
        filtered.forEach(r => {
            if(!fabMap[r.Fabricante]) fabMap[r.Fabricante] = {Fabricante:r.Fabricante, ventas:0, utilidad:0, unidades:0};
            fabMap[r.Fabricante].ventas   += r.ventas;
            fabMap[r.Fabricante].utilidad += r.utilidad;
            fabMap[r.Fabricante].unidades += r.unidades;
            totalSeleccion += r.ventas;
        });

        const sorted = Object.values(fabMap).sort((a,b) => b.ventas - a.ventas);

        // % individual, % acumulado y cuartil de cada fabricante.
        // El cuartil se asigna según el punto donde INICIA su tramo acumulado (no donde termina),
        // así el fabricante más grande siempre cae en el Cuartil 1, sin importar qué tan dominante sea.
        let cum = 0;
        sorted.forEach(f => {
            const start = totalSeleccion>0 ? cum/totalSeleccion : 0;
            cum += f.ventas;
            f.pct    = totalSeleccion>0 ? f.ventas/totalSeleccion : 0;
            f.cumPct = totalSeleccion>0 ? cum/totalSeleccion : 0;
            if(start < 0.25)      f.cuartil = 1;
            else if(start < 0.50) f.cuartil = 2;
            else if(start < 0.75) f.cuartil = 3;
            else                  f.cuartil = 4;
        });

        // Fabricantes que, acumulados de mayor a menor venta, alcanzan (o superan) el 50% del total.
        const paretoList = [];
        let cumPareto = 0;
        for(const f of sorted){
            paretoList.push(f);
            cumPareto += f.ventas;
            if(totalSeleccion>0 && cumPareto/totalSeleccion >= 0.5) break;
        }

        const totVentasPareto = paretoList.reduce((a,r) => a+r.ventas,   0);
        const totUtilPareto   = paretoList.reduce((a,r) => a+r.utilidad, 0);

        document.getElementById('tk3-ventas').textContent = fF(totVentasPareto);
        document.getElementById('tk3-margen').textContent = fP(totVentasPareto>0 ? totUtilPareto/totVentasPareto : 0);
        document.getElementById('tk3-pct').textContent    = `${paretoList.length} de ${sorted.length}`;
        document.getElementById('tk3-pct-sub').textContent = `Cubren el ${fP(totalSeleccion>0 ? totVentasPareto/totalSeleccion : 0)} de las ventas`;

        // ── Tabla completa: todos los fabricantes, agrupados por cuartil (expandible) ──
        const groups = {1:[], 2:[], 3:[], 4:[]};
        sorted.forEach(f => groups[f.cuartil].push(f));

        const maxVentas = sorted.length ? sorted[0].ventas : 0;
        let html = '';
        let cumBeforeGroup = 0; // rastrea el % acumulado antes de cada cuartil, para mostrar su rango real

        [1,2,3,4].forEach(q => {
            const items = groups[q];
            if(!items.length) return;
            const qVentas = items.reduce((a,r) => a+r.ventas,   0);
            const qUtil   = items.reduce((a,r) => a+r.utilidad, 0);
            const qMg     = qVentas>0 ? qUtil/qVentas : 0;
            const qPct    = totalSeleccion>0 ? qVentas/totalSeleccion : 0;
            const rid     = 'fq'+q;
            const qc      = QUART_COLORS[q];
            const expanded = q === 1; // El cuartil de mayor venta inicia desplegado

            // Rango real de % acumulado que cubre este cuartil (ej. 0.00%–24.80% acumulado)
            const startPct = totalSeleccion>0 ? cumBeforeGroup/totalSeleccion : 0;
            cumBeforeGroup += qVentas;
            const endPct   = totalSeleccion>0 ? cumBeforeGroup/totalSeleccion : 0;

            html += `<tr class="fq-parent" onclick="toggleFabQuartil('${rid}')" style="background:${qc.tint}">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">${expanded?'▾':'▸'}</span></td>
                <td data-label="Cuartil"><span class="fq-badge" style="background:${qc.bg};color:${qc.fg}">${QUART_LABELS[q]}</span>
                    <span style="color:#85808c;font-size:.68rem;margin-left:6px;display:block;text-align:left;margin-top:4px">${items.length} fabricante${items.length!==1?'s':''} · ${fP(startPct)}–${fP(endPct)} acumulado</span></td>
                <td class="r" data-label="Ventas $"><b>${fF(qVentas)}</b></td>
                <td class="r" data-label="Margen"><span class="pill ${qMg>=.50?'hi':'mi'}">${fP(qMg)}</span></td>
                <td class="r" data-label="% Participación"><b>${fP(qPct)}</b></td>
            </tr>`;

            items.forEach((f, idx) => {
                const mg  = f.ventas>0 ? f.utilidad/f.ventas : 0;
                const barW = maxVentas>0 ? (f.ventas/maxVentas*100).toFixed(1) : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:${expanded?'table-row':'none'}">
                    <td class="fab-rank-num" data-label="#" style="text-align:right">${idx+1}</td>
                    <td data-label="Fabricante">${f.Fabricante}</td>
                    <td class="r fab-bar-cell" data-label="Ventas $"><span class="fab-bar-bg" style="width:${barW}%;background:${qc.bg}22"></span><b>${fF(f.ventas)}</b></td>
                    <td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td>
                    <td class="r" data-label="% Participación">${fP(f.pct)}</td>
                </tr>`;
            });
        });
        tbody.innerHTML = html;

        // ── Gráfico: solo los fabricantes que concentran el 50% de las ventas ──
        document.getElementById('fab-chart-wrap').style.height = Math.max(220, paretoList.length*34 + 60) + 'px';
        const maxVal = Math.max(...paretoList.map(f => f.ventas)) || 0;

        charts['fabricantes'] = new Chart(document.getElementById('chart-fabricantes'), {
            type: 'bar',
            data: {
                labels: paretoList.map(f => f.Fabricante),
                datasets: [{
                    data: paretoList.map(f => f.ventas),
                    backgroundColor: paretoList.map((f,i) => QUART_COLORS[f.cuartil].bg),
                    borderRadius: 5,
                    barThickness: 16
                }]
            },
            options: {
                indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                layout: { padding: { right: 34 } },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            label: ctx => {
                                const f  = paretoList[ctx.dataIndex];
                                const mg = f.ventas>0 ? f.utilidad/f.ventas : 0;
                                return [' Ventas: ' + fF(f.ventas), ' Margen: ' + fP(mg), ' % Participación: ' + fP(f.pct)];
                            }
                        }
                    },
                    datalabels: {
                        display: true, anchor: 'end', align: 'end',
                        color: '#4f4c54', font: { weight: '600', size: 9.5 },
                        formatter: (value, ctx) => fP(paretoList[ctx.dataIndex].pct)
                    }
                },
                scales: {
                    x: { grid:{display:false}, border:{display:false}, suggestedMax:maxVal*1.15, ticks:{font:{size:9},color:'#85808c',callback:v=>fM(v)} },
                    y: { grid:{display:false}, border:{display:false}, ticks:{font:{size:10,weight:'600'},color:'#222126'} }
                }
            }
        });
        fitTables();
    }

    // ── INICIO ──
    // Se dibuja SOLO la pestaña visible (Objetivos). Las demás se construyen la
    // primera vez que el usuario entra en ellas, vía switchTab(). Esto evita que
    // Chart.js renderice gráficas dentro de contenedores con display:none, lo
    // que las dejaría con dimensiones cero (aplastadas o en blanco).
    buildButtons();
    buildMesButtons();
    updateObjetivos();
    setTimeout(fitTables, 100);
});
</script>
</body>
</html>"""

    html = html_template.replace("__FECHA_VALOR__",   fecha_reporte)
    html = html.replace("__FECHA_INFO__",             fecha_info)
    html = html.replace("__MES_HEADER__",             mes_header)
    html = html.replace("__DATA_JSON__",              data_json)
    html = html.replace("__LINEA_JSON__",             linea_json)
    html = html.replace("__HISTORICO_JSON__",         historico_json)
    html = html.replace("__TOP_ART_JSON__",           top_art_json)
    html = html.replace("__LINEAS_CAT_JSON__",        lineas_cat_json)
    html = html.replace("__FABRICANTES_JSON__",       fabricantes_json)
    html = html.replace("__PRESUPUESTO_JSON__",       presupuesto_json)
    html = html.replace("__PERIODOS_JSON__",          periodos_json)
    html = html.replace("__SUCURSALES_JSON__",        sucursales_json)
    return html


def main():
    try:
        print(f"Leyendo: {EXCEL_PATH}")
        xl  = pd.read_excel(EXCEL_PATH, sheet_name=None)
        vmc = xl["VentasMesCurso"].copy()
        vm  = xl["VentasMensuales"].copy()
        tkt = xl["TicketsMensuales"].copy()
        dim = xl["TablasDimensión"].copy()

        objetivos_df = xl.get("ObjetivosVentas")
        if objetivos_df is None:
            print("⚠️  No se encontró la hoja 'ObjetivosVentas'; el pronóstico se generará sin presupuestos (Presupuesto = 0).")
            objetivos_df = pd.DataFrame(columns=["Sucursal", "Presupuesto"])
        else:
            objetivos_df = objetivos_df.copy()

        suc = dim[["Clave sucursal","Nombre de sucursal"]].dropna().drop_duplicates()
        suc.columns = ["ClaveSucursal","NombreSucursal"]
        suc["ClaveSucursal"] = pd.to_numeric(suc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
        lista_sucursales = sorted(suc["NombreSucursal"].unique().tolist())

        art     = dim[["Artículo","Línea"]].dropna().drop_duplicates("Artículo")
        art_dim = dim[["Artículo","Línea","Categoría","Descripción","Fabricante"]].dropna(subset=["Artículo"]).drop_duplicates("Artículo")

        print("Procesando datos...")
        agg             = procesar_mes_curso(vmc, suc, BOL_EXCLUIR)
        linea_agg       = procesar_lineas(vm, art, suc, BOL_EXCLUIR)
        historico_agg   = procesar_historico(vm, tkt, suc, BOL_EXCLUIR)
        top_art_agg     = procesar_top_articulos(vm, art_dim, suc, BOL_EXCLUIR)
        lineas_cat_agg  = procesar_lineas_categoria(vm, art_dim, suc, BOL_EXCLUIR)
        fabricantes_agg = procesar_fabricantes(vm, art_dim, suc, BOL_EXCLUIR)
        presupuesto_agg = procesar_presupuesto(agg, objetivos_df, suc, FECHA_BASE)

        print("Generando HTML final...")
        fecha_reporte, fecha_info, mes_header = formatear_fechas(FECHA_BASE)
        html = generar_html(agg, linea_agg, historico_agg, top_art_agg, lineas_cat_agg, fabricantes_agg,
                            presupuesto_agg, lista_sucursales, fecha_reporte, fecha_info, mes_header)

        Path(OUTPUT_PATH).write_text(html, encoding="utf-8")
        print(f"✅ Dashboard generado exitosamente en: {OUTPUT_PATH}")

    except Exception as e:
        import traceback
        print(f"❌ Error crítico durante la ejecución: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()