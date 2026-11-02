"""Validaciones no destructivas: nunca alteran los importes de la planilla."""
import math
from collections import defaultdict
import pandas as pd

INGRESOS = list(range(21, 33))  # V:AG, incluido adelantos AF
DESCUENTOS = list(range(33, 39))  # AH:AM
COL_NUMERICAS = list(range(7, 43))  # H:AQ


def _num(value):
    if pd.isna(value) or value == '':
        return 0.0
    try:
        x = float(value)
        return x if math.isfinite(x) else 0.0
    except (TypeError, ValueError):
        return 0.0


def revisar_excel(df, registros, consolidados):
    """Devuelve lista de incidencias sin cambiar datos. Columnas A=0..AQ=42."""
    alertas = []
    def avisar(nivel, fila, documento, asunto, descripcion):
        alertas.append(dict(nivel=nivel, fila=fila, dni=str(documento or ''),
                            asunto=asunto, detalle=descripcion))

    por_documento = defaultdict(list)
    for i in range(5, len(df.index)):
        fila = df.iloc[i]
        dnivalor = fila.iloc[3]
        if pd.isna(dnivalor) or str(dnivalor).strip() == '':
            if not pd.isna(fila.iloc[1]) and str(fila.iloc[1]).strip():
                avisar('error', i+1, '', 'Documento vacío', 'Hay un nombre sin documento. La fila no se importó.')
            continue
        dni = str(dnivalor).strip()
        if dni.endswith('.0'):
            dni = dni[:-2]
        por_documento[dni].append((i+1, fila))
        if pd.isna(fila.iloc[1]) or not str(fila.iloc[1]).strip():
            avisar('error', i+1, dni, 'Nombre vacío', 'El documento tiene un registro sin nombre.')
        for col in COL_NUMERICAS:
            valor = fila.iloc[col]
            if pd.isna(valor) or valor == '':
                continue
            try:
                n = float(valor)
                if not math.isfinite(n):
                    raise ValueError()
            except (ValueError, TypeError):
                avisar('error', i+1, dni, 'Dato numérico inválido',
                      f'Columna {chr(65 + col) if col < 26 else chr(64 + col//26)+chr(65 + col%26)}: «{str(valor)[:40]}». Se trataría como cero.')
        total_ingresos = sum(_num(fila.iloc[c]) for c in INGRESOS)
        total_descuentos = sum(_num(fila.iloc[c]) for c in DESCUENTOS)
        neto = _num(fila.iloc[39])
        if abs((total_ingresos - total_descuentos) - neto) > 0.015:
            avisar('error', i+1, dni, 'Neto de fila no concilia',
                  f'Ingresos S/ {total_ingresos:.2f} – descuentos S/ {total_descuentos:.2f} ≠ neto S/ {neto:.2f}.')
        # Q representa días por registro/área. No se contrasta cada fila
        # con días únicos de H:N porque la jornada puede estar distribuida
        # entre varias áreas y producir advertencias falsas.
        if _num(fila.iloc[14]) < -1e-7 or _num(fila.iloc[15]) < -1e-7:
            avisar('aviso', i+1, dni, 'TN u horas negativas', 'Revisar si corresponde a un ajuste autorizado.')
        # No comparar O contra H:N: podrían ser equivalencias pagables, no TN físicas.
        # No sumar R:U contra P: horas al 100% pueden ser un concepto adicional.
        regimen = str(fila.iloc[4] if not pd.isna(fila.iloc[4]) else '').upper()
        comision = str(fila.iloc[5] if not pd.isna(fila.iloc[5]) else '').strip()
        if 'ONP' in regimen and comision and comision not in ('0','0.0'):
            avisar('aviso', i+1, dni, 'Comisión en registro ONP',
                  f'La columna F contiene «{comision}»; se conserva como dato de origen.')

    mapa = {t['dni']: t for t in consolidados}
    for dni, filas in por_documento.items():
        nombres = {str(f.iloc[1]).strip().upper() for _,f in filas if not pd.isna(f.iloc[1])}
        regimenes = {str(f.iloc[4]).strip().upper() for _,f in filas if not pd.isna(f.iloc[4])}
        comisiones = {str(f.iloc[5]).strip().upper() for _,f in filas if not pd.isna(f.iloc[5]) and str(f.iloc[5]).strip() not in ('0','0.0','')}
        if len(nombres) > 1:
            avisar('error', '/'.join(str(n) for n,_ in filas), dni, 'Nombres distintos',
                  'Un mismo documento tiene distintos nombres en el Excel.')
        if len(regimenes) > 1 or len(comisiones) > 1:
            avisar('aviso', '/'.join(str(n) for n,_ in filas), dni, 'Régimen/Comisión inconsistente',
                  'El documento presenta diferencias en E o F; revisar antes de emitir.')
        t = mapa.get(dni)
        if not t:
            continue
        # El consolidado de días debe ser la suma de Q de todas las áreas,
        # aunque correspondan a una fecha compartida. Solo avisamos cuando
        # absolutamente ningún registro informa días, pero sí hay actividad.
        dias_q = sum(_num(f.iloc[16]) for _, f in filas)
        actividad_registrada = any(
            _num(f.iloc[7 + d]) > 0
            for _, f in filas for d in range(7)
        )
        if dias_q == 0 and actividad_registrada:
            avisar('aviso', '/'.join(str(n) for n, _ in filas), dni,
                  'Días Q sin informar',
                  'Q suma cero en todas las áreas, aunque H:N registra actividad. Revisar el Excel original.')
        valor_ingreso = sum(_num(f.iloc[c]) for _,f in filas for c in INGRESOS)
        valor_desc = sum(_num(f.iloc[c]) for _,f in filas for c in DESCUENTOS)
        if abs(valor_ingreso - valor_desc - t['neto']) > 0.015:
            avisar('error', '/'.join(str(n) for n,_ in filas), dni, 'Neto consolidado no concilia',
                  f'Ingresos S/ {valor_ingreso:.2f}, descuentos S/ {valor_desc:.2f}, neto S/ {t["neto"]:.2f}.')
        # Si AP es referencia mensual, una tarifa distinta en el mismo área amerita revisión.
        valores_ap = defaultdict(set)
        for _,f in filas:
            if _num(f.iloc[15]) > 0:
                area = str(f.iloc[6]).strip().upper()
                valores_ap[area].add(round(_num(f.iloc[41]), 6))
        for area, valores in valores_ap.items():
            valores.discard(0.0)
            if len(valores) > 1:
                avisar('aviso', '/'.join(str(n) for n,_ in filas), dni, 'AP distinto en una misma área',
                      f'{area}: remuneraciones referenciales distintas; se conserva la primera.')
    return alertas


def asegurar_estructura(df):
    if df.shape[1] < 43:
        raise ValueError('El Excel debe contener por lo menos 43 columnas (A:AQ).')
    if df.shape[0] < 6:
        raise ValueError('El Excel no contiene filas de trabajadores desde la fila 6.')
    requisitos = {1:'NOMBRE', 3:'DOC', 6:'AREA', 14:'KILOS', 15:'HORAS', 16:'DIAS',39:'NETO'}
    for columna, palabra in requisitos.items():
        encabezado = str(df.iloc[4,columna]).upper()
        if palabra not in encabezado:
            raise ValueError(f'Encabezado inesperado en columna {columna+1}. Se esperaba un campo de {palabra}.')


def revisar_formulas_sin_resultado(ruta, df):
    """Detecta fórmulas que se importan sin un resultado numérico guardado.

    No evalúa fórmulas ni reescribe la planilla; solo advierte de un archivo
    que debe abrirse/guardarse en Excel para recalcular los resultados.
    """
    from openpyxl import load_workbook
    alertas = []
    wb = load_workbook(ruta, read_only=True, data_only=False)
    try:
        hoja = wb.worksheets[0]
        for fila in hoja.iter_rows(min_row=6, max_col=min(43, hoja.max_column)):
            indice = fila[0].row - 1
            if indice >= len(df.index):
                break
            dni_valor = df.iat[indice,3]
            if pd.isna(dni_valor) or str(dni_valor).strip() == '':
                continue
            dni = str(dni_valor).strip().removesuffix('.0')
            for celda in fila[7:43]:
                if celda.data_type == 'f' and pd.isna(df.iat[indice, celda.column-1]):
                    alertas.append(dict(nivel='error', fila=celda.row, dni=dni,
                                        asunto='Fórmula sin resultado',
                                        detalle=f'{celda.coordinate}: hay una fórmula sin resultado guardado; Excel debe recalcular y guardar el archivo.'))
    finally:
        wb.close()
    return alertas
