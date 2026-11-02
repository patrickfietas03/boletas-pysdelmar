import os
import tempfile
import secrets
import time
import pandas as pd
from datetime import timedelta
from pathlib import Path
from flask import session, abort
from almacen import obtener_clave, verificar_usuario, guardar_planilla, cargar_planilla, historial_usuario
from validaciones import revisar_excel, asegurar_estructura, revisar_formulas_sin_resultado
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    url_for,
    send_from_directory,
    send_file
)


app = Flask(__name__)


app.secret_key = obtener_clave()
app.config.update(
    MAX_CONTENT_LENGTH=15*1024*1024,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=10),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Strict',
    SESSION_COOKIE_SECURE=os.getenv('BOLETAS_HTTPS', '0') == '1',
)
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = str(BASE_DIR / 'uploads')  # Solo logo.jpg y firma.jpg


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def numero(valor):

    if pd.isna(valor) or valor == "":
        return 0.0

    try:

        if isinstance(valor, str):
            valor = valor.replace(",", "").strip()

        return float(valor)

    except Exception:

        return 0.0


def texto(valor):

    if pd.isna(valor):
        return ""

    return str(valor).strip()


def limpiar_dni(valor):

    if pd.isna(valor):
        return ""

    texto_dni = str(valor).strip()

    if texto_dni.endswith(".0"):
        texto_dni = texto_dni[:-2]

    return texto_dni


def agregar_area(lista, area):

    area = texto(area)

    if area and area not in lista:

        lista.append(area)


# ============================================================
# FILTRO PARA MOSTRAR NÚMEROS SIN CEROS DECIMALES
# ============================================================
#
# Ejemplos:
#
# 120.00000  -> 120
# 120.50000  -> 120.5
# 0.56880    -> 0.5688
# 0.62000    -> 0.62
#
# IMPORTANTE:
# No usamos "%.0f" porque un AQ = 0.5688 terminaría como 1.
# ============================================================

@app.template_filter("numero_limpio")
def numero_limpio(valor):

    try:

        numero_valor = float(valor)

        if numero_valor.is_integer():

            return str(
                int(numero_valor)
            )

        return (
            f"{numero_valor:.10f}"
            .rstrip("0")
            .rstrip(".")
        )

    except Exception:

        return ""


# ============================================================
# LECTURA DEL EXCEL
# ============================================================

def leer_excel(ruta):

    df = pd.read_excel(
        ruta,
        sheet_name=0,
        header=None,
        dtype={3: str}
    )


    semana = (

        texto(
            df.iloc[1, 3]
        )

        if len(df.columns) > 3

        else ""

    )


    periodo = (

        texto(
            df.iloc[1, 4]
        )

        if len(df.columns) > 4

        else ""

    )


    datos = df.iloc[5:].copy()


    datos.columns = range(
        len(datos.columns)
    )


    trabajadores = []


    for _, fila in datos.iterrows():

        # ----------------------------------------------------
        # DATOS GENERALES
        # ----------------------------------------------------

        nombre = texto(
            fila.iloc[1]
        )


        tipo_documento = texto(
            fila.iloc[2]
        )


        dni = limpiar_dni(
            fila.iloc[3]
        )


        regimen = texto(
            fila.iloc[4]
        )


        tipo_regimen = texto(
            fila.iloc[5]
        )


        area = texto(
            fila.iloc[6]
        )


        if not dni:

            continue


        # ----------------------------------------------------
        # DÍAS DE LA SEMANA
        # ----------------------------------------------------

        dias_semana = [

            numero(
                fila.iloc[i]
            )

            for i in range(7, 14)

        ]


        # ----------------------------------------------------
        # TOTALES
        # ----------------------------------------------------

        total_tn = numero(
            fila.iloc[14]
        )


        total_horas = numero(
            fila.iloc[15]
        )


        dias_trabajados = numero(
            fila.iloc[16]
        )


        # ----------------------------------------------------
        # HORAS
        # ----------------------------------------------------

        horas_estandar = numero(
            fila.iloc[17]
        )


        horas_25 = numero(
            fila.iloc[18]
        )


        horas_35 = numero(
            fila.iloc[19]
        )


        horas_100 = numero(
            fila.iloc[20]
        )


        # ----------------------------------------------------
        # INGRESOS
        # ----------------------------------------------------

        basica = numero(
            fila.iloc[21]
        )


        asignacion_familiar = numero(
            fila.iloc[22]
        )


        dominical = numero(
            fila.iloc[23]
        )


        gratificacion = numero(
            fila.iloc[24]
        )


        bono = numero(
            fila.iloc[25]
        )


        vacaciones = numero(
            fila.iloc[26]
        )


        cts = numero(
            fila.iloc[27]
        )


        monto_25 = numero(
            fila.iloc[28]
        )


        monto_35 = numero(
            fila.iloc[29]
        )


        monto_100 = numero(
            fila.iloc[30]
        )


        prestamos_adelantos = numero(
            fila.iloc[31]
        )


        condicion_trabajo = numero(
            fila.iloc[32]
        )


        # ----------------------------------------------------
        # DESCUENTOS
        # ----------------------------------------------------

        fondo_spp = numero(
            fila.iloc[33]
        )


        prima_spp = numero(
            fila.iloc[34]
        )


        comision_spp = numero(
            fila.iloc[35]
        )


        descuento_prestamos = numero(
            fila.iloc[36]
        )


        descuentos_comida = numero(
            fila.iloc[37]
        )


        descuentos_varios = numero(
            fila.iloc[38]
        )


        # ----------------------------------------------------
        # NETO / ESSALUD
        # ----------------------------------------------------

        neto = numero(
            fila.iloc[39]
        )


        essalud = numero(
            fila.iloc[40]
        )


        # ====================================================
        # AP / AQ
        #
        # AP = REMUNERACIÓN
        # AQ = VALOR DE TONELADA
        #
        # AP = índice 41
        # AQ = índice 42
        # ====================================================

        remuneracion_area = 0.0


        if len(fila) > 41:

            remuneracion_area = numero(
                fila.iloc[41]
            )


        valor_tn_area = 0.0


        if len(fila) > 42:

            valor_tn_area = numero(
                fila.iloc[42]
            )


        # ====================================================
        # TIPO DE TRABAJO DEL REGISTRO / ÁREA
        # ====================================================

        tiene_tn_area = (
            total_tn > 0
        )


        tiene_horas_area = (
            total_horas > 0
        )


        if (
            tiene_tn_area
            and tiene_horas_area
        ):

            tipo_trabajo_area = (
                "DESTAJO + HORAS"
            )

        elif tiene_tn_area:

            tipo_trabajo_area = "DESTAJO"

        elif tiene_horas_area:

            tipo_trabajo_area = "HORAS"

        else:

            tipo_trabajo_area = "OTRO"


        # ----------------------------------------------------
        # REGISTRO
        # ----------------------------------------------------

        trabajadores.append({

            "nombre":
                nombre,

            "tipo_documento":
                tipo_documento,

            "dni":
                dni,

            "regimen":
                regimen,

            "tipo_regimen":
                tipo_regimen,

            "area":
                area,

            "areas":
                [area]
                if area
                else [],

            # AP
            "remuneracion_area":
                remuneracion_area,

            # AQ
            "valor_tn_area":
                valor_tn_area,

            # Tipo de trabajo de esta fila/área
            "tiene_tn_area":
                tiene_tn_area,

            "tiene_horas_area":
                tiene_horas_area,

            "tipo_trabajo_area":
                tipo_trabajo_area,

            "dias_semana":
                dias_semana,

            "total_tn":
                total_tn,

            "total_horas":
                total_horas,

            "dias_trabajados":
                dias_trabajados,

            "horas_estandar":
                horas_estandar,

            "horas_25":
                horas_25,

            "horas_35":
                horas_35,

            "horas_100":
                horas_100,

            "basica":
                basica,

            "asignacion_familiar":
                asignacion_familiar,

            "dominical":
                dominical,

            "gratificacion":
                gratificacion,

            "bono":
                bono,

            "vacaciones":
                vacaciones,

            "cts":
                cts,

            "monto_25":
                monto_25,

            "monto_35":
                monto_35,

            "monto_100":
                monto_100,

            "prestamos_adelantos":
                prestamos_adelantos,

            "condicion_trabajo":
                condicion_trabajo,

            "fondo_spp":
                fondo_spp,

            "prima_spp":
                prima_spp,

            "comision_spp":
                comision_spp,

            "descuento_prestamos":
                descuento_prestamos,

            "descuentos_comida":
                descuentos_comida,

            "descuentos_varios":
                descuentos_varios,

            "neto":
                neto,

            "essalud":
                essalud

        })


    return (
        semana,
        periodo,
        trabajadores
    )


# ============================================================
# CONSOLIDAR TRABAJADORES
# ============================================================

def consolidar_trabajadores(
    trabajadores
):

    grupos = {}


    for trabajador in trabajadores:

        dni = trabajador["dni"]


        # ====================================================
        # CREAR TRABAJADOR
        # ====================================================

        if dni not in grupos:

            grupos[dni] = {

                "nombre":
                    trabajador["nombre"],

                "tipo_documento":
                    trabajador["tipo_documento"],

                "dni":
                    dni,

                "regimen":
                    trabajador["regimen"],

                "tipo_regimen":
                    trabajador["tipo_regimen"],

                "areas":
                    [],

                "detalle_areas":
                    [],

                "dias_semana":
                    [0] * 7,

                "total_tn":
                    0,

                "total_horas":
                    0,

                "dias_trabajados":
                    0,

                "horas_estandar":
                    0,

                "horas_25":
                    0,

                "horas_35":
                    0,

                "horas_100":
                    0,

                "basica":
                    0,

                "asignacion_familiar":
                    0,

                "dominical":
                    0,

                "gratificacion":
                    0,

                "bono":
                    0,

                "vacaciones":
                    0,

                "cts":
                    0,

                "monto_25":
                    0,

                "monto_35":
                    0,

                "monto_100":
                    0,

                "prestamos_adelantos":
                    0,

                "condicion_trabajo":
                    0,

                "fondo_spp":
                    0,

                "prima_spp":
                    0,

                "comision_spp":
                    0,

                "descuento_prestamos":
                    0,

                "descuentos_comida":
                    0,

                "descuentos_varios":
                    0,

                "neto":
                    0,

                "essalud":
                    0

            }


        grupo = grupos[dni]


        # ====================================================
        # ÁREAS
        # ====================================================

        agregar_area(
            grupo["areas"],
            trabajador["area"]
        )


        # ====================================================
        # DETALLE POR ÁREA
        #
        # REGLA:
        #
        # Solo destajo:
        #   muestra VALOR TN
        #
        # Solo horas:
        #   muestra REMUNERACIÓN
        #
        # Ambos:
        #   muestra REMUNERACIÓN + VALOR TN
        # ====================================================

        area_detalle = (
            trabajador["area"]
            if trabajador["area"]
            else "SIN ÁREA"
        )


        tiene_tn = (
            trabajador["tiene_tn_area"]
        )


        tiene_horas = (
            trabajador["tiene_horas_area"]
        )


        remuneracion_detalle = (
            trabajador["remuneracion_area"]
        )


        valor_tn_detalle = (
            trabajador["valor_tn_area"]
        )


        # No agregamos registros sin
        # ningún tipo de trabajo.

        if (
            tiene_tn
            or tiene_horas
        ):

            detalle_encontrado = None


            # ------------------------------------------------
            # BUSCAR MISMA ÁREA + MISMO TIPO + MISMO VALOR TN
            # ------------------------------------------------

            for detalle in grupo[
                "detalle_areas"
            ]:

                if (

                    detalle["area"]
                    == area_detalle

                    and

                    detalle["tiene_tn"]
                    == tiene_tn

                    and

                    detalle["tiene_horas"]
                    == tiene_horas

                    and

                    abs(
                        detalle["valor_tn"]
                        - valor_tn_detalle
                    ) < 0.0000001

                ):

                    detalle_encontrado = detalle

                    break


            # ------------------------------------------------
            # SI YA EXISTE, ACUMULAR REMUNERACIÓN
            # ------------------------------------------------

            if detalle_encontrado:

                # AP es remuneración referencial, no un ingreso por fila.
                # No se debe acumular cuando el área se repite.
                if not detalle_encontrado['remuneracion'] and remuneracion_detalle:
                    detalle_encontrado['remuneracion'] = remuneracion_detalle


                detalle_encontrado[
                    "total_tn"
                ] += trabajador["total_tn"]


                detalle_encontrado[
                    "total_horas"
                ] += trabajador["total_horas"]


            # ------------------------------------------------
            # SI NO EXISTE, CREAR REGISTRO
            # ------------------------------------------------

            else:

                grupo[
                    "detalle_areas"
                ].append({

                    "area":
                        area_detalle,

                    "remuneracion":
                        remuneracion_detalle,

                    # Cantidades trabajadas en esta área
                    "total_tn":
                        trabajador["total_tn"],

                    "total_horas":
                        trabajador["total_horas"],

                    "valor_tn":
                        valor_tn_detalle,

                    "tiene_tn":
                        tiene_tn,

                    "tiene_horas":
                        tiene_horas,

                    "tipo_trabajo":
                        (
                            "DESTAJO + HORAS"
                            if (
                                tiene_tn
                                and tiene_horas
                            )
                            else
                            "DESTAJO"
                            if tiene_tn
                            else
                            "HORAS"
                        )

                })


        # ====================================================
        # DATOS GENERALES
        # ====================================================

        if not grupo["nombre"]:

            grupo["nombre"] = (
                trabajador["nombre"]
            )


        if not grupo["tipo_documento"]:

            grupo["tipo_documento"] = (
                trabajador[
                    "tipo_documento"
                ]
            )


        if not grupo["regimen"]:

            grupo["regimen"] = (
                trabajador["regimen"]
            )


        if not grupo["tipo_regimen"]:

            grupo["tipo_regimen"] = (
                trabajador["tipo_regimen"]
            )


        # ====================================================
        # DÍAS
        # ====================================================

        for i in range(7):

            grupo["dias_semana"][i] += (
                trabajador[
                    "dias_semana"
                ][i]
            )

        # ====================================================
        # TONELADAS / HORAS
        # ====================================================

        grupo["total_tn"] += (
            trabajador["total_tn"]
        )


        grupo["total_horas"] += (
            trabajador["total_horas"]
        )


        grupo["horas_estandar"] += (
            trabajador[
                "horas_estandar"
            ]
        )


        grupo["horas_25"] += (
            trabajador["horas_25"]
        )


        grupo["horas_35"] += (
            trabajador["horas_35"]
        )


        grupo["horas_100"] += (
            trabajador["horas_100"]
        )


        # ====================================================
        # DÍAS TRABAJADOS
        # ====================================================

        # La columna Q contiene días informados por registro/área.
        # Si el documento se repite, se SUMAN todas sus filas,
        # incluso cuando dos áreas correspondan al mismo día.
        # No sustituimos Q por los días únicos de H:N.
        grupo["dias_trabajados"] += (
            trabajador["dias_trabajados"]
        )


        # ====================================================
        # INGRESOS
        # ====================================================

        grupo["basica"] += (
            trabajador["basica"]
        )


        grupo["asignacion_familiar"] += (
            trabajador[
                "asignacion_familiar"
            ]
        )


        grupo["dominical"] += (
            trabajador["dominical"]
        )


        grupo["gratificacion"] += (
            trabajador["gratificacion"]
        )


        grupo["bono"] += (
            trabajador["bono"]
        )


        grupo["vacaciones"] += (
            trabajador["vacaciones"]
        )


        grupo["cts"] += (
            trabajador["cts"]
        )


        grupo["monto_25"] += (
            trabajador["monto_25"]
        )


        grupo["monto_35"] += (
            trabajador["monto_35"]
        )


        grupo["monto_100"] += (
            trabajador["monto_100"]
        )


        grupo["prestamos_adelantos"] += (
            trabajador[
                "prestamos_adelantos"
            ]
        )


        grupo["condicion_trabajo"] += (
            trabajador[
                "condicion_trabajo"
            ]
        )


        # ====================================================
        # DESCUENTOS
        # ====================================================

        grupo["fondo_spp"] += (
            trabajador["fondo_spp"]
        )


        grupo["prima_spp"] += (
            trabajador["prima_spp"]
        )


        grupo["comision_spp"] += (
            trabajador["comision_spp"]
        )


        grupo["descuento_prestamos"] += (
            trabajador[
                "descuento_prestamos"
            ]
        )


        grupo["descuentos_comida"] += (
            trabajador[
                "descuentos_comida"
            ]
        )


        grupo["descuentos_varios"] += (
            trabajador[
                "descuentos_varios"
            ]
        )


        # ====================================================
        # NETO / ESSALUD
        # ====================================================

        grupo["neto"] += (
            trabajador["neto"]
        )


        grupo["essalud"] += (
            trabajador["essalud"]
        )


    # ========================================================
    # RESULTADO FINAL
    # ========================================================

    resultado = []


    for trabajador in grupos.values():

        tiene_toneladas = (
            trabajador["total_tn"] > 0
        )


        tiene_horas = (
            trabajador["total_horas"] > 0
        )


        if (
            tiene_toneladas
            and tiene_horas
        ):

            tipo_trabajo = (
                "DESTAJO + HORAS"
            )

        elif tiene_toneladas:

            tipo_trabajo = "DESTAJO"

        elif tiene_horas:

            tipo_trabajo = "HORAS"

        else:

            tipo_trabajo = "OTRO"


        # Distribuir según el espacio estimado: una fila mixta ocupa más
        # alto; los nombres largos pueden necesitar dos líneas.
        presupuesto = 6
        principales = []
        adicionales = []
        for detalle in trabajador['detalle_areas']:
            costo = 1
            if detalle['tiene_tn'] and detalle['tiene_horas']:
                costo += 1
            if len(detalle['area']) > 28:
                costo += 1
            if costo <= presupuesto and not adicionales:
                principales.append(detalle)
                presupuesto -= costo
            else:
                adicionales.append(detalle)
        trabajador['detalle_areas_principal'] = principales
        trabajador['detalle_areas_anexos'] = [
            adicionales[i:i+18] for i in range(0, len(adicionales), 18)
        ]

        trabajador[
            "tipo_trabajo"
        ] = tipo_trabajo


        trabajador[
            "areas_texto"
        ] = " / ".join(
            trabajador["areas"]
        )


        resultado.append(
            trabajador
        )


    return resultado


# ============================================================
# ACCESO / SESIONES
# ============================================================

_INTENTOS = {}

@app.before_request
def comprobar_acceso():
    # Archivos estáticos de interfaz no contienen información salarial.
    if request.endpoint in ('static', 'login'):
        return
    if not session.get('usuario_id'):
        return redirect(url_for('login', siguiente=request.path))


@app.after_request
def cabeceras_privadas(respuesta):
    respuesta.headers['X-Content-Type-Options'] = 'nosniff'
    respuesta.headers['X-Frame-Options'] = 'DENY'
    respuesta.headers['Referrer-Policy'] = 'same-origin'
    if request.endpoint != 'static':
        respuesta.headers['Cache-Control'] = 'no-store, max-age=0'
    return respuesta


def token_csrf():
    if 'csrf' not in session:
        session['csrf'] = secrets.token_urlsafe(28)
    return session['csrf']

app.jinja_env.globals['csrf_token'] = token_csrf


def validar_csrf():
    if not secrets.compare_digest(str(session.get('csrf', '')),
                                  str(request.form.get('csrf_token', ''))) or not session.get('csrf'):
        abort(400, description='La sesión del formulario expiró. Actualiza la página.')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('usuario_id'):
        return redirect(url_for('index'))
    error = ''
    ip = request.remote_addr or 'desconocido'
    if request.method == 'POST':
        validar_csrf()
        ahora = time.time()
        fallas = [t for t in _INTENTOS.get(ip, []) if ahora-t < 600]
        if len(fallas) >= 8:
            error = 'Demasiados intentos. Intenta nuevamente dentro de unos minutos.'
        else:
            usuario = verificar_usuario(request.form.get('usuario',''),
                                       request.form.get('contrasena',''))
            if not usuario:
                fallas.append(ahora)
                _INTENTOS[ip] = fallas
                error = 'Usuario o contraseña incorrectos.'
            else:
                _INTENTOS.pop(ip, None)
                session.clear()
                session.permanent = True
                session['usuario_id'] = usuario['id']
                session['usuario_nombre'] = usuario['usuario']
                return redirect(url_for('index'))
    return render_template('login.html', error=error)


@app.route('/salir', methods=['POST'])
def salir():
    validar_csrf()
    session.clear()
    return redirect(url_for('login'))


def planilla_actual():
    return cargar_planilla(session.get('planilla_id'), session['usuario_id'])


def requerir_planilla():
    data = planilla_actual()
    if not data:
        flash('Selecciona una planilla del historial o importa un Excel.')
    return data


# ============================================================
# NAVEGACIÓN / HISTORIAL
# ============================================================

@app.route('/')
def index():
    return render_template('index.html', usuario=session.get('usuario_nombre'))


# ============================================================
# DESCARGAR PLANTILLA EXCEL (MISMA ESTRUCTURA A:AQ)
# ============================================================

@app.route('/descargar-plantilla', methods=['GET'])
def descargar_plantilla():
    # La comprobación de sesión se realiza en comprobar_acceso().
    # Archivo sin datos personales: conserva encabezados 4 y 5,
    # semana en D2, período en E2 y registros desde la fila 6.
    ruta = BASE_DIR / 'plantillas' / 'PLANTILLA_BOLETAS_PYS_DEL_MAR.xlsx'
    return send_file(
        ruta,
        as_attachment=True,
        download_name='PLANTILLA_BOLETAS_PYS_DEL_MAR.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/historial')
def historial():
    return render_template('historial.html', planillas=historial_usuario(session['usuario_id']),
                           actual=session.get('planilla_id'))


@app.route('/historial/<int:planilla_id>/abrir', methods=['POST'])
def abrir_planilla(planilla_id):
    validar_csrf()
    if not cargar_planilla(planilla_id, session['usuario_id']):
        abort(404)
    session['planilla_id'] = planilla_id
    return redirect(url_for('resultados'))


# ============================================================
# PROCESAMIENTO CON VALIDACIÓN NO DESTRUCTIVA
# ============================================================

@app.route('/procesar', methods=['POST'])
def procesar():
    validar_csrf()
    archivo = request.files.get('excel_destajo')
    if not archivo or not archivo.filename:
        flash('Selecciona un archivo Excel.')
        return redirect(url_for('index'))
    extension = Path(archivo.filename).suffix.lower()
    if extension != '.xlsx':
        flash('Solo se permiten archivos .xlsx. Guarda tu plantilla en ese formato.')
        return redirect(url_for('index'))
    # Se utiliza un archivo temporal único para evitar colisiones de nombre.
    ruta = None
    try:
        with tempfile.NamedTemporaryFile(prefix='boleta_', suffix='.xlsx',
                                         delete=False) as temporal:
            ruta = temporal.name
            archivo.save(temporal)
        df = pd.read_excel(ruta, sheet_name=0, header=None, dtype={3: str})
        asegurar_estructura(df)
        semana, periodo, registros = leer_excel(ruta)
        if not registros:
            raise ValueError('El archivo no contiene trabajadores con documento.')
        consolidados = consolidar_trabajadores(registros)
        alertas = revisar_excel(df, registros, consolidados)
        alertas.extend(revisar_formulas_sin_resultado(ruta, df))
        identificador = guardar_planilla(session['usuario_id'], semana, periodo,
                                         Path(archivo.filename).name, consolidados, alertas)
        session['planilla_id'] = identificador
        flash(f'Se importaron {len(consolidados)} trabajadores. Alertas de revisión: {len(alertas)}.')
        return redirect(url_for('validaciones'))
    except Exception as exc:
        # No exponemos tracebacks, rutas de servidor ni información personal.
        flash('No se pudo procesar el Excel: ' + str(exc)[:250])
        return redirect(url_for('index'))
    finally:
        if ruta:
            try: os.unlink(ruta)
            except FileNotFoundError: pass


@app.route('/validaciones')
def validaciones():
    datos = requerir_planilla()
    if not datos:
        return redirect(url_for('historial'))
    return render_template('validaciones.html', datos=datos)


@app.route('/resultados')
def resultados():
    datos = requerir_planilla()
    if not datos:
        return redirect(url_for('historial'))
    return render_template('resultados.html', semana=datos['semana'],
                           periodo=datos['periodo'], trabajadores=datos['trabajadores'],
                           alertas=datos['alertas'])


@app.route('/trabajadores')
def trabajadores():
    datos = requerir_planilla()
    if not datos:
        return redirect(url_for('historial'))
    return render_template('trabajadores.html', datos=datos)


@app.route('/boleta/<dni>')
def boleta(dni):
    datos = requerir_planilla()
    if not datos:
        return redirect(url_for('historial'))
    if any(a['nivel'] == 'error' for a in datos['alertas']):
        flash('La planilla tiene errores de datos. Corrige el Excel antes de emitir.')
        return redirect(url_for('validaciones'))
    trabajador = next((t for t in datos['trabajadores'] if str(t['dni']) == str(dni)), None)
    if not trabajador:
        abort(404)
    return render_template('boleta.html', trabajador=trabajador,
                           semana=datos['semana'], periodo=datos['periodo'])


@app.route('/boletas/masivas')
def boletas_masivas():
    datos = requerir_planilla()
    if not datos:
        return redirect(url_for('historial'))
    if any(a['nivel'] == 'error' for a in datos['alertas']):
        flash('La planilla tiene errores de datos. Corrige el Excel antes de emitir.')
        return redirect(url_for('validaciones'))
    return render_template('boletas_masivas.html', trabajadores=datos['trabajadores'],
                           semana=datos['semana'], periodo=datos['periodo'])


@app.route('/uploads/<filename>')
def serve_uploads(filename):
    if filename not in ('logo.jpg', 'firma.jpg'):
        abort(404)
    return send_from_directory(UPLOAD_FOLDER, filename)


if __name__ == '__main__':
    app.run(debug=False, host=os.environ.get('BOLETAS_HOST', '127.0.0.1'),
            port=int(os.environ.get('PORT', '5000')))
