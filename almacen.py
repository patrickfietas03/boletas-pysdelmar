"""Persistencia privada y cuentas locales del sistema de boletas (SQLite)."""
import json
import os
import secrets
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from werkzeug.security import check_password_hash, generate_password_hash

BASE = Path(__file__).resolve().parent
DATA = Path(os.environ.get('BOLETAS_DATA_DIR', str(BASE / 'data')))
DATA.mkdir(parents=True, exist_ok=True)
DB = DATA / 'boletas.sqlite3'


def obtener_clave():
    variable = os.environ.get('SECRET_KEY')
    if variable and len(variable) >= 24:
        return variable
    ruta = DATA / '.clave_sesion'
    if ruta.exists():
        return ruta.read_text(encoding='utf-8').strip()
    clave = secrets.token_hex(32)
    fd = os.open(str(ruta), os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    with os.fdopen(fd, 'w', encoding='utf-8') as destino:
        destino.write(clave)
    return clave


def conexion():
    con = sqlite3.connect(DB, timeout=20)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys = ON')
    return con


def iniciar():
    with conexion() as con:
        con.execute('''CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE COLLATE NOCASE,
            contrasena_hash TEXT NOT NULL,
            creado_en TEXT NOT NULL
        )''')
        con.execute('''CREATE TABLE IF NOT EXISTS planillas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER NOT NULL REFERENCES usuarios(id),
            semana TEXT NOT NULL,
            periodo TEXT NOT NULL,
            nombre_archivo TEXT NOT NULL,
            creado_en TEXT NOT NULL,
            total_trabajadores INTEGER NOT NULL,
            datos_json TEXT NOT NULL,
            alertas_json TEXT NOT NULL
        )''')
        con.execute('CREATE INDEX IF NOT EXISTS planillas_usuario ON planillas(usuario_id, id DESC)')


def crear_usuario(usuario, contrasena):
    if not usuario or not usuario.strip() or len(contrasena) < 12:
        raise ValueError('Indica un usuario y una contraseña de al menos 12 caracteres.')
    with conexion() as con:
        con.execute('INSERT INTO usuarios (usuario, contrasena_hash, creado_en) VALUES (?,?,?)',
                    (usuario.strip(), generate_password_hash(contrasena),
                     datetime.now(timezone.utc).isoformat(timespec='seconds')))


def verificar_usuario(usuario, contrasena):
    with conexion() as con:
        dato = con.execute('SELECT * FROM usuarios WHERE usuario = ? COLLATE NOCASE',
                           (usuario.strip(),)).fetchone()
        if dato and check_password_hash(dato['contrasena_hash'], contrasena):
            return dict(dato)
        return None


def guardar_planilla(usuario_id, semana, periodo, nombre_archivo, trabajadores, alertas):
    datos = json.dumps(trabajadores, ensure_ascii=False, allow_nan=False)
    avisos = json.dumps(alertas, ensure_ascii=False, allow_nan=False)
    with conexion() as con:
        cursor = con.execute('''INSERT INTO planillas
            (usuario_id, semana, periodo, nombre_archivo, creado_en,
             total_trabajadores, datos_json, alertas_json) VALUES (?,?,?,?,?,?,?,?)''',
            (usuario_id, str(semana), str(periodo), nombre_archivo,
             datetime.now(timezone.utc).isoformat(timespec='seconds'),
             len(trabajadores), datos, avisos))
        return cursor.lastrowid


def cargar_planilla(planilla_id, usuario_id):
    if not planilla_id:
        return None
    with conexion() as con:
        dato = con.execute('SELECT * FROM planillas WHERE id=? AND usuario_id=?',
                           (planilla_id, usuario_id)).fetchone()
    if dato is None:
        return None
    resultado = dict(dato)
    resultado['trabajadores'] = json.loads(resultado.pop('datos_json'))
    resultado['alertas'] = json.loads(resultado.pop('alertas_json'))
    return resultado


def historial_usuario(usuario_id):
    with conexion() as con:
        filas = con.execute('''SELECT id, semana, periodo, nombre_archivo, creado_en,
                                     total_trabajadores, alertas_json
                              FROM planillas WHERE usuario_id=? ORDER BY id DESC''',
                            (usuario_id,)).fetchall()
    return [{**dict(f), 'n_alertas': len(json.loads(f['alertas_json']))} for f in filas]


iniciar()
