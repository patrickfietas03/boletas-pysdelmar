"""Ejecutar una vez antes de iniciar el sistema: python configurar_usuario.py"""
import getpass
import sqlite3
from almacen import crear_usuario

if __name__ == '__main__':
    print('Configuración de acceso al sistema de boletas')
    nombre = input('Usuario: ').strip()
    password = getpass.getpass('Contraseña (mínimo 12 caracteres): ')
    confirmar = getpass.getpass('Confirmar contraseña: ')
    if password != confirmar:
        raise SystemExit('Las contraseñas no coinciden; no se creó ningún usuario.')
    try:
        crear_usuario(nombre, password)
    except (ValueError, sqlite3.IntegrityError) as exc:
        raise SystemExit(f'No fue posible crear la cuenta: {exc}')
    print('Cuenta creada. Inicia el sistema con: python app.py')
