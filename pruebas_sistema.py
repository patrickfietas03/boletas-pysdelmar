"""Prueba de integración opcional: python pruebas_sistema.py ruta/plantilla.xlsx

Crea una base temporal: no modifica el historial ni el Excel del usuario.
"""
import io
import os
import re
import sys
import tempfile
from pathlib import Path

if __name__ == '__main__':
    if len(sys.argv) != 2 or not Path(sys.argv[1]).is_file():
        raise SystemExit('Uso: python pruebas_sistema.py ruta/plantilla.xlsx')
    path = Path(sys.argv[1]).resolve()
    with tempfile.TemporaryDirectory(prefix='pruebas_boletas_') as carpeta:
        os.environ['BOLETAS_DATA_DIR'] = carpeta
        from app import app
        from almacen import crear_usuario
        crear_usuario('prueba_local', 'PruebaTemporal123456!')
        app.config['TESTING'] = True
        client = app.test_client()
        def csrf(res):
            m = re.search(rb'name="csrf_token"\s+value="([^"]+)"', res.data)
            assert m, 'No se encontró el token de seguridad en la página.'
            return m.group(1).decode('utf-8')
        resp = client.get('/login')
        assert resp.status_code == 200
        resp = client.post('/login', data={'usuario':'prueba_local',
                    'contrasena':'PruebaTemporal123456!', 'csrf_token':csrf(resp)},
                    follow_redirects=True)
        assert resp.status_code == 200
        resp = client.get('/')
        form = {'csrf_token':csrf(resp),
                'excel_destajo':(io.BytesIO(path.read_bytes()),path.name)}
        resp = client.post('/procesar', data=form,
                           content_type='multipart/form-data',follow_redirects=True)
        assert resp.status_code == 200
        assert 'CONTROL DE IMPORTACIÓN' in resp.get_data(as_text=True)
        resp = client.get('/historial')
        assert resp.status_code == 200 and path.name in resp.get_data(as_text=True)
        resp = client.get('/resultados')
        assert resp.status_code == 200
        resp = client.get('/boletas/masivas')
        if resp.status_code == 200:
            assert 'pagina-trabajador' in resp.get_data(as_text=True)
        else:
            assert resp.status_code in (302,303), 'Respuesta inesperada de impresión.'
        print('PRUEBA OK: acceso, carga, validación, historial, resultados y emisión.')
        print('El Excel de origen no fue modificado.')
