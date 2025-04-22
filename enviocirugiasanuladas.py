from sqlalchemy import create_engine
import pandas as pd
from datetime import datetime
import smtplib
from email.message import EmailMessage
import os

# Datos de conexión a MySQL con SQLAlchemy
usuario = 'luisp'
clave = 'iopa2023$'
host = '192.168.1.18'
base_datos = 'tablaquirurgica'

engine = create_engine(f'mysql+pymysql://{usuario}:{clave}@{host}/{base_datos}')

query = """
SELECT
    a.id_bloque AS id_bloque,
    a.idsolicitud AS idsolicitud,
    a.estado AS codigo_estado,
    CASE
        WHEN a.estado = 9 THEN 'Anulado por cantidad de llamados'
        ELSE 'Anulado'
    END AS estado,
    a.sucursal,
    CASE
        WHEN a.sucursal = 1 THEN 'Los Leones'
        WHEN a.sucursal = 2 THEN 'La Florida'
        WHEN a.sucursal = 3 THEN 'Huerfanos'
        ELSE ''
    END AS nombre_sucursal,
    s.fechasp AS fechasp,
    s.rut AS rut_paciente,
    s.nombre AS nombre_paciente,
    s.sexo,
    s.correo_electronico,
    s.telefono_trabajo AS telefono_trabajo,
    a.primer_cirujano AS medico_rut,
    CONCAT_WS(" ", m.name, m.lastname) AS medico_nombre,
    s.cirugia AS procedimiento,
    a.fecha AS fecha_cirugia,
    a.desde,
    a.hasta,
    s.codigo_presupuesto,
    s.nombre_procedencia
FROM agenda a
LEFT JOIN solicitud AS s ON s.id = a.idsolicitud
INNER JOIN pabellon AS p ON a.sala = p.id_pabellon
LEFT JOIN medico AS m ON m.rut_num = a.primer_cirujano
WHERE
    a.pos = 1
    AND a.estado IN (3, 9)
    AND a.fecha = CURDATE() - INTERVAL 1 DAY;
"""

# Leer resultados con pandas y SQLAlchemy
df = pd.read_sql(query, engine)

# Guardar en Excel
fecha_hoy = datetime.now().strftime('%Y-%m-%d')
excel_path = f'anulados_{fecha_hoy}.xlsx'
df.to_excel(excel_path, index=False)

# Lista de destinatarios
destinatarios = ['nstuardo@gmail.com', 'yeremi.ortega@iopa.cl']

# Configuración de email
msg = EmailMessage()
msg['Subject'] = f'Reporte de anulados del {fecha_hoy}'
msg['From'] = 'envios@iopa.cl'
msg['To'] = ", ".join(destinatarios)
msg.set_content('Adjunto encontrarás el reporte en Excel.')

# Adjuntar archivo
with open(excel_path, 'rb') as f:
    file_data = f.read()
    msg.add_attachment(
        file_data,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=os.path.basename(excel_path)
    )

# Enviar por SMTP normal (no Gmail)
smtp_host = 'mail.iopa.cl'
smtp_port = 587  # O 25 si el servidor no acepta TLS
smtp_user = 'envios@iopa.cl'
smtp_pass = '$NSloteria2015$'

with smtplib.SMTP(smtp_host, smtp_port) as server:
    server.starttls()  # TLS para puertos 587 o 25
    server.login(smtp_user, smtp_pass)
    server.send_message(msg)

print("✅ Correo enviado con éxito.")
