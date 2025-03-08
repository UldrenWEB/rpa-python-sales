# filepath: c:\Users\uldre\Desktop\Uldren\Desarrollo\Entornos\VSCODE\2025A\AI\rpa (Ventas)\main.py
import os
import json
import pandas as pd
import matplotlib.pyplot as plt
from twilio.rest import Client
import dropbox

with open('credentials.json') as config_file:
    config = json.load(config_file)

account_sid = config['TWILIO_SID']
auth_token = config['TWILIO_AUTH_TOKEN']
client = Client(account_sid, auth_token)

def send_whatsapp_message(body, media_url=None):
    message = client.messages.create(
        to='whatsapp:+584121528916',
        from_='whatsapp:+14155238886',
        body=body,
        media_url=media_url
    )
    print(message)

# Leer datos desde el archivo Excel
file_path = 'Ventas - Fundamentos.xlsx'
ventas_df = pd.read_excel(file_path, sheet_name='VENTAS')
vehiculos_df = pd.read_excel(file_path, sheet_name='VEHICULOS')
nuevos_registros_df = pd.read_excel(file_path, sheet_name='NUEVOS REGISTROS')

# Cálculo de ingresos, costos y beneficios totales
ingresos_totales = ventas_df['Precio Venta Real'].sum()
costos_totales = ventas_df['Costo Vehículo'].sum()
beneficios_totales = ingresos_totales - costos_totales

# Análisis por Canal de Venta
ventas_por_canal = ventas_df.groupby('Canal')['Precio Venta Real'].sum()

# Análisis por Segmento de Clientes
ventas_por_segmento = ventas_df.groupby('Segmento')['Precio Venta Real'].sum()

# Tendencias de Venta Mensuales
ventas_df['Fecha'] = pd.to_datetime(ventas_df['Fecha'])
ventas_df['Mes'] = ventas_df['Fecha'].dt.to_period('M')
ventas_mensuales = ventas_df.groupby('Mes')['Precio Venta Real'].sum()

# Crear la carpeta de salida si no existe
output_folder = 'output'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Crear el DataFrame resumen_financiero
resumen_financiero = pd.DataFrame({
    'Métrica': ['Ingresos Totales', 'Costos Totales', 'Beneficios Totales'],
    'Valor': [ingresos_totales, costos_totales, beneficios_totales]
})

# Guardar el resumen financiero en un archivo Excel
with pd.ExcelWriter(f'{output_folder}/Reporte_Financiero.xlsx') as writer:
    resumen_financiero.to_excel(writer, sheet_name='Resumen Financiero', index=False)
    ventas_por_canal.to_excel(writer, sheet_name='Ventas por Canal')
    ventas_por_segmento.to_excel(writer, sheet_name='Ventas por Segmento')
    ventas_mensuales.to_excel(writer, sheet_name='Ventas Mensuales')

# Guardar los gráficos en archivos locales
plt.figure(figsize=(10, 6))
ventas_mensuales.plot(kind='bar', title='Tendencias de Venta Mensuales')
ventas_mensuales_path = f'{output_folder}/Ventas_Mensuales.png'
plt.savefig(ventas_mensuales_path)

plt.figure(figsize=(10, 6))
ventas_por_canal.plot(kind='bar', title='Ventas por Canal')
ventas_por_canal_path = f'{output_folder}/Ventas_Por_Canal.png'
plt.savefig(ventas_por_canal_path)

plt.figure(figsize=(10, 6))
ventas_por_segmento.plot(kind='bar', title='Ventas por Segmento')
ventas_por_segmento_path = f'{output_folder}/Ventas_Por_Segmento.png'
plt.savefig(ventas_por_segmento_path)

print(f'Archivos guardados en: {output_folder}')

# Funcion utilitarias para poder conectarse a dropbox y hacer las solicitudes correspondientes
dropbox_access_token = config['DROPBOX_TOKEN']
dbx = dropbox.Dropbox(dropbox_access_token)

# Subir archivo a dropbox y obtener la url publica
def upload_to_dropbox(file_path, dropbox_path):
    with open(file_path, 'rb') as f:
        dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite, mute=True)
    shared_link_metadata = dbx.sharing_create_shared_link_with_settings(dropbox_path)
    return shared_link_metadata.url.replace('?dl=0', '?dl=1')

public_url = upload_to_dropbox(f'{output_folder}/Reporte_Financiero.xlsx', '/Reporte_Financiero.xlsx')
ventas_mensuales_url = upload_to_dropbox(ventas_mensuales_path, '/Ventas_Mensuales.png')
ventas_por_canal_url = upload_to_dropbox(ventas_por_canal_path, '/Ventas_Por_Canal.png')
ventas_por_segmento_url = upload_to_dropbox(ventas_por_segmento_path, '/Ventas_Por_Segmento.png')

print(f'Public URL: {public_url}')

# Se envia los reportes via whatsapp
send_whatsapp_message(body=f'Aquí está el reporte financiero: {public_url}')
graficos_ventas_texto = f"""
Gráficos de Ventas:

Ventas mensuales: {ventas_mensuales_url}

Ventas por canal: {ventas_por_canal_url}

Ventas por segmento: {ventas_por_segmento_url}
"""
send_whatsapp_message(body=graficos_ventas_texto)

resultados_texto = f"""
Ingresos Totales: {ingresos_totales:,.2f}
Costos Totales: {costos_totales:,.2f}
Beneficios Totales: {beneficios_totales:,.2f}

Ventas por Canal:
{ventas_por_canal}

Ventas por Segmento:
{ventas_por_segmento}

Ventas Mensuales:
{ventas_mensuales}
"""
send_whatsapp_message(body=resultados_texto)