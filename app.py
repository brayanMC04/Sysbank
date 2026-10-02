import pandas as pd
import io
from flask import Flask, render_template, request, redirect, url_for, send_file, session
import pyodbc
from flask import send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from decimal import Decimal


app = Flask(__name__)
app.secret_key = 'clave_secreta_sysbank_digital'

# Configuración de tu conexión a SQL Server
conexion_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;"
    "DATABASE=SysbankDigital;"
    "UID=brayan;"
    "PWD=123123"
)

# ... (tu cadena de conexión y otras rutas se mantienen igual) ...

@app.route('/query', methods=['GET', 'POST'])
def query_interface():
    resultado = None
    columnas = []
    error = None
    
    if request.method == 'POST':
        sql_query = request.form.get('query')
        # Guardamos la consulta en la sesión para poder exportarla a Excel después
        session['last_query'] = sql_query
        
        try:
            conexion = pyodbc.connect(conexion_str)
            cursor = conexion.cursor()
            cursor.execute(sql_query)
            
            if sql_query.strip().upper().startswith("SELECT"):
                columnas = [column[0] for column in cursor.description]
                resultado = cursor.fetchall()
            else:
                conexion.commit()
                resultado = "Consulta ejecutada exitosamente (Cambios aplicados en la BD)."
                
            cursor.close()
            conexion.close()
        except Exception as e:
            error = str(e)
            
    return render_template('query.html', resultado=resultado, columnas=columnas, error=error)

# Nueva ruta para exportar el resultado actual a Excel
@app.route('/descargar-excel')
def descargar_excel():
    try:
        sql_query = session.get('last_query')
        if not sql_query or not sql_query.strip().upper().startswith("SELECT"):
            return "No hay una consulta SELECT previa para exportar.", 400
            
        # Conectamos y leemos directamente con pandas
        conexion = pyodbc.connect(conexion_str)
        df = pd.read_sql(sql_query, conexion)
        conexion.close()
        
        # Guardamos el DataFrame en un buffer de memoria en formato Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado SQL')
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name="resultado_consulta_sysbank.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return f"Error al generar el archivo Excel: {e}"

@app.route('/descargar-extracto')
def descargar_extracto():
    try:
        # Creamos un buffer en memoria para el PDF
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=letter)
        
        # Cabecera del PDF
        pdf.setFont("Helvetica-Bold", 20)
        pdf.drawString(50, 750, "SYSBANK DIGITAL - EXTRACTO BANCARIO")
        
        pdf.setFont("Helvetica", 12)
        pdf.drawString(50, 725, "Resumen oficial de transacciones y cuentas")
        pdf.line(50, 715, 560, 715)
        
        # Consultamos datos reales de la BD para el reporte
        conexion = pyodbc.connect(conexion_str)
        cursor = conexion.cursor()
        
        cursor.execute("SELECT TOP 10 ID_Movimiento, ID_CuentaOrigen, Valor, FechaMvto FROM MOVIMIENTO ORDER BY FechaMvto DESC")
        filas = cursor.fetchall()
        
        y = 670
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(50, y, "ID Movimiento")
        pdf.drawString(150, y, "Cuenta Origen")
        pdf.drawString(300, y, "Valor ($)")
        pdf.drawString(450, y, "Fecha")
        
        pdf.setFont("Helvetica", 10)
        for fila in filas:
            y -= 25
            pdf.drawString(50, y, str(fila[0]))
            pdf.drawString(150, y, str(fila[1]))
            pdf.drawString(300, y, f"${fila[2]:,.2f}")
            pdf.drawString(450, y, str(fila[3]))
            
        cursor.close()
        conexion.close()
        
        pdf.save()
        buffer.seek(0)
        
        return send_file(buffer, as_attachment=True, download_name="extracto_sysbank.pdf", mimetype='application/pdf')
    except Exception as e:
        return f"Error al generar el PDF: {e}"

# Ruta del Login (Muestra la vista y procesa el formulario para ir al dashboard)
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        # Aquí puedes validar usuario/contraseña si lo deseas más adelante
        return redirect(url_for('dashboard'))
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    try:
        conexion = pyodbc.connect(conexion_str)
        cursor = conexion.cursor()
        
        # 1. Calcular el saldo total disponible sumando las cuentas
        cursor.execute("SELECT SUM(Saldo) FROM CUENTA")
        resultado_saldo = cursor.fetchone()
        saldo_total = resultado_saldo[0] if resultado_saldo and resultado_saldo[0] is not None else 0.00
        
        # 2. Obtener la lista de cuentas
        cursor.execute("SELECT Numero_Cuenta, Saldo, Estado FROM CUENTA")
        cuentas = cursor.fetchall()
        
        # 3. Obtener los movimientos recientes haciendo JOIN con TIPO_MOVIMIENTO
        query_movimientos = """
            SELECT TOP 5 
                m.ID_Movimiento, 
                m.ID_CuentaOrigen, 
                t.NombreTipoMvto, 
                m.Valor, 
                m.FechaMvto 
            FROM MOVIMIENTO m
            JOIN TIPO_MOVIMIENTO t ON m.ID_TipoMvto = t.ID_TipoMvto
            ORDER BY m.FechaMvto DESC
        """
        cursor.execute(query_movimientos)
        
        # Convertimos cada fila a un diccionario limpio y compatible con JSON
        movimientos_raw = cursor.fetchall()
        movimientos = []
        for row in movimientos_raw:
            movimientos.append({
                'id': row[0],
                'cuenta': row[1],
                'tipo': row[2],
                'valor': float(row[3]) if row[3] is not None else 0.0,
                'fecha': str(row[4]) if row[4] is not None else ''
            })
        
        cursor.close()
        conexion.close()
        
        return render_template('dashboard.html', saldo_total=saldo_total, cuentas=cuentas, movimientos=movimientos)
    except Exception as e:
        return f"Error al cargar el dashboard: {e}"

@app.route('/transaccion', methods=['GET', 'POST'])
def transaccion():
    mensaje = None
    error = None
    
    conexion = pyodbc.connect(conexion_str)
    cursor = conexion.cursor()
    
    # Obtenemos las cuentas disponibles para mostrarlas en el formulario
    cursor.execute("SELECT Numero_Cuenta, Saldo FROM CUENTA WHERE Estado = 'Activa'")
    cuentas = cursor.fetchall()
    
    # Obtenemos los tipos de movimiento (ej: Depósito, Retiro, Transferencia)
    cursor.execute("SELECT ID_TipoMvto, NombreTipoMvto FROM TIPO_MOVIMIENTO")
    tipos_movimiento = cursor.fetchall()
    
    if request.method == 'POST':
        cuenta_origen = request.form.get('cuenta_origen')
        tipo_mvto = int(request.form.get('tipo_mvto'))
        valor = float(request.form.get('valor'))
        descripcion = request.form.get('descripcion', 'Transacción web')
        valor = Decimal(request.form.get('valor'))
        descripcion = request.form.get('descripcion', 'Transacción web')
        
        try:
            # Validar saldo si es un retiro o salida
            cursor.execute("SELECT Saldo FROM CUENTA WHERE Numero_Cuenta = ?", (cuenta_origen,))
            row_cuenta = cursor.fetchone()
            if not row_cuenta:
                raise Exception("La cuenta de origen no existe.")
            
            saldo_actual = row_cuenta[0]
            
            # Asumiremos tipos de movimiento comunes (ej: 1 = Deposito/Consignacion, 2 = Retiro, etc. Ajusta según tus IDs de SQL)
            # Validamos si requiere fondos suficientes (ejemplo básico para retiros)
            if tipo_mvto == 2 and saldo_actual < valor: # Ajusta el ID según tu tabla TIPO_MOVIMIENTO
                raise Exception("Fondos insuficientes para realizar el retiro.")
            
            # Insertar el movimiento
            query_insert = """
                INSERT INTO MOVIMIENTO (ID_TipoMvto, ID_CuentaOrigen, ID_CuentaDestino, FechaMvto, Valor, Descripcion)
                VALUES (?, ?, NULL, GETDATE(), ?, ?)
            """
            cursor.execute(query_insert, (tipo_mvto, cuenta_origen, valor, descripcion))
            
            # Actualizar el saldo de la cuenta
            if tipo_mvto == 1: # Si es depósito/consignación suma
                nuevo_saldo = saldo_actual + valor
            else: # Si es retiro resta
                nuevo_saldo = saldo_actual - valor
                
            cursor.execute("UPDATE CUENTA SET Saldo = ?, FechaUltmMovimiento = GETDATE() WHERE Numero_Cuenta = ?", (nuevo_saldo, cuenta_origen))
            
            conexion.commit()
            mensaje = "¡Transacción realizada con éxito!"
        except Exception as e:
            conexion.rollback()
            error = str(e)
            
    cursor.close()
    conexion.close()
    
    return render_template('transaccion.html', cuentas=cuentas, tipos_movimiento=tipos_movimiento, mensaje=mensaje, error=error)

if __name__ == '__main__':
    app.run(debug=True)