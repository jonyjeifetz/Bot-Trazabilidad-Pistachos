import os
import io
import threading
import pandas as pd
from flask import Flask
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from telegram.request import HTTPXRequest

# Configuración de Flask para mantener el servicio activo en la nube
app_flask = Flask(__name__)

@app_flask.route('/')
def home():
    return "¡El bot de trazabilidad y pedidos está activo y funcionando!"

def run_flask():
    port = int(os.environ.get("PORT", 10000))
    app_flask.run(host="0.0.0.0", port=port)

# Token y configuración de archivos
TOKEN = os.environ.get("TELEGRAM_TOKEN")
EXCEL_PATH = "datos_pedidos.xlsx"
# Compatible con ambos entornos (trazabilidad normal o pistachos)
NOMBRE_HOJA = "Pedidos_Pistachos (2)" 

# Memoria global de sesiones y teléfonos verificados (persiste en ejecución)
user_sessions = {}
telefonos_verificados = {} # {user_id: {"telefono": "...", "rol": "...", "nombre_vendedor": "..."}}

# Definición de roles por número de teléfono
ADMIN_PHONES = {
    "541153248379",
    "541156680181",
    "541133635660"
}

VENDEDORES_MAP = {
    # "541112345678": "Nombre Vendedor En Excel"
}

def limpiar_telefono(tel_str):
    if not tel_str:
        return ""
    import re
    return re.sub(r'\D', '', str(tel_str))

def cargar_datos():
    if os.path.exists(EXCEL_PATH):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=NOMBRE_HOJA, engine='openpyxl')
            df.columns = df.columns.astype(str).str.strip()
            
            for col in df.columns:
                df[col] = df[col].astype(str).str.replace('\xa0', ' ', regex=False).str.replace(r'\s+', ' ', regex=True).str.strip()
                df[col] = df[col].str.replace(r'\.0$', '', regex=True)
                
            return df
        except Exception as e:
            print(f"❌ Error al leer la hoja '{NOMBRE_HOJA}' del Excel: {e}")
            return pd.DataFrame()
    else:
        print("⚠️ No se encontró el archivo Excel en la ruta especificada.")
        return pd.DataFrame()

def construir_texto_boton(icon, row, query_texto):
    pedido = row.get('Pedido', 'S/N')
    cliente = row.get('RazonSocial', 'Desconocido')
    vendedor = row.get('Vendedor', 'S/V')
    fecha = row.get('Vtas_Fact_fecha', 'S/F')
    
    if len(str(fecha)) >= 10 and '-' in str(fecha):
        partes_fecha = str(fecha).split()[0].split('-')
        if len(partes_fecha) == 3:
            fecha = f"{partes_fecha[2]}/{partes_fecha[1]}/{partes_fecha[0]}"

    texto = f"{icon} 📅 {fecha} | {cliente} | Ped: {pedido} | {vendedor}"
    if len(texto) > 60:
        texto = texto[:57] + "..."
    return texto

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    # Si ya compartió su contacto alguna vez, lo saludamos humano y directo
    if user_id in telefonos_verificados:
        await mostrar_menu_segun_rol(update, context, user_id)
        return

    contacto_btn = KeyboardButton("📱 Compartir mi contacto (Solo por única vez)", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[contacto_btn]], resize_keyboard=True, one_time_keyboard=True)
    
    await update.message.reply_text(
        "👋 ¡Hola! Bienvenido al sistema de trazabilidad y gestión.\n\n"
        "Para mantener la seguridad y asignarte los permisos correspondientes, por favor compartí tu número de contacto tocando el botón de abajo (solo te lo pediré esta única vez):",
        reply_markup=reply_markup
    )

async def recibir_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message.contact:
        return
        
    telefono = limpiar_telefono(message.contact.phone_number)
    user_id = update.effective_user.id
    
    rol = 'estandar'
    nombre_vendedor = ''
    
    if telefono in ADMIN_PHONES:
        rol = 'admin'
    elif telefono in VENDEDORES_MAP:
        rol = 'vendedor'
        nombre_vendedor = VENDEDORES_MAP[telefono]
        
    telefonos_verificados[user_id] = {
        "telefono": telefono,
        "rol": rol,
        "nombre_vendedor": nombre_vendedor
    }
    
    await message.reply_text(
        "✅ ¡Verificado con éxito! Ya quedó registrado tu acceso seguro.",
        reply_markup=ReplyKeyboardRemove()
    )
    await mostrar_menu_segun_rol(update, context, user_id)

async def mostrar_menu_segun_rol(update_or_query, context, user_id):
    info_usuario = telefonos_verificados.get(user_id, {})
    rol = info_usuario.get('rol', 'estandar')
    
    if rol == 'vendedor':
        await mostrar_menu_principal_vendedor(update_or_query, context)
    else:
        await mostrar_menu_principal_admin(update_or_query, context)

async def mostrar_menu_principal_admin(update_or_query, context):
    keyboard = [
        [InlineKeyboardButton("🔍 Buscar por Vendedor", callback_data="menu_buscar_vendedor")],
        [InlineKeyboardButton("🏢 Buscar por Razón Social (Cliente)", callback_data="menu_buscar_cliente")],
        [InlineKeyboardButton("🔢 Buscar por Nro de Pedido / Factura", callback_data="menu_buscar_pedido")],
        [InlineKeyboardButton("🌐 Búsqueda Libre General", callback_data="menu_buscar_libre")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    texto = "🎛️ **Menú Principal**\n\n¿Qué te gustaría consultar hoy?"
    
    await enviar_o_editar_mensaje(update_or_query, context, texto, reply_markup)

async def mostrar_menu_principal_vendedor(update_or_query, context):
    keyboard = [
        [InlineKeyboardButton("📦 Ver mis Pedidos Asignados", callback_data="vendedor_ver_mis_pedidos")],
        [InlineKeyboardButton("🔍 Buscar por Razón Social", callback_data="menu_buscar_cliente")],
        [InlineKeyboardButton("🔢 Buscar por Nro de Pedido", callback_data="menu_buscar_pedido")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    texto = "👤 **Menú de Vendedor**\n\nSeleccioná una opción para gestionar tus cuentas:"
    
    await enviar_o_editar_mensaje(update_or_query, context, texto, reply_markup)

async def enviar_o_editar_mensaje(update_or_query, context, texto, reply_markup):
    if hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(texto, reply_markup=reply_markup, parse_mode="Markdown")
    elif hasattr(update_or_query, 'edit_message_text'):
        try:
            await update_or_query.edit_message_text(texto, reply_markup=reply_markup, parse_mode="Markdown")
        except Exception:
            chat_id = update_or_query.message.chat_id if update_or_query.message else update_or_query.effective_chat.id
            await context.bot.send_message(chat_id=chat_id, text=texto, reply_markup=reply_markup, parse_mode="Markdown")

async def manejar_interaccion_chat(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    texto_crudo = update.message.text.strip()
    texto_lower = texto_crudo.lower()

    if user_id not in telefonos_verificados:
        await start(update, context)
        return

    if texto_lower in ["cancelar", "salir", "reiniciar"]:
        if user_id in user_sessions:
            user_sessions[user_id].clear()
        await update.message.reply_text("🚫 Operación cancelada. ¿En qué te puedo ayudar?")
        return

    saludos = ["hola", "buen dia", "buenas", "buenas tardes", "buenos dias", "hey", "hi", "saludos", "que tal", "menu"]
    if texto_lower in saludos:
        await mostrar_menu_segun_rol(update, context, user_id)
        return

    session = user_sessions.get(user_id, {})
    modo = session.get('modo_busqueda')

    # Si está esperando texto libre porque escribió directamente o eligió filtrar por texto
    if modo:
        await realizar_busqueda_optimizada(update, context, modo, texto_crudo)
    else:
        # Si mandó texto suelto sin elegir menú, interpretamos saludo o búsqueda libre amigable
        await realizar_busqueda_optimizada(update, context, 'libre', texto_crudo)

async def manejar_botones(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if user_id not in telefonos_verificados:
        await query.message.reply_text("⚠️ Por favor iniciá escribiendo /start o compartiendo tu contacto.")
        return

    if data == "cancelar_proceso":
        if user_id in user_sessions:
            user_sessions[user_id].clear()
        await query.edit_message_text("🚫 Operación cancelada. Escribí 'hola' cuando gustes para volver al menú.")
        return

    if data.startswith("menu_buscar_"):
        tipo = data.replace("menu_buscar_", "")
        if user_id not in user_sessions:
            user_sessions[user_id] = {}
        user_sessions[user_id]['modo_busqueda'] = tipo

        df = cargar_datos()
        if df.empty:
            await query.edit_message_text("⚠️ No hay datos cargados en el Excel actualmente.")
            return

        # Generar lista rápida de opciones según el tipo de búsqueda elegida
        keyboard = []
        columna_objetivo = None
        if tipo == 'vendedor':
            columna_objetivo = 'Vendedor'
        elif tipo == 'cliente':
            columna_objetivo = 'RazonSocial'
        elif tipo == 'pedido':
            columna_objetivo = 'Pedido'

        if columna_objetivo and columna_objetivo in df.columns:
            # Obtener valores únicos ordenados (máximo los primeros 30 para no saturar el bot)
            unicos = df[columna_objetivo].dropna().astype(str).unique()
            unicos = [u for u in unicos if u.strip() and u.lower() != 'nan'][:30]
            
            for item in unicos:
                # Callback corto para filtrar directamente por ese valor exacto
                keyboard.append([InlineKeyboardButton(f"📌 {item}", callback_data=f"filtrar_val_{tipo}_{item}")])

        keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
        
        texto_opciones = (
            f"📋 **Seleccioná un {tipo} de la lista** o simplemente **escribí en el chat** el nombre/número que buscás para agilizar:"
        )
        await query.edit_message_text(texto_opciones, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
        return

    if data.startswith("filtrar_val_"):
        _, tipo, valor_seleccionado = data.split("_", 2)
        user_sessions[user_id]['modo_busqueda'] = tipo
        await realizar_busqueda_optimizada(query, context, tipo, valor_seleccionado)
        return

    if data == "vendedor_ver_mis_pedidos":
        info_usuario = telefonos_verificados.get(user_id, {})
        vendedor_nombre = info_usuario.get('nombre_vendedor', '')
        await realizar_busqueda_optimizada(query, context, 'vendedor', vendedor_nombre)
        return

    if data == "post_si":
        if user_id in user_sessions:
            user_sessions[user_id].clear()
        await mostrar_menu_segun_rol(query, context, user_id)
        return

    if data == "post_no":
        if user_id in user_sessions:
            user_sessions[user_id].clear()
        await query.edit_message_text("¡Genial! Escribime un 'hola' cuando necesites consultar algo más. 👋")
        return

    session = user_sessions.get(user_id, {})
    df_encontrado = session.get('df_encontrado')

    if data.startswith("sel_row_"):
        row_idx = int(data.replace("sel_row_", ""))
        query_texto = session.get('query_texto', '')
        if row_idx in session['filas_seleccionadas']:
            session['filas_seleccionadas'].remove(row_idx)
        else:
            session['filas_seleccionadas'].append(row_idx)
            
        keyboard = []
        for index, row in df_encontrado.iterrows():
            icon = "✅" if index in session['filas_seleccionadas'] else "🔲"
            texto_btn = construir_texto_boton(icon, row, query_texto)
            keyboard.append([InlineKeyboardButton(texto_btn, callback_data=f"sel_row_{index}")])
            
        keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
        keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
        await query.edit_message_reply_markup(reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "conf_filas":
        if not session.get('filas_seleccionadas'):
            await query.answer("⚠️ Tildá al menos un elemento antes de continuar.", show_alert=True)
            return
        
        columnas_excluidas = {
            'Ítem', 'Artículo', 'CodColor', 'Vtas_Fact_PDF', 'Vtas_Remito_PDF',
            'Rom_Mov_Tipo', 'Rom_Mov_Numero', 'Vtas_Fact_fecha', 'PD_Id',
            'Alta', 'Asignado', 'EnPr', 'Prep', 'ARom', 'Roma', 'Fac',
            'Lib', 'Ent', 'desh', 'Con', 'Baja', 'Anul'
        }

        campos_brutos = list(df_encontrado.columns)
        if 'Estado' not in campos_brutos:
            campos_brutos.append('Estado')
            
        campos_disponibles = [col for col in campos_brutos if col not in columnas_excluidas]
        session['campos_disponibles'] = campos_disponibles
        session['paso'] = 'columnas'
        session['columnas_seleccionadas'] = campos_disponibles.copy() # Seleccionadas por defecto para mayor agilidad
        
        keyboard = []
        for campo in campos_disponibles:
            icon = "✅" # Por defecto marcadas
            keyboard.append([InlineKeyboardButton(f"{icon} {campo}", callback_data=f"sel_col_{campo}")])
        keyboard.append([InlineKeyboardButton("🚀 Generar Reporte", callback_data="conf_columnas")])
        keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
        
        try:
            await query.message.delete()
        except Exception:
            pass

        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="✅ Columnas cargadas. Podés tocar alguna para quitarla o darle directamente a generar:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif data.startswith("sel_col_"):
        campo = data.replace("sel_col_", "")
        if campo in session['columnas_seleccionadas']:
            session['columnas_seleccionadas'].remove(campo)
        else:
            session['columnas_seleccionadas'].append(campo)
            
        keyboard = []
        for campo_op in session['campos_disponibles']:
            icon = "✅" if campo_op in session['columnas_seleccionadas'] else "🔲"
            keyboard.append([InlineKeyboardButton(f"{icon} {campo_op}", callback_data=f"sel_col_{campo_op}")])
        keyboard.append([InlineKeyboardButton("🚀 Generar Reporte", callback_data="conf_columnas")])
        keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
        
        await query.edit_message_reply_markup(reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "conf_columnas":
        if not session.get('columnas_seleccionadas'):
            await query.answer("⚠️ Seleccioná al menos una columna.", show_alert=True)
            return

        filas_sel = session['filas_seleccionadas']
        cols_sel = session['columnas_seleccionadas']
        df_filtrado = df_encontrado.loc[filas_sel].copy()

        columnas_estado_posibles = ['Anul', 'Baja', 'Ent', 'Lib', 'Roma', 'ARom', 'Prep', 'EnPr', 'Asignado', 'Alta']
        
        def calcular_estado_actual(row):
            for col in columnas_estado_posibles:
                if col in row and pd.notna(row[col]) and str(row[col]).strip() != "" and str(row[col]).lower() != "nan" and str(row[col]) != "0":
                    val_col = str(row[col]).strip()
                    if val_col not in ["1", "True", "x", "X"]:
                        return f"{col} ({val_col})"
                    else:
                        return col
            return "Pendiente"

        df_filtrado['EstadoCalculado'] = df_filtrado.apply(calcular_estado_actual, axis=1)
        await query.delete_message()

        for index, row in df_filtrado.iterrows():
            pedido_val = row.get('Pedido', '-')
            cliente_val = row.get('RazonSocial', '-')

            msg_ind = f"📦 **Detalle de Pedido**\n🔢 Pedido: {pedido_val}\n👤 Cliente: {cliente_val}\n\n"
            for col in cols_sel:
                if col not in ['Pedido', 'RazonSocial']:
                    val = row['EstadoCalculado'] if col == 'Estado' else row.get(col, '-')
                    if pd.isna(val) or str(val).lower() == 'nan' or str(val) == '':
                        val = '-'
                    msg_ind += f"• {col}: {val}\n"

            await context.bot.send_message(chat_id=query.message.chat_id, text=msg_ind, parse_mode="Markdown")

        cols_a_exportar = [c if c != 'Estado' else 'EstadoCalculado' for c in cols_sel]
        buffer_excel = io.BytesIO()
        df_filtrado[cols_a_exportar].to_excel(buffer_excel, index=False)
        buffer_excel.seek(0)

        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=buffer_excel,
            filename="Reporte_Consolidado.xlsx",
            caption="📥 Acá tenés el archivo Excel con tu selección."
        )
        
        session.clear()
        teclado_continuar = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Sí, consultar otro", callback_data="post_si"), InlineKeyboardButton("❌ No, gracias", callback_data="post_no")]
        ])
        await context.bot.send_message(chat_id=query.message.chat_id, text="¿Te puedo ayudar con otra consulta?", reply_markup=teclado_continuar)

async def realizar_busqueda_optimizada(update_or_query, context, tipo_filtro, query_texto):
    chat_id = update_or_query.effective_chat.id if hasattr(update_or_query, 'effective_chat') else update_or_query.message.chat_id
    user_id = update_or_query.effective_user.id if hasattr(update_or_query, 'effective_user') else update_or_query.from_user.id
    
    query_lower = str(query_texto).lower().strip()
    df = cargar_datos()
    if df.empty:
        msg = "⚠️ No se pudo acceder a los datos de los pedidos."
        if hasattr(update_or_query, 'message') and update_or_query.message:
            await update_or_query.message.reply_text(msg)
        else:
            await context.bot.send_message(chat_id=chat_id, text=msg)
        return

    info_usuario = telefonos_verificados.get(user_id, {})
    if info_usuario.get('rol') == 'vendedor':
        vendedor_asignado = info_usuario.get('nombre_vendedor', '').lower()
        if vendedor_asignado and 'Vendedor' in df.columns:
            df = df[df['Vendedor'].str.lower().str.contains(vendedor_asignado, na=False)]

    if tipo_filtro == 'vendedor' and 'Vendedor' in df.columns:
        resultados = df[df['Vendedor'].str.lower().str.contains(query_lower, na=False)].copy()
    elif tipo_filtro == 'cliente' and 'RazonSocial' in df.columns:
        resultados = df[df['RazonSocial'].str.lower().str.contains(query_lower, na=False)].copy()
    elif tipo_filtro == 'pedido' and 'Pedido' in df.columns:
        resultados = df[df['Pedido'].str.lower().str.contains(query_lower, na=False)].copy()
    else:
        mask = pd.Series(False, index=df.index)
        for col in df.columns:
            mask = mask | df[col].str.lower().str.contains(query_lower, na=False)
        resultados = df[mask].copy()

    if resultados.empty:
        msg = f"❌ No encontré coincidencias para '{query_texto}'."
        if hasattr(update_or_query, 'message') and update_or_query.message:
            await update_or_query.message.reply_text(msg)
        else:
            await context.bot.send_message(chat_id=chat_id, text=msg)
        return

    try:
        resultados['Pedido_num'] = pd.to_numeric(resultados['Pedido'], errors='coerce')
        resultados = resultados.sort_values(by='Pedido_num', ascending=True)
        resultados = resultados.drop(columns=['Pedido_num'])
    except Exception:
        pass

    if user_id not in user_sessions:
        user_sessions[user_id] = {}
        
    user_sessions[user_id].update({
        'df_encontrado': resultados,
        'query_texto': query_texto,
        'filas_seleccionadas': list(resultados.index)[:15], # Seleccionados por defecto los primeros para agilizar
        'columnas_seleccionadas': [],
        'paso': 'filas'
    })

    keyboard = []
    for index, row in resultados.iterrows():
        icon = "✅" if index in user_sessions[user_id]['filas_seleccionadas'] else "🔲"
        texto_btn = construir_texto_boton(icon, row, query_texto)
        keyboard.append([InlineKeyboardButton(texto_btn, callback_data=f"sel_row_{index}")])
    
    keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
    keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
    
    texto_resp = f"🔍 Encontré {len(resultados)} resultados. Seleccioná los que necesites:"
    if hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(texto_resp, reply_markup=InlineKeyboardMarkup(keyboard))
    else:
        await context.bot.send_message(chat_id=chat_id, text=texto_resp, reply_markup=InlineKeyboardMarkup(keyboard))

def main():
    if not TOKEN:
        print("❌ Error: Falta configurar la variable TELEGRAM_TOKEN.")
        return
        
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()

    request_config = HTTPXRequest(read_timeout=30.0, connect_timeout=30.0)
    app = Application.builder().token(TOKEN).request(request_config).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.CONTACT, recibir_contacto))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, manejar_interaccion_chat))
    app.add_handler(CallbackQueryHandler(manejar_botones))
    
    print("🤖 Bot optimizado y activo. Presioná Ctrl+C para salir.")
    app.run_polling()

if __name__ == '__main__':
    main()