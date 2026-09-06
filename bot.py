import os
import io
import threading
import pandas as pd
from flask import Flask
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from telegram.request import HTTPXRequest

# Configuración de Flask para que Render detecte un puerto abierto y mantenga el servicio gratis
app_flask = Flask(__name__)

@app_flask.route('/')
def home():
    return "¡El bot de trazabilidad está activo y funcionando!"

def run_flask():
    port = int(os.environ.get("PORT", 10000))
    app_flask.run(host="0.0.0.0", port=port)

# Leemos el token de forma segura desde las variables de entorno de la nube o del sistema
TOKEN = os.environ.get("TELEGRAM_TOKEN")
EXCEL_PATH = "datos_pedidos.xlsx"
NOMBRE_HOJA = "Pedidos_Pistachos (2)"

user_sessions = {}

# Definición de roles por número de teléfono
ADMIN_PHONES = {
    "541153248379",
    "541156680181",
    "541133635660"
}

# Mapeo opcional de vendedores autorizados (si se requiere restringir a vendedores específicos)
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
    """Genera un texto adaptativo y limpio para el botón ordenado exactamente según el tipo de búsqueda"""
    pedido = row.get('Pedido', 'S/N')
    cliente = row.get('RazonSocial', 'Desconocido')
    vendedor = row.get('Vendedor', 'S/V')
    fecha = row.get('Vtas_Fact_fecha', 'S/F')
    
    if len(str(fecha)) >= 10 and '-' in str(fecha):
        partes_fecha = str(fecha).split()[0].split('-')
        if len(partes_fecha) == 3:
            fecha = f"{partes_fecha[2]}/{partes_fecha[1]}/{partes_fecha[0]}"

    es_pedido = any(char.isdigit() for char in query_texto) and len(query_texto) >= 3 and query_texto in str(pedido).lower()
    es_vendedor = query_texto in str(vendedor).lower()

    if es_pedido:
        texto = f"{icon} 📅 {fecha} | {cliente} | {vendedor} | Ped: {pedido}"
    elif es_vendedor:
        texto = f"{icon} 📅 {fecha} | {cliente} | Ped: {pedido} | {vendedor}"
    else:
        texto = f"{icon} Ped: {pedido} | 📅 {fecha} | {vendedor} | {cliente}"

    if len(texto) > 60:
        texto = texto[:57] + "..."
        
    return texto

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id in user_sessions and 'telefono' in user_sessions[user_id]:
        tel = user_sessions[user_id]['telefono']
        if tel in ADMIN_PHONES:
            await mostrar_menu_principal_admin(update, context)
            return

    contacto_btn = KeyboardButton("📱 Compartir mi contacto para ingresar", request_contact=True)
    reply_markup = ReplyKeyboardMarkup([[contacto_btn]], resize_keyboard=True, one_time_keyboard=True)
    
    await update.message.reply_text(
        "👋 ¡Hola! Bienvenido al bot de trazabilidad de Pistachos.\n\n"
        "Para garantizar la seguridad y asignarte los permisos correspondientes, por favor compartí tu número de teléfono tocando el botón de abajo:",
        reply_markup=reply_markup
    )

async def recibir_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    if not message.contact:
        return
        
    telefono = limpiar_telefono(message.contact.phone_number)
    user_id = update.effective_user.id
    
    if user_id not in user_sessions:
        user_sessions[user_id] = {}
    user_sessions[user_id]['telefono'] = telefono
    
    if telefono in ADMIN_PHONES:
        user_sessions[user_id]['rol'] = 'admin'
        await message.reply_text(
            "✅ ¡Verificado como Administrador!",
            reply_markup=ReplyKeyboardRemove()
        )
        await mostrar_menu_principal_admin(update, context)
    elif telefono in VENDEDORES_MAP:
        vendedor_nombre = VENDEDORES_MAP[telefono]
        user_sessions[user_id]['rol'] = 'vendedor'
        user_sessions[user_id]['nombre_vendedor'] = vendedor_nombre
        await message.reply_text(
            f"✅ ¡Verificado como Vendedor ({vendedor_nombre})!",
            reply_markup=ReplyKeyboardRemove()
        )
        await mostrar_menu_principal_vendedor(update, context)
    else:
        # Permitir acceso general o restringir según prefieras. Aquí permitimos rol estándar con menús guiados:
        user_sessions[user_id]['rol'] = 'estandar'
        await message.reply_text(
            "✅ ¡Verificado correctamente!",
            reply_markup=ReplyKeyboardRemove()
        )
        await mostrar_menu_principal_admin(update, context)

async def mostrar_menu_principal_admin(update_or_query, context):
    keyboard = [
        [InlineKeyboardButton("🔍 Buscar por Vendedor", callback_data="menu_buscar_vendedor")],
        [InlineKeyboardButton("🏢 Buscar por Razón Social (Cliente)", callback_data="menu_buscar_cliente")],
        [InlineKeyboardButton("🔢 Buscar por Nro de Pedido / Factura", callback_data="menu_buscar_pedido")],
        [InlineKeyboardButton("🌐 Búsqueda Libre General", callback_data="menu_buscar_libre")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    texto = (
        "🎛️ **Menú Principal de Búsqueda**\n\n"
        "Seleccioná cómo querés realizar tu consulta para agilizar el proceso:"
    )
    
    if hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(texto, reply_markup=reply_markup, parse_mode="Markdown")
    elif hasattr(update_or_query, 'edit_message_text'):
        try:
            await update_or_query.edit_message_text(texto, reply_markup=reply_markup, parse_mode="Markdown")
        except Exception:
            await context.bot.send_message(chat_id=update_or_query.message.chat_id, text=texto, reply_markup=reply_markup, parse_mode="Markdown")

async def mostrar_menu_principal_vendedor(update_or_query, context):
    keyboard = [
        [InlineKeyboardButton("📦 Ver mis Pedidos Asignados", callback_data="vendedor_ver_mis_pedidos")],
        [InlineKeyboardButton("🔍 Buscar por Razón Social", callback_data="menu_buscar_cliente")],
        [InlineKeyboardButton("🔢 Buscar por Nro de Pedido", callback_data="menu_buscar_pedido")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    texto = "👤 **Menú de Vendedor**\n\nSeleccioná una opción:"
    
    if hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(texto, reply_markup=reply_markup, parse_mode="Markdown")
    elif hasattr(update_or_query, 'edit_message_text'):
        try:
            await update_or_query.edit_message_text(texto, reply_markup=reply_markup, parse_mode="Markdown")
        except Exception:
            await context.bot.send_message(chat_id=update_or_query.message.chat_id, text=texto, reply_markup=reply_markup, parse_mode="Markdown")

async def manejar_saludos_o_busqueda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto_crudo = update.message.text.strip()
    texto = texto_crudo.lower()
    user_id = update.effective_user.id
    
    if texto in ["cancelar", "salir", "reiniciar"]:
        if user_id in user_sessions:
            user_sessions[user_id].pop('modo_busqueda', None)
            user_sessions[user_id].pop('paso', None)
        await update.message.reply_text("🚫 Operación cancelada. Escribí /start para volver al menú principal.")
        return

    if user_id not in user_sessions or 'telefono' not in user_sessions[user_id]:
        await update.message.reply_text("⚠️ Por favor, envianos tu contacto primero usando el comando /start.")
        return

    session = user_sessions[user_id]
    
    if session.get('paso') == 'columnas' or session.get('paso') == 'filas':
        try:
            await update.message.delete()
        except Exception:
            pass
        await update.message.reply_text("⚠️ Tenés una sesión pendiente. Usá los botones de arriba o escribí 'cancelar'.")
        return

    saludos = ["hola", "buen dia", "buenas", "buenas tardes", "buenos dias", "hey", "hi", "saludos", "que tal"]
    if texto in saludos:
        await mostrar_menu_principal_admin(update, context)
        return

    modo = session.get('modo_busqueda', 'libre')
    await realizar_busqueda_optimizada(update, context, modo, texto_crudo)

async def realizar_busqueda_optimizada(update_or_query, context, tipo_filtro, query_texto):
    chat_id = update_or_query.effective_chat.id if hasattr(update_or_query, 'effective_chat') else update_or_query.message.chat_id
    user_id = update_or_query.effective_user.id if hasattr(update_or_query, 'effective_user') else update_or_query.from_user.id
    
    query_lower = query_texto.lower().strip()
    if len(query_lower) < 2:
        msg = "Mmm, pusiste muy pocos caracteres. Tirame un dato más completo (mínimo 2 letras o números)."
        if hasattr(update_or_query, 'message') and update_or_query.message:
            await update_or_query.message.reply_text(msg)
        else:
            await context.bot.send_message(chat_id=chat_id, text=msg)
        return

    df = cargar_datos()
    if df.empty:
        msg = "⚠️ El archivo Excel no se encuentra disponible o está vacío en este momento."
        if hasattr(update_or_query, 'message') and update_or_query.message:
            await update_or_query.message.reply_text(msg)
        else:
            await context.bot.send_message(chat_id=chat_id, text=msg)
        return

    session = user_sessions.get(user_id, {})
    rol = session.get('rol', 'estandar')
    
    if rol == 'vendedor':
        vendedor_asignado = session.get('nombre_vendedor', '').lower()
        if vendedor_asignado and 'Vendedor' in df.columns:
            df = df[df['Vendedor'].str.lower().str.contains(vendedor_asignado, na=False)]

    if tipo_filtro == 'vendedor' and 'Vendedor' in df.columns:
        resultados = df[df['Vendedor'].str.lower().str.contains(query_lower, na=False)].copy()
    elif tipo_filtro == 'cliente' and 'RazonSocial' in df.columns:
        resultados = df[df['RazonSocial'].str.lower().str.contains(query_lower, na=False)].copy()
    elif tipo_filtro == 'pedido' and 'Pedido' in df.columns:
        resultados = df[df['Pedido'].str.lower().str.contains(query_lower, na=False)].copy()
    elif tipo_filtro == 'mis_pedidos':
        resultados = df.copy()
    else:
        mask = pd.Series(False, index=df.index)
        for col in df.columns:
            mask = mask | df[col].str.lower().str.contains(query_lower, na=False)
        resultados = df[mask].copy()

    if resultados.empty:
        msg = f"❌ No encontré ningún registro con '{query_texto}'."
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

    user_sessions[user_id] = {
        **session,
        'df_encontrado': resultados,
        'query_texto': query_texto,
        'filas_seleccionadas': [],
        'columnas_seleccionadas': [],
        'paso': 'filas'
    }

    keyboard = []
    for index, row in resultados.iterrows():
        texto_btn = construir_texto_boton("🔲", row, query_texto)
        keyboard.append([InlineKeyboardButton(texto_btn, callback_data=f"sel_row_{index}")])
    
    keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
    keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
    
    texto_resp = f"🔍 Encontré {len(resultados)} coincidencias. Tildá las que necesites y dale a confirmar:"
    if hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(texto_resp, reply_markup=InlineKeyboardMarkup(keyboard))
    else:
        await context.bot.send_message(chat_id=chat_id, text=texto_resp, reply_markup=InlineKeyboardMarkup(keyboard))

async def manejar_botones(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if data == "cancelar_proceso":
        if user_id in user_sessions:
            user_sessions[user_id].pop('modo_busqueda', None)
            user_sessions[user_id].pop('paso', None)
        await query.edit_message_text("🚫 Operación cancelada. Escribí /start para volver al menú principal.")
        return

    if data.startswith("menu_buscar_"):
        tipo = data.replace("menu_buscar_", "")
        if user_id not in user_sessions:
            user_sessions[user_id] = {}
        user_sessions[user_id]['modo_busqueda'] = tipo
        
        nombres_campos = {
            "vendedor": "el nombre del vendedor",
            "cliente": "la razón social o nombre del cliente",
            "pedido": "el número de pedido o factura",
            "libre": "cualquier palabra clave para búsqueda general"
        }
        campo_str = nombres_campos.get(tipo, "el dato")
        await query.edit_message_text(f"✍️ Perfecto. Escribí {campo_str} que querés buscar:")
        return

    if data == "vendedor_ver_mis_pedidos":
        session = user_sessions.get(user_id, {})
        vendedor_nombre = session.get('nombre_vendedor', '')
        await realizar_busqueda_optimizada(query, context, 'vendedor', vendedor_nombre)
        return

    if data == "post_si":
        await mostrar_menu_principal_admin(query, context)
        return

    if data == "post_no":
        await query.edit_message_text("¡Genial! Que tengas un excelente día. Escribí /start cuando necesites consultar algo más. 👋")
        return

    if user_id not in user_sessions or 'df_encontrado' not in user_sessions[user_id]:
        await query.edit_message_text("⚠️ Ups, esta sesión expiró o no es válida. Escribí /start de nuevo.")
        return

    session = user_sessions[user_id]
    df_encontrado = session['df_encontrado']
    query_texto = session.get('query_texto', '')

    if data.startswith("sel_row_"):
        row_idx = int(data.replace("sel_row_", ""))
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
        if not session['filas_seleccionadas']:
            await query.answer("⚠️ Tenés que tildar al menos un pedido antes de confirmar.", show_alert=True)
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
        
        keyboard = []
        for campo in campos_disponibles:
            keyboard.append([InlineKeyboardButton(f"🔲 {campo}", callback_data=f"sel_col_{campo}")])
        keyboard.append([InlineKeyboardButton("🚀 Generar Reporte", callback_data="conf_columnas")])
        keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
        
        try:
            await query.message.delete()
        except Exception:
            pass

        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="Perfecto. Ahora elegí qué columnas querés que salgan en el informe:",
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
        if not session['columnas_seleccionadas']:
            await query.answer("⚠️ Tildá al menos una columna para generar el reporte.", show_alert=True)
            return

        filas_sel = session['filas_seleccionadas']
        cols_sel = session['columnas_seleccionadas']

        df_filtrado = df_encontrado.loc[filas_sel].copy()

        columnas_estado_posibles = ['Anul', 'Baja', 'Ent', 'Lib', 'Roma', 'ARom', 'Prep', 'EnPr', 'Asignado', 'Alta']
        
        def calcular_estado_actual(row):
            for col in columnas_estado_posibles:
                if col in row and pd.notna(row[col]) and str(row[col]).strip() != "" and str(row[col]).lower() != "nan" and str(row[col]) != "0" and str(row[col]).lower() != "nat":
                    val_col = str(row[col]).strip()
                    if val_col not in ["1", "True", "x", "X"]:
                        return f"{col} ({val_col})"
                    else:
                        return col
            return "Pendiente / Sin iniciar"

        df_filtrado['EstadoCalculado'] = df_filtrado.apply(calcular_estado_actual, axis=1)

        await query.delete_message()

        for index, row in df_filtrado.iterrows():
            pedido_val = row.get('Pedido', '-')
            cliente_val = row.get('RazonSocial', '-')

            mensaje_individual = f"📦 Reporte de Pedido\n"
            mensaje_individual += f"🔢 Pedido: {pedido_val}\n"
            mensaje_individual += f"👤 Cliente: {cliente_val}\n\n"
            
            for col in cols_sel:
                if col not in ['Pedido', 'RazonSocial']:
                    if col == 'Estado':
                        valor_col = row['EstadoCalculado']
                    else:
                        valor_col = row.get(col, '-')
                        if pd.isna(valor_col) or str(valor_col).lower() == 'nan' or str(valor_col) == '':
                            valor_col = '-'
                    mensaje_individual += f"• {col}: {valor_col}\n"

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text=mensaje_individual
            )

        cols_a_exportar = [c if c != 'Estado' else 'EstadoCalculado' for c in cols_sel]
        
        buffer_excel = io.BytesIO()
        df_filtrado[cols_a_exportar].to_excel(buffer_excel, index=False)
        buffer_excel.seek(0)

        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=buffer_excel,
            filename="Reporte_Filtrado.xlsx",
            caption="📥 Acá tenés el Excel listo con el consolidado."
        )
        
        if user_id in user_sessions:
            user_sessions[user_id].pop('modo_busqueda', None)
            user_sessions[user_id].pop('paso', None)
        
        teclado_continuar = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Sí", callback_data="post_si"), InlineKeyboardButton("❌ No", callback_data="post_no")]
        ])
        
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="¿Te puedo ayudar con algo más?",
            reply_markup=teclado_continuar
        )
        return

def main():
    if not TOKEN:
        print("❌ Error: No se encontró la variable de entorno TELEGRAM_TOKEN.")
        return
        
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()

    request_config = HTTPXRequest(read_timeout=30.0, connect_timeout=30.0)
    app = Application.builder().token(TOKEN).request(request_config).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.CONTACT, recibir_contacto))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, manejar_saludos_o_busqueda))
    app.add_handler(CallbackQueryHandler(manejar_botones))
    
    print("🤖 Bot listo y servidor web falso corriendo... Presioná Ctrl+C para detenerlo.")
    app.run_polling()

if __name__ == '__main__':
    main()