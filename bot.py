import os
import io
import threading
import pandas as pd
from flask import Flask
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
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
    """Genera un texto adaptativo y limpio para el botón según lo que se esté buscando"""
    pedido = row.get('Pedido', 'S/N')
    cliente = row.get('RazonSocial', 'Desconocido')
    vendedor = row.get('Vendedor', 'S/V')
    fecha = row.get('Vtas_Fact_fecha', 'S/F')
    
    # Limpiamos la fecha para que quede cortita (ej: '15/05/2026')
    if len(str(fecha)) >= 10 and '-' in str(fecha):
        partes_fecha = str(fecha).split()[0].split('-')
        if len(partes_fecha) == 3:
            fecha = f"{partes_fecha[2]}/{partes_fecha[1]}/{partes_fecha[0]}"

    # Armado dinámico inteligente con fecha, pedido, cliente y vendedor
    if any(char.isdigit() for char in query_texto) and len(query_texto) >= 3 and query_texto in str(pedido).lower():
        texto = f"{icon} 📅 {fecha} | Ped: {pedido} | {cliente} ({vendedor})"
    elif query_texto in str(vendedor).lower():
        texto = f"{icon} 📅 {fecha} | Vendedor: {vendedor} | Ped: {pedido} - {cliente}"
    else:
        texto = f"{icon} 📅 {fecha} | Ped: {pedido} | {cliente} ({vendedor})"

    # Control de límite estricto de Telegram para botones
    if len(texto) > 60:
        texto = texto[:57] + "..."
        
    return texto

async def enviar_bienvenida(update_or_query, context):
    texto_bienvenida = (
        "👋 ¡Hola! Soy tu bot de trazabilidad.\n\n"
        "🔎 Escribí parte del nombre de un vendedor, un número de pedido, factura o cualquier dato para buscar:"
    )
    if hasattr(update_or_query, 'message') and update_or_query.message:
        await update_or_query.message.reply_text(texto_bienvenida)
    elif hasattr(update_or_query, 'edit_message_text'):
        await context.bot.send_message(chat_id=update_or_query.message.chat_id, text=texto_bienvenida)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await enviar_bienvenida(update, context)

async def manejar_saludos_o_busqueda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto_crudo = update.message.text.strip()
    texto = texto_crudo.lower()
    user_id = update.effective_user.id
    
    if texto in ["cancelar", "salir", "reiniciar"]:
        if user_id in user_sessions:
            del user_sessions[user_id]
        await update.message.reply_text("🚫 Operación cancelada. Escribime cuando quieras para hacer otra búsqueda.")
        return

    saludos = ["hola", "buen dia", "buenas", "buenas tardes", "buenos dias", "hey", "hi", "saludos", "que tal"]

    if user_id in user_sessions:
        try:
            await update.message.delete()
        except Exception:
            pass

        session = user_sessions[user_id]
        await update.message.reply_text(
            "⚠️ Che, tenés una sesión pendiente. Te vuelvo a mostrar las opciones para que puedas continuar (o mandá 'cancelar' para cortar)."
        )
        if session.get('columnas_disponibles') or len(session.get('columnas_seleccionadas', [])) > 0 or session.get('paso') == 'columnas':
            session['paso'] = 'columnas'
            await mostrar_menu_columnas(update.effective_chat.id, context, session)
        else:
            session['paso'] = 'filas'
            await mostrar_menu_filas(update.effective_chat.id, context, session)
        return

    if texto in saludos or texto in ["?", "ayuda", "help", "estoy perdido", "como es esto"]:
        if texto in saludos:
            await enviar_bienvenida(update, context)
        else:
            await update.message.reply_text(
                "¿Te perdiste? No pasa nada. Mandame el número de pedido, el nombre/razón social del cliente o el nombre del vendedor que querés consultar."
            )
        return

    await recibir_texto_core(update, context)

async def mostrar_menu_filas(chat_id, context, session):
    df_encontrado = session['df_encontrado']
    query_texto = session.get('query_texto', '')
    keyboard = []
    for index, row in df_encontrado.iterrows():
        icon = "✅" if index in session['filas_seleccionadas'] else "🔲"
        texto_boton = construir_texto_boton(icon, row, query_texto)
        keyboard.append([InlineKeyboardButton(texto_boton, callback_data=f"sel_row_{index}")])
    
    keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
    keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
    
    await context.bot.send_message(
        chat_id=chat_id,
        text="📋 Acá tenés nuevamente el menú de selección de pedidos:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def mostrar_menu_columnas(chat_id, context, session):
    keyboard = []
    for campo_op in session['campos_disponibles']:
        icon = "✅" if campo_op in session['columnas_seleccionadas'] else "🔲"
        keyboard.append([InlineKeyboardButton(f"{icon} {campo_op}", callback_data=f"sel_col_{campo_op}")])
    keyboard.append([InlineKeyboardButton("🚀 Generar Reporte", callback_data="conf_columnas")])
    keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
    
    await context.bot.send_message(
        chat_id=chat_id,
        text="📋 Acá tenés nuevamente el menú de selección de columnas:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def recibir_texto_core(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query_texto = ' '.join(update.message.text.strip().lower().split())
    
    if len(query_texto) < 2:
        await update.message.reply_text("Mmm, pusiste muy pocos caracteres para buscar. Tirame un dato un poco más completo (mínimo 2 letras o números).")
        return

    df = cargar_datos()
    if df.empty:
        await update.message.reply_text("⚠️ Che, el archivo Excel no se encuentra disponible o está vacío en este momento.")
        return

    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        serie_texto = df[col].str.lower()
        mask = mask | serie_texto.str.contains(query_texto, na=False)

    resultados = df[mask].copy()

    if resultados.empty:
        await update.message.reply_text(
            f"❌ No encontré ningún registro con '{query_texto}'.\n"
            "Fijate si está bien escrito o probá buscando por otra cosa (vendedor, cliente, factura)."
        )
        return

    try:
        resultados['Pedido_num'] = pd.to_numeric(resultados['Pedido'], errors='coerce')
        resultados = resultados.sort_values(by='Pedido_num', ascending=True)
        resultados = resultados.drop(columns=['Pedido_num'])
    except Exception:
        pass

    user_id = update.effective_user.id
    user_sessions[user_id] = {
        'df_encontrado': resultados,
        'query_texto': query_texto,
        'filas_seleccionadas': [],
        'columnas_seleccionadas': [],
        'paso': 'filas'
    }

    keyboard = []
    for index, row in resultados.iterrows():
        texto_boton = construir_texto_boton("🔲", row, query_texto)
        keyboard.append([InlineKeyboardButton(texto_boton, callback_data=f"sel_row_{index}")])
    
    keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
    keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
    
    await update.message.reply_text(
        f"🔍 Encontré {len(resultados)} coincidencias (ordenadas de más vieja a más nueva). Tildá las que necesites y dale a confirmar:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def manejar_botones(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    if data == "cancelar_proceso":
        if user_id in user_sessions:
            del user_sessions[user_id]
        await query.edit_message_text("🚫 Operación cancelada. Escribime cuando quieras para hacer otra búsqueda.")
        return

    if data == "post_si":
        await query.edit_message_text("Entendido. Ingresá el número de pedido, vendedor o nombre de cliente para la nueva búsqueda:")
        return

    if data == "post_no":
        await query.edit_message_text("¡Genial! Que tengas un excelente día. Escribime un 'Hola' cuando necesites consultar algo más. 👋")
        return

    if user_id not in user_sessions:
        await query.edit_message_text("⚠️ Ups, esta sesión expiró. Escribime tu búsqueda de nuevo por favor.")
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
            texto_boton = construir_texto_boton(icon, row, query_texto)
            keyboard.append([InlineKeyboardButton(texto_boton, callback_data=f"sel_row_{index}")])
            
        keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
        keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])
        await query.edit_message_reply_markup(reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "conf_filas":
        if not session['filas_seleccionadas']:
            keyboard = []
            for index, row in df_encontrado.iterrows():
                icon = "✅" if index in session['filas_seleccionadas'] else "🔲"
                texto_boton = construir_texto_boton(icon, row, query_texto)
                keyboard.append([InlineKeyboardButton(texto_boton, callback_data=f"sel_row_{index}")])
            
            keyboard.append([InlineKeyboardButton("✅ Confirmar Selección", callback_data="conf_filas")])
            keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="⚠️ **¡Atención!** No tildaste ningún pedido.\nPor favor, elegí al menos uno de la lista acá abajo:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
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
            keyboard = []
            for campo_op in session['campos_disponibles']:
                icon = "✅" if campo_op in session['columnas_seleccionadas'] else "🔲"
                keyboard.append([InlineKeyboardButton(f"{icon} {campo_op}", callback_data=f"sel_col_{campo_op}")])
            keyboard.append([InlineKeyboardButton("🚀 Generar Reporte", callback_data="conf_columnas")])
            keyboard.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_proceso")])

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="⚠️ **¡Atención!** No elegiste ninguna columna.\nPor favor, tildá al menos una columna acá abajo:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
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
            del user_sessions[user_id]
        
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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, manejar_saludos_o_busqueda))
    app.add_handler(CallbackQueryHandler(manejar_botones))
    
    print("🤖 Bot listo y servidor web falso corriendo... Presioná Ctrl+C para detenerlo.")
    app.run_polling()

if __name__ == '__main__':
    main()