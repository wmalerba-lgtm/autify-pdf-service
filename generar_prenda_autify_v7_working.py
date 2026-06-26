#!/usr/bin/env python3
"""
AUTIFY — Generador de Contrato de Prenda v7.0
El script lee coordenadas y configuración desde parametros_contrato_prenda.xlsx.
El Excel es la única fuente de verdad — editá el Excel, no este script.

Uso:
  python3 generar_prenda_autify.py <solicitud.pdf> <template.pdf> <output.pdf> \
                                   [fecha dd/mm/yyyy] [tipo UVA_PI|UVA_PRE|FIJA_PI|FIJA_PRE] \
                                   [carta_aprobacion.pdf] [mutuo_prendario.pdf]

Tipos de operación:
  UVA_PI   → UVA + Prenda Inscripta   (default)
  UVA_PRE  → UVA + Pre-prenda
  FIJA_PI  → FIJA + Prenda Inscripta
  FIJA_PRE → FIJA + Pre-prenda

v6.0 — Cambios:
  - H1_BIEN: agrega Marca Motor / Marca Chasis con placeholder "COMPLETAR CORRECTAMENTE"
    (deben verificarse manualmente contra el Título del Automotor).
  - Hoja 3, fila ~85.5: la fecha (DD/AA) ahora se toma de la Carta de Aprobación,
    separada de la fecha de encabezado (que sigue siendo XX/XX a completar por el cliente).
  - H3_UVA_VALOR_LETRAS: si el texto excede el ancho disponible, continúa en la línea
    siguiente sin pisar el contenido de abajo.

v7.0 — Cambios:
  - Nuevo parser parsear_mutuo_prendario() para casos de Pre-prenda (UVA_PRE / FIJA_PRE),
    que trae su propio "corte" de UVAs más cercano al desembolso.
  - Para tipo_op == '*_PRE': H3_UVAS_LETRAS/NUM, H3_FECHA85_DD/MM/AA y
    H3_UVA_VALOR_LETRAS/NUM se toman del Mutuo Prendario (UVAS EQUIVALENTE,
    FECHA VALOR UVA, COTIZACION UVA) en lugar de la Solicitud/Carta de Aprobación.
  - H1_VENCIMIENTOS también puede tomar el primer vencimiento desde el Mutuo
    Prendario cuando tipo_op es *_PRE (mismo valor, distinta fuente).
  - Nueva variable H5_CONYUGE_BOX: si el deudor es casado, dibuja un recuadro
    semitransparente sobre la sección de "Régimen Patrimonial" (hoja 5,
    de columna 12 fila 40 a columna 50 fila 49) indicando que debe completarse
    a mano.
"""

import sys, re, io, os
from datetime import date
import pdfplumber
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.colors import red
from pypdf import PdfReader, PdfWriter

# ── CONSTANTES ──────────────────────────────────────────────
MM   = 2.8346
STEP = 2.5 * MM   # 1 celda grilla = 7.0866 pt
W, H = 595.0, 842.0

# Mapeo tipo_op → columna del Excel (D=4, E=5, F=6, G=7)
# Columnas del Excel (7 columnas, sin Variable_ID):
# A=Coordenadas B=Variable C=Hoja D=UVA_PI E=UVA_PRE F=FIJA_PI G=FIJA_PRE
TIPO_COL = {
    'UVA_PI':  4,   # D
    'UVA_PRE': 5,   # E
    'FIJA_PI': 6,   # F
    'FIJA_PRE':7,   # G
}

MESES = ['','enero','febrero','marzo','abril','mayo','junio',
         'julio','agosto','septiembre','octubre','noviembre','diciembre']


# ── HELPERS DE GRILLA ───────────────────────────────────────
def gx(col): return col * STEP
def gy(row): return H - row * STEP

def draw(cv, col, row, texto, size=6.5, bold=True):
    if not str(texto).strip(): return
    cv.setFont("Helvetica-Bold" if bold else "Helvetica", size)
    cv.setFillColor(red)
    cv.drawString(gx(col), gy(row), str(texto))

def draw_box(cv, col1, row1, col2, row2, fill_color=None, alpha=0.15,
              stroke_color=red, stroke_alpha=0.6, line_width=0.75):
    """
    Dibuja un rectángulo semitransparente desde (col1,row1) [esquina
    superior-izquierda en coordenadas de grilla] hasta (col2,row2)
    [esquina inferior-derecha]. Usado para resaltar secciones que el
    cliente debe completar a mano (ej: Régimen Patrimonial si es casado).
    """
    from reportlab.lib.colors import Color
    x1, y1 = gx(col1), gy(row1)
    x2, y2 = gx(col2), gy(row2)
    x, y = min(x1, x2), min(y1, y2)
    w, h = abs(x2 - x1), abs(y2 - y1)

    cv.saveState()
    if fill_color is not None:
        fr, fg, fb = fill_color.red, fill_color.green, fill_color.blue
        cv.setFillColor(Color(fr, fg, fb, alpha=alpha))
        cv.setStrokeColor(Color(stroke_color.red, stroke_color.green, stroke_color.blue, alpha=stroke_alpha))
        cv.setLineWidth(line_width)
        cv.rect(x, y, w, h, fill=1, stroke=1)
    else:
        cv.setStrokeColor(Color(stroke_color.red, stroke_color.green, stroke_color.blue, alpha=stroke_alpha))
        cv.setLineWidth(line_width)
        cv.rect(x, y, w, h, fill=0, stroke=1)
    cv.restoreState()


# ── CONVERSIÓN NUMÉRICA ─────────────────────────────────────
def numero_a_letras(n):
    uns  = ['','UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
            'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS',
            'DIECISIETE','DIECIOCHO','DIECINUEVE']
    dec  = ['','DIEZ','VEINTE','TREINTA','CUARENTA','CINCUENTA',
            'SESENTA','SETENTA','OCHENTA','NOVENTA']
    cent = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
            'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']
    def h999(n):
        if n==0: return ''
        if n==100: return 'CIEN'
        c,r = divmod(n,100)
        s = (cent[c]+' ') if c else ''
        if r<20: s+=uns[r]
        else:
            d2,u2 = divmod(r,10)
            s+=dec[d2]+(' Y '+uns[u2] if u2 else '')
        return s.strip()
    if n==0: return 'CERO'
    p=[]
    mill,miles,resto = n//1_000_000,(n%1_000_000)//1_000,n%1_000
    if mill==1:  p.append('UN MILLÓN')
    elif mill>1: p.append(h999(mill)+' MILLONES')
    if miles==1:  p.append('MIL')
    elif miles>1: p.append(h999(miles)+' MIL')
    if resto: p.append(h999(resto))
    return ' '.join(p)

def monto_letras(s):
    try:
        v=float(str(s).replace(',','.'))
        ent,dec=int(v),round((v-int(v))*100)
        return f"{numero_a_letras(ent)} CON {dec:02d}/100"
    except: return str(s)

def fmt_num(s):
    try:
        v=float(str(s).replace(',','.'))
        ent,dec=int(v),round((v-int(v))*100)
        return f"$ {ent:,}".replace(',','.')+f",{dec:02d}"
    except: return str(s)

def split_text(texto, max_chars):
    """Divide en (linea1, linea2) respetando palabras."""
    if len(texto)<=max_chars: return texto,''
    c=texto[:max_chars].rfind(' ')
    if c<0: c=max_chars
    return texto[:c], texto[c+1:]

def add_spaces(s):
    if not s: return s
    # Prefijos comunes concatenados (más largos primero para evitar solapamiento)
    s = re.sub(r'(OESTE|NORTE|ESTE|SUR|BRIG|ALTE|CNEL|DEL|LAS|LOS|SAN|EL |LA |DE |AV |DR |GEN|ING)(?=[A-ZÁÉÍÓÚÑ])', r'\1 ', s)
    s = re.sub(r'([a-záéíóúñ])([A-ZÁÉÍÓÚÑ])', r'\1 \2', s)
    s = re.sub(r'([A-ZÁÉÍÓÚÑa-záéíóúñ])(\d)', r'\1 \2', s)
    # Colapsar múltiples espacios
    s = re.sub(r'  +', ' ', s)
    return s.strip()

def space_alnum(s):
    """Espaciado genérico letra↔número (ej: 'HILUXL/162.8DC4X2TDISRX' ->
    'HILUX L/16 2.8 DC 4 X 2 TDI SRX'). Usado para campos de modelo de
    vehículo que vienen concatenados del PDF."""
    if not s: return s
    s = re.sub(r'([A-Za-zÁÉÍÓÚáéíóúÑñ])(\d)', r'\1 \2', s)
    s = re.sub(r'(\d)([A-Za-zÁÉÍÓÚáéíóúÑñ])', r'\1 \2', s)
    return re.sub(r'\s+', ' ', s).strip()


# ── PARSER SOLICITUD ────────────────────────────────────────
def parsear_solicitud(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        t1 = pdf.pages[0].extract_text() or ''
        t2 = pdf.pages[1].extract_text() if len(pdf.pages)>1 else ''
        t4 = pdf.pages[3].extract_text() if len(pdf.pages)>3 else ''

    def ex(txt,pat,g=1):
        m=re.search(pat,txt,re.IGNORECASE)
        return m.group(g).strip() if m else ''

    def ex2(txt,pat,extra_flags=0,g=1):
        m=re.search(pat,txt,re.IGNORECASE|extra_flags)
        return m.group(g).strip() if m else ''

    nc=ex(t1,r'ApellidoyNombre:\s*([A-ZÁÉÍÓÚÑ ,]+?)\s*Fecha')
    nc = nc.strip() if nc else nc
    if ',' in nc:
        apellido,nombre=nc.split(',',1)
        apellido,nombre=apellido.strip(),add_spaces(nombre.strip())
    else:
        p=nc.split(); apellido=p[0]; nombre=add_spaces(' '.join(p[1:]))

    tipo_doc = ex(t1, r'TipoyN[°º]Documento:\s*([A-Za-z\.]+)') or 'DNI'
    dni      = ex(t1,r'DNI\s*(\d+)')
    cuit_raw = ex(t1,r'CUIT/CUIL:\s*([\d\-]+)')
    # Formatear CUIT: 20389577627 → 20-38957762-7
    if cuit_raw and '-' not in cuit_raw and len(cuit_raw)==11:
        cuit = f"{cuit_raw[:2]}-{cuit_raw[2:10]}-{cuit_raw[10]}"
    else:
        cuit = cuit_raw

    fecha_nac = ex(t1,r'FechaNacimiento:\s*(\d{2}/\d{2}/\d{4})')
    est_civil = ex(t1,r'EstadoCivil:\s*(\w[\w/]*)').replace('/a','').strip()
    nac       = ex(t1,r'Nacionalidad:\s*(\w+)\s*Pais')
    activ_raw = ex(t1,r'Actividad\s*Principal:\s*([^\n]+?)(?:Antig|$)').split('Antig')[0].strip()
    activ     = add_spaces(activ_raw) if activ_raw else 'Autónomo'

    # Aplanar saltos de línea para capturar domicilios que se parten en dos líneas
    t1_flat = t1.replace('\n', ' ')
    dom_raw = ex(t1_flat,r'Domicilio:\s*(.+?)\s*Piso:')
    # dom_raw ya está acotado hasta "Piso:", se puede usar regex permisivo
    # para separar calle de número. El Nº: con dos puntos identifica el número.
    dm=re.match(r'(.+?)\s*N[º°]:\s*(\d+)',dom_raw) if dom_raw else None
    dom_calle = add_spaces(dm.group(1).strip()) if dm else add_spaces(ex(t1_flat,r'Domicilio:\s*(.+?)\s*N[º°]'))
    dom_num   = dm.group(2).strip() if dm else ex(t1_flat,r'Domicilio:.*?N[º°]:\s*(\d+)')

    # Si alguna "palabra" del dom_calle es demasiado larga para ser una sola
    # palabra real (> 8 chars sin espacio), reconstruir con x_tolerance=1.5
    # que separa correctamente (ej: "REPÚBLICADEL" → "REPÚBLICA DEL").
    if dom_calle and any(len(w) > 8 for w in dom_calle.split()):
        try:
            words_low = pdf.pages[0].extract_words(x_tolerance=1.5)
            for idx, w in enumerate(words_low):
                if w['text'].rstrip(':') in ('Domicilio', 'Domicilio:'):
                    parts = []
                    for j in range(idx+1, min(idx+15, len(words_low))):
                        wt = words_low[j]['text']
                        if re.search(r'^(Nº|N°|Nro|Piso|Depto|Código|CódigoPostal)', wt):
                            break
                        if re.match(r'^[A-ZÁÉÍÓÚÜÑ0-9]+$', wt):
                            parts.append(wt)
                        else:
                            break
                    if len(parts) >= 2:
                        dom_calle = ' '.join(parts)
                    break
        except Exception:
            pass  # Fallback al valor extraído originalmente
    dom_piso  = ex(t1,r'Piso:\s*(\d+)')
    # Depto puede venir vacío ("Piso:Depto:CódigoPostal..."); solo capturar
    # si hay un valor real antes del próximo label conocido.
    depto_m = re.search(r'Depto:\s*([A-Za-z0-9]*?)(?=CódigoPostal|Localidad|Provincia|\n|$)', t1)
    dom_depto = depto_m.group(1).strip() if depto_m else ''
    cod_postal = ex(t1,r'C[óo]digoPostal:\s*(\d+)')
    localidad = ex(t1,r'Localidad:\s*([^\s].+?)\s*Provincia').strip()
    # Fix concatenaciones (ej: "QUILMESOESTE" -> "QUILMES OESTE")
    if localidad and any(len(w) > 8 for w in localidad.split()):
        try:
            if 'words_low' not in dir():
                words_low = pdf.pages[0].extract_words(x_tolerance=1.5)
            for idx, w in enumerate(words_low):
                if w['text'].rstrip(':') in ('Localidad', 'Localidad:'):
                    parts = []
                    for j in range(idx+1, min(idx+8, len(words_low))):
                        wt = words_low[j]['text']
                        if re.search(r'^(Provincia|Provincia:|Buenos|Capital|Código)', wt):
                            break
                        if re.match(r'^[A-ZÁÉÍÓÚÜÑ0-9]+$', wt):
                            parts.append(wt)
                        else:
                            break
                    if len(parts) >= 2:
                        localidad = ' '.join(parts)
                    break
        except Exception:
            pass
    provincia = add_spaces(ex(t1,r'Provincia:\s*([A-ZÁÉÍÓÚÑA-Za-z ]+?)(?:Tel|Ingr|$)').strip())

    tna       = ex(t1,r'TNA:\s*([\d.,]+)%').replace(',','.')
    importe   = ex(t1,r'Importedel[Pp]r[eé]stamo:\s*\$\s*([\d.,]+)')
    uvas      = ex(t1,r'UVAsEquivalentes:\s*([\d.,]+)')
    monto_ins = ex(t1,r'MontoInsc\.Prenda:\s*\$\s*([\d.,]+)')
    cuotas    = ex(t1,r'CantidaddeCuotas:\s*(\d+)')
    prim_venc = ex(t1,r'PrimerVencimiento:\s*(\d{2}/\d{2}/\d{4})')
    tasa_tipo = ex(t1,r'TasadeInter[eé]s:\s*(\w+)').upper() or 'FIJA'
    cuota_pura= ex(t1,r'CuotaPura:\s*\$\s*([\d.,]+)')
    cftea_raw = ex(t1,r'CFTEA:\s*([\d.,]+)%').replace(',','.')
    uva_valor = ex(t1,r'ValorUVAsDía:\s*([\d.,]+)')

    marca  = ex(t1,r'Marca:\s*([A-Z]+)')
    mod_m  = re.search(r'Modelo:\s*([^\n]+)',t1)
    modelo = space_alnum(re.split(r'(?:Valor|Tipo:|N[º°])',mod_m.group(1))[0].strip()) if mod_m else ''
    anio   = ex(t1,r'A[ñn]o:\s*(\d{4})')
    nuevo_m= re.search(r'Nuevo:\s*(SI|NO)',t1,re.I)
    categoria='AUTOMÓVIL NUEVO' if (nuevo_m and nuevo_m.group(1).upper()=='SI') else 'AUTOMÓVIL USADO'
    chasis = ex(t1,r'N[º°]\s*Chasis:\s*([A-Z0-9]+)')
    motor  = ex(t1,r'N[º°]\s*Motor:\s*([A-Z0-9]+)')
    uso    = ex(t1,r'Uso:\s*(\w+)')
    dominio= ex(t4,r'Dominio\s*([A-Z]{2}\d{3}[A-Z]{2}|[A-Z]{3}\d{3})')
    if not dominio: dominio=ex(t4,r'PrendariaDominio\s*([A-Z0-9]+)')
    # Fallback: buscar en todas las páginas (ej. solicitudes sin TyC donde
    # la Entrega de Fondos está en página 3, no 4)
    if not dominio:
        for i in range(len(pdf.pages)):
            t_tmp = pdf.pages[i].extract_text() or ''
            dominio = ex(t_tmp,r'PrendariaDominio\s*([A-Z0-9]+)')
            if not dominio: dominio = ex(t_tmp,r'PrendariaDominio([A-Z]{2}\d{3}[A-Z]{2})')
            if dominio: break

    # Cónyuge
    # Formato nuevo (SolicitudSinTyC): sección "DATOSCONYUGE" en página 1 con
    # "...ApellidoyNombre:MAYOCCHI,LORENA" y "TipoyNºDocumento:DNI24704151".
    conyuge_nc = ex2(t1, r'DATOSCONYUGE.*?ApellidoyNombre:\s*([A-ZÁÉÍÓÚÑ ,]+)', re.DOTALL)
    if not conyuge_nc:
        # Formato anterior: campos sueltos NombreConyuge / ApellidoConyuge,
        # o tabla en página 2 (hoja del cónyuge del contrato modelo).
        conyuge_nomb = ex(t1,r'NombreConyuge:\s*([^\n]+)') or ex(t2,r'Nombres?:\s*([^\n]+)')
        conyuge_apel = ex(t1,r'ApellidoConyuge:\s*([^\n]+)') or ex(t2,r'Apellidos?:\s*([^\n]+)')
    elif ',' in conyuge_nc:
        conyuge_apel, conyuge_nomb = conyuge_nc.split(',',1)
        conyuge_apel, conyuge_nomb = conyuge_apel.strip(), add_spaces(conyuge_nomb.strip())
    else:
        p=conyuge_nc.split(); conyuge_apel=p[0]; conyuge_nomb=add_spaces(' '.join(p[1:]))

    conyuge_dni_m = re.search(r'DATOSCONYUGE.*?TipoyN[ºo°]Documento:\s*[A-Za-zÁÉÍÓÚáéíóú]*(\d+)', t1, re.IGNORECASE|re.DOTALL)
    conyuge_dni = conyuge_dni_m.group(1).strip() if conyuge_dni_m else (
        ex(t1,r'DocumentoConyuge:\s*([\w ]+)') or ex(t2,r'DNI:\s*(\d+)'))
    nupcias_m    = re.search(r'(primera|segunda)\s*nupcias',t1,re.I)
    nupcias      = nupcias_m.group(1).upper() if nupcias_m else 'PRIMERA'

    try:
        d2,mo,yy=fecha_nac.split('/')
        hoy=date.today()
        edad=str(hoy.year-int(yy)-((hoy.month,hoy.day)<(int(mo),int(d2))))
    except: edad=''

    try:
        cftea_ent=str(int(float(cftea_raw)))
        cftea_dec=f"{round((float(cftea_raw)%1)*100):02d}"
    except: cftea_ent=cftea_dec=''

    return {
        'nombre_completo':  f"{apellido}, {nombre}".strip(', '),
        'apellido': apellido, 'nombre': nombre,
        'dni': dni, 'tipo_doc': tipo_doc, 'cuit': cuit,
        'fecha_nac': fecha_nac, 'edad': edad,
        'estado_civil': est_civil, 'nacionalidad': nac, 'actividad': activ,
        'dom_calle': dom_calle, 'dom_num': dom_num,
        'dom_piso': dom_piso, 'dom_depto': dom_depto,
        'dom_completo': f"{dom_calle} Nº: {dom_num}  {localidad}  {provincia}",
        'cod_postal': cod_postal,
        'localidad': localidad, 'provincia': provincia,
        'tna': tna, 'importe': importe, 'uvas': uvas,
        'monto_insc': monto_ins, 'cuotas': cuotas,
        'prim_venc': prim_venc, 'tasa_tipo': tasa_tipo,
        'cuota_pura': cuota_pura,
        'cftea': cftea_raw, 'cftea_ent': cftea_ent, 'cftea_dec': cftea_dec,
        'uva_valor': uva_valor,
        'marca': marca, 'modelo': modelo, 'anio': anio,
        'categoria': categoria, 'chasis': chasis, 'motor': motor,
        'uso': uso, 'dominio': dominio,
        'conyuge_nombre': conyuge_nomb, 'conyuge_apellido': conyuge_apel,
        'conyuge_dni': conyuge_dni,
        'tiene_conyuge': bool(conyuge_nomb or conyuge_apel or 'CASADO' in est_civil.upper()),
        'nupcias': nupcias,
        'fecha_hoy': date.today().strftime('%d/%m/%Y'),
        # Marca Motor / Marca Chasis no se pueden leer automáticamente
        # (el Título del Automotor se adjunta como imagen, sin OCR en pipeline).
        # Deben verificarse y completarse a mano contra el Título adjunto.
        'marca_motor': 'COMPLETAR CORRECTAMENTE',
        'marca_chasis': 'COMPLETAR CORRECTAMENTE',
        # Fecha de la Carta de Aprobación (para Hoja 3, fila ~85,5)
        'fecha_carta_dd': '', 'fecha_carta_mm': '', 'fecha_carta_aa': '',
    }


# ── PARSER CARTA DE APROBACIÓN ──────────────────────────────
def parsear_carta_aprobacion(pdf_path):
    """
    Extrae la fecha de la Carta de Aprobación (formato 'Fecha:DD/MM/AAAA' o
    'Fecha: DD/MM/AAAA'). Devuelve dict con fecha_carta_dd y fecha_carta_aa
    (últimos 2 dígitos del año). Si no se puede leer, devuelve campos vacíos.

    También extrae datos del vehículo (Marca/Modelo/Año/Nuevo/Uso) como
    fallback para versiones de la Solicitud que no incluyen la sección
    "DESTINO DEL PRÉSTAMO - INFORMACIÓN DEL AUTOMOTOR" (ej: SolicitudSinTyC).
    Estos se devuelven con prefijo 'carta_' y solo se usan si el dato
    correspondiente vino vacío de la Solicitud.
    """
    vacio_fecha = {'fecha_carta_dd': '', 'fecha_carta_mm': '', 'fecha_carta_aa': ''}
    vacio_veh = {'carta_marca': '', 'carta_modelo': '', 'carta_anio': '',
                  'carta_categoria': '', 'carta_uso': ''}
    if not pdf_path or not os.path.exists(pdf_path):
        return {**vacio_fecha, **vacio_veh}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            t = pdf.pages[0].extract_text() or ''
    except Exception:
        return {**vacio_fecha, **vacio_veh}

    m = re.search(r'Fecha:\s*(\d{2})/(\d{2})/(\d{4})', t)
    if m:
        dd, mm, yyyy = m.groups()
        out_fecha = {'fecha_carta_dd': dd, 'fecha_carta_mm': mm, 'fecha_carta_aa': yyyy[-2:]}
    else:
        out_fecha = vacio_fecha

    carta_marca = ''
    carta_modelo = ''
    carta_anio = ''
    carta_categoria = ''
    carta_uso = ''

    mm_ = re.search(r'Marca:\s*([A-Z]+)', t)
    if mm_: carta_marca = mm_.group(1).strip()

    mod_m = re.search(r'Modelo:\s*([^\n]+)', t)
    if mod_m:
        carta_modelo = space_alnum(re.split(r'(?:A[ñn]o|Valor|Tipo|N[º°])', mod_m.group(1))[0].strip())

    an_m = re.search(r'A[ñn]o:\s*(\d{4})\s*Nuevo:\s*(SI|NO)', t, re.IGNORECASE)
    if an_m:
        carta_anio = an_m.group(1)
        carta_categoria = 'AUTOMÓVIL NUEVO' if an_m.group(2).upper()=='SI' else 'AUTOMÓVIL USADO'

    uso_m = re.search(r'Uso:\s*([A-Za-zÁÉÍÓÚáéíóú]+?)(?:Valor|$)', t)
    if uso_m: carta_uso = uso_m.group(1).strip()

    out_veh = {'carta_marca': carta_marca, 'carta_modelo': carta_modelo,
               'carta_anio': carta_anio, 'carta_categoria': carta_categoria,
               'carta_uso': carta_uso}
    return {**out_fecha, **out_veh}


# ── PARSER MUTUO PRENDARIO (casos Pre-prenda) ───────────────
def parsear_mutuo_prendario(pdf_path):
    """
    Extrae datos del Mutuo Prendario, usado en casos de Pre-prenda (*_PRE).
    Trae su propio "corte" de UVAs más cercano a la fecha de desembolso:
      - UVAS EQUIVALENTE
      - FECHA VALOR UVA (dd/mm/aaaa)
      - COTIZACION UVA
      - MONTO DE LA PRENDA
      - Primer vencimiento (texto "La 1° cuota vence el dd/mm/aaaa")
    Si no se puede leer, devuelve campos vacíos (el resolver hará fallback
    silencioso a cadenas vacías).
    """
    vacio = {
        'mutuo_uvas': '', 'mutuo_uva_valor': '',
        'mutuo_fecha_dd': '', 'mutuo_fecha_mm': '', 'mutuo_fecha_aa': '',
        'mutuo_monto_prenda': '', 'mutuo_prim_venc': '',
    }
    if not pdf_path or not os.path.exists(pdf_path):
        return vacio
    try:
        with pdfplumber.open(pdf_path) as pdf:
            t = pdf.pages[0].extract_text() or ''
    except Exception:
        return vacio

    def ex(pat, g=1):
        m = re.search(pat, t, re.IGNORECASE)
        return m.group(g).strip() if m else ''

    uvas      = ex(r'UVASEQUIVALENTE:\s*([\d.,]+)')
    uva_valor = ex(r'COTIZACIONUVA:\s*\$\s*([\d.,]+)')
    monto_pr  = ex(r'MONTODELAPRENDA:\s*\$\s*([\d.,]+)')
    prim_venc = ex(r'La1[ºo°]?\s*cuotavenceel\s*(\d{2}/\d{2}/\d{4})')

    fecha_dd = fecha_mm = fecha_aa = ''
    fm = re.search(r'FECHAVALORUVA:\s*(\d{2})/(\d{2})/(\d{4})', t, re.IGNORECASE)
    if fm:
        fecha_dd, fecha_mm, yyyy = fm.groups()
        fecha_aa = yyyy[-2:]

    return {
        'mutuo_uvas': uvas, 'mutuo_uva_valor': uva_valor,
        'mutuo_fecha_dd': fecha_dd, 'mutuo_fecha_mm': fecha_mm, 'mutuo_fecha_aa': fecha_aa,
        'mutuo_monto_prenda': monto_pr, 'mutuo_prim_venc': prim_venc,
    }


# ── RESOLUCIÓN DE VARIABLES ─────────────────────────────────
def resolver(variable_id, d):
    """
    Recibe el Variable_ID del Excel y los datos parseados.
    Devuelve una lista de (col, row, texto) listos para dibujar.
    Coordenadas multi-línea se expresan como lista de tuplas.
    """
    hoy = d['fecha_hoy']
    try: dd,mm,yy = hoy.split('/')
    except: dd=mm=yy='01'
    aa = yy[-2:]
    mes_txt = MESES[int(mm)]

    R = []  # lista de (col, row, texto, size)

    def add(col, row, texto, size=6.5): R.append((col, row, texto, size))

    # ── HOJA 1 ──────────────────────────────────────────────
    if variable_id == 'H1_PROVINCIA_FECHA':
        # Fecha de encabezado: queda "XX/XX" literal para que el cliente
        # la complete a mano el día que entrega la documentación.
        add(46, 15, f"{d['provincia']},  XX/XX")

    elif variable_id == 'H1_ANIO':
        add(69, 15, yy)

    elif variable_id == 'H1_MONTO_NUM':
        add(23, 20, fmt_num(d['monto_insc']), 7)

    elif variable_id == 'H1_MONTO_LETRAS':
        letras = monto_letras(d['monto_insc'])
        l1, l2 = split_text(letras, 52)
        add(40,   21,   l1)
        add(18,   23.5, l2)

    elif variable_id == 'H1_NOMBRE_DEUDOR':
        n1, n2 = split_text(d['nombre_completo'], 28)
        add(52,   23.5, n1)
        if n2: add(18, 26, n2)

    elif variable_id == 'H1_ACREEDOR_FIJO_28':
        add(44, 28, 'BANCO SUPERVIELLE S.A.')

    elif variable_id == 'H1_BIEN':
        bien = (f"{d['categoria']}  Marca: {d['marca']}  Modelo: {d['modelo']}  "
                f"Marca Motor: {d['marca_motor']}  Nº Motor: {d['motor']}  "
                f"Marca Chasis: {d['marca_chasis']}  Nº Chasis: {d['chasis']}  "
                f"Dominio: {d['dominio']}  Año: {d['anio']}  Uso: {d['uso']}")
        rows_bien = [35, 37, 39, 41]
        resto = bien
        for i, max_c in enumerate([68, 68, 68, 56]):
            if not resto: break
            if len(resto) <= max_c:
                add(18, rows_bien[i], resto); resto=''; break
            c = resto[:max_c].rfind(' ')
            if c<0: c=max_c
            add(18, rows_bien[i], resto[:c])
            resto = resto[c+1:]

    elif variable_id == 'H1_PROVINCIA_BIEN':
        add(35, 42, d['provincia'])

    elif variable_id == 'H1_LOCALIDAD_UBICACION':
        add(21, 44, d['localidad'])

    elif variable_id == 'H1_CIUDAD':
        add(35, 47, d['localidad'])

    elif variable_id == 'H1_CALLE':
        add(48, 47, d['dom_calle'])

    elif variable_id == 'H1_DOM_NUM':
        add(65, 47, d['dom_num'])

    elif variable_id == 'H1_CUOTAS':
        txt = (f"{d['cuotas']} cuotas mensuales, iguales y consecutivas de "
               f"{fmt_num(d['cuota_pura'])} + IVA")
        l1, l2 = split_text(txt, 52)
        add(35, 53, l1)
        if l2: add(18, 55, l2)

    elif variable_id == 'H1_VENCIMIENTOS':
        # Para *_PRE el primer vencimiento se toma del Mutuo Prendario
        # (misma fecha que la Solicitud en la práctica, distinta fuente).
        prim_venc = d['prim_venc']
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_prim_venc'):
            prim_venc = d['mutuo_prim_venc']
        txt = (f"Venciendo la primera el día {prim_venc} y las restantes el mismo día "
               f"o primer día hábil posterior de los meses subsiguientes hasta la total y "
               f"definitiva cancelación de la deuda")
        l1, l2 = split_text(txt, 58)
        add(32, 57, l1)
        add(18, 59, l2)

    elif variable_id == 'H1_TNA':
        add(40, 62, d['tna'])

    elif variable_id == 'H1_CUIT_BANCO':
        add(24, 71.5, '33-50000517-9')

    elif variable_id == 'H1_CUIT_DEUDOR':
        add(52, 71.5, f"CUIT/CUIL: {d['cuit']}")

    elif variable_id == 'H1_ACREEDOR_FIJO_73':
        add(26, 73, 'BANCO SUPERVIELLE S.A.')

    elif variable_id == 'H1_NOMBRE_DEUDOR_73':
        add(56, 73, d['nombre_completo'])

    elif variable_id == 'H1_ESTADO_CIVIL':
        add(53, 76, d['estado_civil'])

    elif variable_id == 'H1_ACTIVIDAD':
        add(85, 76, d['actividad'])

    elif variable_id == 'H1_NACIONALIDAD':
        add(53, 78, d['nacionalidad'])

    elif variable_id == 'H1_EDAD':
        add(66, 78, d['edad'])

    elif variable_id == 'H1_DOM_BANCO':
        add(23, 80, 'RECONQUISTA 330 (C1003ABH) C.A.B.A.')

    elif variable_id == 'H1_DOM_DEUDOR':
        add(53, 80, d['dom_completo'])

    elif variable_id == 'H1_IGJ':
        add(30, 82, 'IGJ N°7333 L28 TdeS por A 27-06-05', 6)

    elif variable_id == 'H1_DNI_DEUDOR':
        add(55, 82, f"DNI {d['dni']}")

    elif variable_id == 'H1_FIRMA_DEUDOR':
        add(60, 86, 'X', 14)

    # ── HOJA 2 ──────────────────────────────────────────────
    elif variable_id == 'H2_NOMBRE_SOLICITANTE':
        if d['tiene_conyuge']: add(27, 24, d['nombre_completo'])

    elif variable_id == 'H2_NUPCIAS_PRIMERA':
        if d['tiene_conyuge'] and d['nupcias'] != 'SEGUNDA':
            add(16, 25.5, 'X', 10)

    elif variable_id == 'H2_NUPCIAS_SEGUNDA':
        if d['tiene_conyuge'] and d['nupcias'] == 'SEGUNDA':
            add(16, 26.5, 'X', 10)

    elif variable_id == 'H2_CONYUGE_NOMBRE':
        if d['tiene_conyuge']: add(20, 31, d['conyuge_nombre'])

    elif variable_id == 'H2_CONYUGE_APELLIDO':
        if d['tiene_conyuge']: add(20, 33, d['conyuge_apellido'])

    elif variable_id == 'H2_CONYUGE_DNI':
        if d['tiene_conyuge']: add(20, 35, d['conyuge_dni'])

    elif variable_id == 'H2_FIRMA_CONYUGE':
        if d['tiene_conyuge']: add(50, 35, 'X', 14)

    elif variable_id == 'H2_FIRMA_DEUDOR':
        if d['tiene_conyuge']: add(61, 60, 'X', 14)

    # ── HOJA 3 ──────────────────────────────────────────────
    elif variable_id == 'H3_MONTO_NUM':
        add(16, 5, fmt_num(d['monto_insc']), 7)

    elif variable_id == 'H3_FECHA_DD':
        # Fecha de encabezado del contrato: queda "XX" literal para que
        # el cliente la complete a mano el día que entrega la documentación.
        add(33, 5, 'XX')

    elif variable_id == 'H3_FECHA_MES':
        add(36, 5, 'XX')

    elif variable_id == 'H3_FECHA_AA':
        add(47, 5, 'XX')

    elif variable_id == 'H3_ACREEDOR_FIJO':
        add(16, 6.5, 'BANCO SUPERVIELLE S.A.')

    elif variable_id == 'H3_NOMBRE_DEUDOR':
        add(16, 8, d['nombre_completo'])

    elif variable_id == 'H3_CFTEA_ENT':
        add(44, 82, d['cftea_ent'])

    elif variable_id == 'H3_CFTEA_DEC':
        add(48, 82, d['cftea_dec'])

    elif variable_id == 'H3_IMPORTE_LETRAS':
        letras = monto_letras(d['importe'])
        l1, l2 = split_text(letras, 40)
        add(42,   83,   l1)
        add(13,   84.5, l2)

    elif variable_id == 'H3_IMPORTE_NUM':
        add(32, 84.5, fmt_num(d['importe']))

    elif variable_id == 'H3_UVAS_LETRAS':
        # Para *_PRE, las UVAs equivalentes vienen del Mutuo Prendario
        # (UVAS EQUIVALENTE), no de la Solicitud.
        uvas = d['uvas']
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_uvas'):
            uvas = d['mutuo_uvas']
        letras = monto_letras(uvas)
        l1, l2 = split_text(letras, 34)
        add(46,   84.5, l1)
        add(13,   85.5, l2)

    elif variable_id == 'H3_UVAS_NUM':
        uvas = d['uvas']
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_uvas'):
            uvas = d['mutuo_uvas']
        add(32, 85.5, uvas)

    elif variable_id == 'H3_FECHA85_DD':
        # UVA_PI / FIJA_PI: fecha de la Carta de Aprobación.
        # *_PRE: FECHA VALOR UVA del Mutuo Prendario.
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_fecha_dd'):
            add(45, 85.5, d['mutuo_fecha_dd'])
        else:
            add(45, 85.5, d['fecha_carta_dd'])

    elif variable_id == 'H3_FECHA85_MM':
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_fecha_mm'):
            add(47.5, 85.5, d['mutuo_fecha_mm'])
        else:
            add(47.5, 85.5, d['fecha_carta_mm'])

    elif variable_id == 'H3_FECHA85_AA':
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_fecha_aa'):
            add(49, 85.5, d['mutuo_fecha_aa'])
        else:
            add(49, 85.5, d['fecha_carta_aa'])

    elif variable_id == 'H3_UVA_VALOR_LETRAS':
        # Para *_PRE, el valor de la UVA del día viene de COTIZACION UVA
        # del Mutuo Prendario, no del "Valor UVAs Día" de la Solicitud.
        uva_valor = d['uva_valor']
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_uva_valor'):
            uva_valor = d['mutuo_uva_valor']
        letras = monto_letras(uva_valor)
        # Ancho disponible en la fila 86.5 es de col 32 a col 49 (17 columnas).
        # Si el texto no entra, continúa en la fila siguiente (87), col 13,
        # siguiendo la misma convención usada por H3_IMPORTE_LETRAS / H3_UVAS_LETRAS.
        max_c1 = 17
        if len(letras) <= max_c1:
            add(32, 86.5, letras)
        else:
            l1, l2 = split_text(letras, max_c1)
            add(32, 86.5, l1)
            add(13, 87, l2)

    elif variable_id == 'H3_UVA_VALOR_NUM':
        uva_valor = d['uva_valor']
        if d['tipo_op'].endswith('_PRE') and d.get('mutuo_uva_valor'):
            uva_valor = d['mutuo_uva_valor']
        add(49, 86.5, uva_valor)

    # ── HOJA 4 ──────────────────────────────────────────────
    elif variable_id == 'H4_CUOTAS':
        add(61, 7.5, d['cuotas'])

    elif variable_id == 'H4_FIRMA_DEUDOR':
        add(30, 110, 'X', 14)

    elif variable_id == 'H4_FIRMA_CONYUGE':
        add(44, 110, 'X', 14)

    # ── HOJA 5 ──────────────────────────────────────────────
    elif variable_id == 'H5_MONTO_NUM':
        add(18, 18.5, fmt_num(d['monto_insc']), 7)

    elif variable_id == 'H5_FECHA_DD':
        # Fecha de encabezado: queda "XX" literal para completar a mano.
        add(33, 18.5, 'XX')

    elif variable_id == 'H5_FECHA_MES':
        add(37, 18.5, 'XX')

    elif variable_id == 'H5_FECHA_AA':
        add(48, 18.5, 'XX')

    elif variable_id == 'H5_ACREEDOR_FIJO':
        add(15, 20, 'BANCO SUPERVIELLE S.A.')

    elif variable_id == 'H5_NOMBRE_DEUDOR':
        add(15, 22, d['nombre_completo'])

    elif variable_id == 'H5_DOM_BANCO':
        add(13, 91, 'RECONQUISTA 330 (C1003ABH) C.A.B.A.')

    elif variable_id == 'H5_DOM_DEUDOR':
        add(14, 92, d['dom_completo'])

    elif variable_id == 'H5_FIRMA_DEUDOR':
        add(30, 100, 'X', 14)

    elif variable_id == 'H5_FIRMA_CONYUGE':
        add(44, 100, 'X', 14)

    elif variable_id == 'H5_CONYUGE_BOX':
        # Si el deudor es casado, resalta con un recuadro semitransparente
        # la sección "RÉGIMEN PATRIMONIAL" (declaraciones a completar a mano:
        # comunidad/separación de bienes, unión convivencial, etc.)
        if d['tiene_conyuge']:
            R.append((12, 40, '__BOX__', 50, 49))

    # ── HOJA 6 ──────────────────────────────────────────────
    elif variable_id == 'H6_ACREEDOR_FIJO':
        add(15, 19, 'BANCO SUPERVIELLE S.A.')

    elif variable_id == 'H6_NOMBRE_DEUDOR':
        add(15, 21, d['nombre_completo'])

    elif variable_id == 'H6_FIRMA_DEUDOR':
        add(64, 108, 'X', 14)

    else:
        print(f"  ⚠ Variable_ID desconocida: '{variable_id}'")

    return R



# ── MAPEO COORDENADAS → VARIABLE_ID ─────────────────────────
# Clave: (hoja_int, coords_string)  →  Variable_ID del resolver
# Normalizar claves del dict para que usen \n literal (2 chars) en vez de newline real
def _build_coords_vid():
    raw = {
    # HOJA 1
    (1, '46.15'):                                    'H1_PROVINCIA_FECHA',
    (1, '69.15'):                                    'H1_ANIO',
    (1, '23.20'):                                    'H1_MONTO_NUM',
    (1, '40.21–75.21\n18.23,5–49.23,5'):             'H1_MONTO_LETRAS',
    (1, '52.23,5–75.23,5\n18.26–65.26'):             'H1_NOMBRE_DEUDOR',
    (1, '44.28'):                                    'H1_ACREEDOR_FIJO_28',
    (1, '18.35–75.35\n18.37–75.37\n18.39–75.39\n18.41–65.41'): 'H1_BIEN',
    (1, '35.42'):                                    'H1_PROVINCIA_BIEN',
    (1, '21.44'):                                    'H1_LOCALIDAD_UBICACION',
    (1, '35.47'):                                    'H1_CIUDAD',
    (1, '48.47'):                                    'H1_CALLE',
    (1, '65.47'):                                    'H1_DOM_NUM',
    (1, '35.53–75.53\n18.55–75.55'):                 'H1_CUOTAS',
    (1, '32.57–75.57\n18.59–75.59'):                 'H1_VENCIMIENTOS',
    (1, '40.62'):                                    'H1_TNA',
    (1, '24.71,5–45.71,5'):                          'H1_CUIT_BANCO',
    (1, '52.71,5'):                                  'H1_CUIT_DEUDOR',
    (1, '26.73'):                                    'H1_ACREEDOR_FIJO_73',
    (1, '56.73'):                                    'H1_NOMBRE_DEUDOR_73',
    (1, '53.76'):                                    'H1_ESTADO_CIVIL',
    (1, '85.76'):                                    'H1_ACTIVIDAD',
    (1, '53.78'):                                    'H1_NACIONALIDAD',
    (1, '66.78'):                                    'H1_EDAD',
    (1, '23.80'):                                    'H1_DOM_BANCO',
    (1, '53.80'):                                    'H1_DOM_DEUDOR',
    (1, '30.82–46.82'):                              'H1_IGJ',
    (1, '55.82'):                                    'H1_DNI_DEUDOR',
    (1, '60.86'):                                    'H1_FIRMA_DEUDOR',
    # HOJA 2
    (2, '—'):                                        'H2_CONDICIONAL',
    (2, '27.24'):                                    'H2_NOMBRE_SOLICITANTE',
    (2, '16.25,5'):                                  'H2_NUPCIAS_PRIMERA',
    (2, '16.26,5'):                                  'H2_NUPCIAS_SEGUNDA',
    (2, '20.31'):                                    'H2_CONYUGE_NOMBRE',
    (2, '20.33'):                                    'H2_CONYUGE_APELLIDO',
    (2, '20.35'):                                    'H2_CONYUGE_DNI',
    (2, '50.35'):                                    'H2_FIRMA_CONYUGE',
    (2, '61.60'):                                    'H2_FIRMA_DEUDOR',
    # HOJA 3
    (3, '16.5'):                                     'H3_MONTO_NUM',
    (3, '33.5'):                                     'H3_FECHA_DD',
    (3, '36.5'):                                     'H3_FECHA_MES',
    (3, '47.5'):                                     'H3_FECHA_AA',
    (3, '16.6,5'):                                   'H3_ACREEDOR_FIJO',
    (3, '16.8'):                                     'H3_NOMBRE_DEUDOR',
    (3, '44.82'):                                    'H3_CFTEA_ENT',
    (3, '48.82'):                                    'H3_CFTEA_DEC',
    (3, '42.83–70.83\n13.84,5–30.84,5'):             'H3_IMPORTE_LETRAS',
    (3, '32.84,5–39.84,5'):                          'H3_IMPORTE_NUM',
    (3, '46.84,5–70.84,5\n13.85,5–30.85,5'):         'H3_UVAS_LETRAS',
    (3, '32.85,5–39.85,5'):                          'H3_UVAS_NUM',
    (3, '45.85,5–47.85,5'):                          'H3_FECHA85_DD',
    (3, '47,5.85,5–48,5.85,5'):                      'H3_FECHA85_MM',
    (3, '49.85,5–52.85,5'):                          'H3_FECHA85_AA',
    (3, '32.86,5–48.86,5'):                          'H3_UVA_VALOR_LETRAS',
    (3, '49.86,5–54.86,5'):                          'H3_UVA_VALOR_NUM',
    # HOJA 4
    (4, '61.7,5–63.7,5'):                            'H4_CUOTAS',
    (4, '30.110'):                                   'H4_FIRMA_DEUDOR',
    (4, '44.110'):                                   'H4_FIRMA_CONYUGE',
    # HOJA 5
    (5, '18.18,5'):                                  'H5_MONTO_NUM',
    (5, '33.18,5'):                                  'H5_FECHA_DD',
    (5, '37.18,5'):                                  'H5_FECHA_MES',
    (5, '48.18,5'):                                  'H5_FECHA_AA',
    (5, '15.20'):                                    'H5_ACREEDOR_FIJO',
    (5, '15.22'):                                    'H5_NOMBRE_DEUDOR',
    (5, '13.91'):                                    'H5_DOM_BANCO',
    (5, '14.92'):                                    'H5_DOM_DEUDOR',
    (5, '30.100'):                                   'H5_FIRMA_DEUDOR',
    (5, '44.100'):                                   'H5_FIRMA_CONYUGE',
    (5, '12.40–50.49'):                              'H5_CONYUGE_BOX',
    # HOJA 6
    (6, '15.19'):                                    'H6_ACREEDOR_FIJO',
    (6, '15.21'):                                    'H6_NOMBRE_DEUDOR',
    (6, '64.108'):                                   'H6_FIRMA_DEUDOR',
    }
    # Normalizar claves: reemplazar newline real con \\n
    return {(h, k.replace(chr(10), chr(92)+chr(110))): v for (h,k),v in raw.items()}
COORDS_TO_VID = _build_coords_vid()

# ── LEER CONFIGURACIÓN DEL EXCEL ────────────────────────────
def cargar_config_excel(xlsx_path, tipo_op):
    """
    Lee el Excel y devuelve lista de entradas activas para el tipo_op dado.
    Retorna: [ {hoja, variable_id, coords_str}, ... ]
    Solo incluye filas donde la columna del tipo_op NO es vacía ni '(por completar)'.
    """
    col_tipo = TIPO_COL[tipo_op]
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    entradas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Estructura: A=coords B=variable C=hoja D=UVA_PI E=UVA_PRE F=FIJA_PI G=FIJA_PRE
        if len(row) < 4: continue
        coords, variable, hoja = row[0], row[1], row[2]
        valores = row[3:]  # D en adelante

        if not coords or not hoja: continue
        if str(coords).strip().startswith('HOJA') or str(coords).strip().startswith('  HOJA'): continue
        try: hoja = int(float(hoja))
        except: continue

        # La columna del tipo_op: col_tipo=4 → índice 0 en valores
        val_idx = col_tipo - 4
        val = valores[val_idx] if val_idx < len(valores) else None
        # Nota: 'XX' / 'XX/XX' NO se excluyen — son placeholders literales que el
        # cliente completa a mano (fecha de entrega de documentación).
        if not val or str(val).strip() in ('', '(por completar)', 'N/A', 'None'): continue

        # Mapear coords → Variable_ID usando el diccionario
        coords_norm = str(coords).strip().replace(chr(10), chr(92)+chr(110))
        variable_id = COORDS_TO_VID.get((hoja, coords_norm)) or COORDS_TO_VID.get((hoja, str(coords).strip()))
        if not variable_id:
            continue  # Coordenada no mapeada, ignorar

        entradas.append({
            'hoja':        int(hoja),
            'variable_id': variable_id,
            'coords_str':  str(coords),
        })

    wb.close()

    # H5_CONYUGE_BOX: aparece en el Excel sin valor (fila de referencia visual).
    # Se agrega siempre; el resolver la dibuja solo si tiene_conyuge=True.
    entradas.append({'hoja': 5, 'variable_id': 'H5_CONYUGE_BOX', 'coords_str': '12.40–50.49'})

    return entradas


# ── GENERADOR PRINCIPAL ─────────────────────────────────────
def generar_pdf(solicitud_path, template_path, output_path,
                xlsx_path, tipo_op='UVA_PI', fecha_firma=None, carta_path=None, mutuo_path=None):

    if tipo_op not in TIPO_COL:
        print(f"⚠ Tipo '{tipo_op}' no válido. Usando UVA_PI.")
        tipo_op = 'UVA_PI'

    print(f"Tipo de operación : {tipo_op}")
    print(f"Config Excel      : {xlsx_path}")
    print(f"Parseando solicitud...")
    d = parsear_solicitud(solicitud_path)
    if fecha_firma: d['fecha_hoy'] = fecha_firma
    d['tipo_op'] = tipo_op

    print("Parseando carta de aprobación...")
    carta_data = parsear_carta_aprobacion(carta_path)
    d.update(carta_data)

    # Fallback: si la Solicitud no trae los datos del vehículo (ej. versión
    # "SolicitudSinTyC"), completar desde la Carta de Aprobación.
    if not d.get('marca') and carta_data.get('carta_marca'):
        d['marca'] = carta_data['carta_marca']
    if not d.get('modelo') and carta_data.get('carta_modelo'):
        d['modelo'] = carta_data['carta_modelo']
    if not d.get('anio') and carta_data.get('carta_anio'):
        d['anio'] = carta_data['carta_anio']
    if not d.get('categoria') and carta_data.get('carta_categoria'):
        d['categoria'] = carta_data['carta_categoria']
    if not d.get('uso') and carta_data.get('carta_uso'):
        d['uso'] = carta_data['carta_uso']
    # Nº Chasis / Nº Motor no figuran ni en la Solicitud ni en la Carta de
    # Aprobación de algunas variantes — quedan a completar a mano.
    if not d.get('chasis'):
        d['chasis'] = 'COMPLETAR CORRECTAMENTE'
    if not d.get('motor'):
        d['motor'] = 'COMPLETAR CORRECTAMENTE'

    if tipo_op.endswith('_PRE'):
        print("Parseando mutuo prendario (caso Pre-prenda)...")
        d.update(parsear_mutuo_prendario(mutuo_path))
    else:
        d.update({'mutuo_uvas': '', 'mutuo_uva_valor': '',
                  'mutuo_fecha_dd': '', 'mutuo_fecha_mm': '', 'mutuo_fecha_aa': '',
                  'mutuo_monto_prenda': '', 'mutuo_prim_venc': ''})

    print("\nDatos extraídos:")
    for k,v in d.items(): print(f"  {k:22s}: {v}")

    # Cargar config del Excel
    entradas = cargar_config_excel(xlsx_path, tipo_op)
    print(f"\nCampos activos para {tipo_op}: {len(entradas)}")

    # Agrupar por hoja
    por_hoja = {}
    for e in entradas:
        por_hoja.setdefault(e['hoja'], []).append(e)

    # Generar PDF
    reader = PdfReader(template_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        hoja_num = i + 1
        packet = io.BytesIO()
        cv = canvas.Canvas(packet, pagesize=(W, H))

        for e in por_hoja.get(hoja_num, []):
            puntos = resolver(e['variable_id'], d)
            for col, row, texto, *rest in puntos:
                if texto == '__BOX__':
                    col2, row2 = rest[0], rest[1]
                    draw_box(cv, col, row, col2, row2, fill_color=red, alpha=0.12)
                else:
                    size = rest[0] if rest else 6.5
                    draw(cv, col, row, texto, size)

        cv.showPage(); cv.save(); packet.seek(0)
        page.merge_page(PdfReader(packet).pages[0])
        writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)
    print(f"\n✓ PDF generado: {output_path}")


# ── ENTRY POINT ─────────────────────────────────────────────
if __name__ == '__main__':
    sol      = sys.argv[1] if len(sys.argv)>1 else '/mnt/user-data/uploads/2539942_7701FCF0EFC181FA78BE45BC36E79A35_Solicitud__1_.pdf'
    template = sys.argv[2] if len(sys.argv)>2 else '/mnt/user-data/uploads/EJEMPLO_CONTRATO_UVA_.pdf'
    out      = sys.argv[3] if len(sys.argv)>3 else '/home/claude/contrato_completado.pdf'
    xlsx     = sys.argv[4] if len(sys.argv)>4 else '/mnt/user-data/outputs/parametros_contrato_prenda.xlsx'
    fecha    = sys.argv[5] if len(sys.argv)>5 else None
    tipo_op  = sys.argv[6] if len(sys.argv)>6 else 'UVA_PI'
    carta    = sys.argv[7] if len(sys.argv)>7 else '/mnt/user-data/uploads/2539942_7701FCF0EFC181FA78BE45BC36E79A35_CartaAprobacion.pdf'
    mutuo    = sys.argv[8] if len(sys.argv)>8 else None
    generar_pdf(sol, template, out, xlsx, tipo_op, fecha, carta, mutuo)
