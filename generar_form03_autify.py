#!/usr/bin/env python3
"""
AUTIFY — Generador de Formulario "03" (Solicitud de Inscripción Contrato
Prendario, Decreto 15348/46 ratificado por Ley 12962) v1.0

Reutiliza los parsers de generar_prenda_autify_v7 (parsear_solicitud,
parsear_carta_aprobacion, parsear_mutuo_prendario) para no duplicar la
extracción de datos. El Excel "Parametros 03" (hoja del mismo libro
parametros_contrato_prenda.xlsx) es la única fuente de verdad para
coordenadas y valores.

Uso:
  python3 generar_form03_autify.py <solicitud.pdf> <template.pdf> <output.pdf> \
                                    <xlsx_path> [tipo UVA_PI|UVA_PRE|FIJA_PI|FIJA_PRE] \
                                    [carta_aprobacion.pdf] [mutuo_prendario.pdf]

Notas:
  - Por ahora solo están definidos los valores para UVA_PI / UVA_PRE.
    FIJA_PI / FIJA_PRE quedan "(por completar)" — el script avisa y no
    imprime esos campos hasta que se definan en el Excel.
  - Campos con valor "COMPLETAR" en el Excel (TIPO, MARCA MOTOR, MARCA
    CHASIS) se imprimen como "COMPLETAR" literal — el cliente los
    completa a mano contra el Título del Automotor.
  - Recuadros rojos semitransparentes (checkboxes tipo de documento y
    fecha de nacimiento/estado civil del deudor) se imprimen siempre,
    para todos los casos, indicando que deben completarse a mano.
"""

import os, sys, re, io
import openpyxl
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import red, Color

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generar_prenda_autify_v7 import (
    parsear_solicitud, parsear_carta_aprobacion, parsear_mutuo_prendario,
    fmt_num, MM, STEP, W, H, gx, gy, draw, draw_box, TIPO_COL,
)

SHEET_NAME = 'Parametros 03'

# Hoja única en form03: todas las coordenadas son de la página 1.
HOJA_FORM03 = 1


# ── MAPEO Variable (col B del Excel) -> Variable_ID interno ─
# El Variable_ID interno es una versión normalizada (sin espacios/acentos)
# de la columna "Variable" del Excel, prefijada con F03_.
def normalizar_variable(nombre):
    n = nombre.strip().upper()
    n = (n.replace('Á','A').replace('É','E').replace('Í','I')
           .replace('Ó','O').replace('Ú','U').replace('Ñ','N'))
    n = re.sub(r'[^A-Z0-9]+', '_', n).strip('_')
    return f'F03_{n}'


# ── CARGA DE CONFIGURACIÓN DESDE EXCEL ──────────────────────
def cargar_config_form03(xlsx_path, tipo_op):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la hoja '{SHEET_NAME}' en {xlsx_path}")
    ws = wb[SHEET_NAME]

    col_idx = TIPO_COL.get(tipo_op, 4)  # default columna D (UVA_PI)

    entradas = []
    for r in range(3, ws.max_row + 1):
        coords = ws.cell(row=r, column=1).value
        variable = ws.cell(row=r, column=2).value
        val = ws.cell(row=r, column=col_idx).value

        if coords is None or str(coords).strip() in ('', 'None'):
            continue
        if not variable:
            continue
        if val is None or str(val).strip() in ('', '(por completar)', 'N/A', 'None'):
            continue

        coords_str = str(coords)
        variable_id = normalizar_variable(str(variable))

        entradas.append({
            'variable_id': variable_id,
            'coords_str': coords_str,
            'valor_excel': str(val),
        })

    wb.close()
    return entradas


# ── PARSEO DE COORDENADAS ───────────────────────────────────
def parse_coords(coords_str):
    """
    Convierte un string de coordenadas en lista de tuplas.
    Soporta:
      - '22.16'              -> [(22, 16)]                  punto simple
      - '20.37-45.37'        -> [(20, 37, 45, 37)]          rango de 1 línea
                                  (ancho disponible para wrap; misma fila)
      - '20.28-45.28\n20.32-45.32\n20.34-45.34'
                              -> [(20,28,45,28),(20,32,45,32),(20,34,45,34)]
      - '50.50-80.54'        -> [(50,50,80,54)]             recuadro
                                  (filas distintas: r1 != r2)
    Tanto los rangos de 1 línea como los recuadros se devuelven como
    4-tuplas (col1,row1,col2,row2); el llamador decide cómo usarlos según
    el campo (texto con ancho disponible vs. recuadro).
    """
    out = []
    for line in coords_str.split('\n'):
        line = line.strip()
        if not line: continue
        if '-' in line:
            a, b = line.split('-', 1)
            c1, r1 = a.split('.'); c2, r2 = b.split('.')
            out.append((float(c1), float(r1), float(c2), float(r2)))
        else:
            c, r = line.split('.')
            out.append((float(c), float(r)))
    return out


# ── RESOLUCIÓN DE CADA VARIABLE ─────────────────────────────
def pt(c):
    """Extrae (col,row) de una tupla de coordenadas, sea punto (2) o rango/recuadro (4)."""
    return (c[0], c[1])


def resolver_form03(entrada, d):
    """
    Devuelve una lista de "comandos de dibujo":
      ('text', col, row, texto, size)
      ('box',  col1, row1, col2, row2)
    """
    vid   = entrada['variable_id']
    val   = entrada['valor_excel']
    coords = parse_coords(entrada['coords_str'])
    R = []

    # ── Recuadros (rangos de 4 valores en una sola línea de coords) ──
    if vid in ('F03_RECUADRO_CHECKBOXES_TIPO_DE_DOCUMENTO_DEL_DEUDOR',
               'F03_RECUADRO_FECHA_NACIMIENTO_Y_ESTADO_CIVIL_DEL_DEUDOR'):
        for c in coords:
            if len(c) == 4:
                R.append(('box', c[0], c[1], c[2], c[3]))
        return R

    # ── Casos especiales que requieren lógica con datos parseados (d) ──

    if vid == 'F03_DIA':
        for cc in coords:
            c, r = pt(cc); R.append(('text', c, r, 'XX'))
        return R
    if vid == 'F03_MES':
        for cc in coords:
            c, r = pt(cc); R.append(('text', c, r, 'XX'))
        return R
    if vid == 'F03_ANO':
        for cc in coords:
            c, r = pt(cc); R.append(('text', c, r, 'XX'))
        return R

    if vid == 'F03_MONTO_INSC_PRENDA_NUMERO_EJ_56_433_264_00':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, fmt_num(d['monto_insc']), 7))
        return R

    if vid in ('F03_UNIDAD_GARANTIA_PRENDARIA_DOMINIO_EJ_AG018VE', 'F03_DOMINIO'):
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['dominio'], 7))
        return R

    if vid == 'F03_CUIT_ACREEDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, '33-50000517-9'))
        return R

    if vid == 'F03_APELLIDO_Y_NOMBRE_ACREEDOR':
        return _wrap_3lineas(coords, 'BANCO SUPERVIELLE S.A.')

    if vid == 'F03_CALLE_ACREEDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'RECONQUISTA'))
        return R

    if vid == 'F03_NUMERO_ACREEDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, '330'))
        return R

    if vid == 'F03_CP_ACREEDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'C1003ABH'))
        return R

    if vid == 'F03_LOCALIDAD_ACREEDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'C.A.B.A.'))
        return R

    if vid == 'F03_PERSONERIA_OTORGADA':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'I.G.J'))
        return R

    if vid == 'F03_DATOS_DE_INSCRIPCION':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'N°7333 L28 TdeS por A 27-06-05', 6))
        return R

    if vid == 'F03_DIA_INSCRIPCION':
        for cc in coords:
            c, r = pt(cc); R.append(('text', c, r, '27'))
        return R
    if vid == 'F03_MES_INSCRIPCION':
        for cc in coords:
            c, r = pt(cc); R.append(('text', c, r, '06'))
        return R
    if vid == 'F03_ANO_INSCRIPCION':
        for cc in coords:
            c, r = pt(cc); R.append(('text', c, r, '05'))
        return R

    if vid == 'F03_CUIL_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['cuit']))
        return R

    if vid == 'F03_APELLIDO_Y_NOMBRE_DEUDOR':
        return _wrap_3lineas(coords, d['nombre_completo'])

    if vid == 'F03_CALLE_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['dom_calle']))
        return R

    if vid == 'F03_NUMERO_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['dom_num']))
        return R

    if vid == 'F03_PISO_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['dom_piso']))
        return R

    if vid == 'F03_DPTO_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['dom_depto']))
        return R

    if vid == 'F03_CP_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d.get('cod_postal', '')))
        return R

    if vid == 'F03_LOCALIDAD_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['localidad']))
        return R

    if vid == 'F03_PROVINCIA_DEUDOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['provincia']))
        return R

    if vid == 'F03_DNI':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, f"DNI {d['dni']}"))
        return R

    if vid in ('F03_DIA_NACIMIENTO','F03_MES_NACIMIENTO','F03_ANO_NACIMIENTO'):
        dd_n, mm_n, yyyy_n = ('', '', '')
        if d.get('fecha_nac'):
            try:
                dd_n, mm_n, yyyy_n = d['fecha_nac'].split('/')
            except ValueError:
                pass
        valores = {'F03_DIA_NACIMIENTO': dd_n,
                   'F03_MES_NACIMIENTO': mm_n,
                   'F03_ANO_NACIMIENTO': yyyy_n[-2:] if yyyy_n else ''}
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, valores[vid]))
        return R

    if vid == 'F03_FIRMA':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'X', 14))
        return R

    if vid == 'F03_MARCA':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['marca']))
        return R

    if vid == 'F03_TIPO':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'COMPLETAR'))
        return R

    if vid == 'F03_MODELO':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['modelo']))
        return R

    if vid == 'F03_MARCA_MOTOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'COMPLETAR'))
        return R

    if vid == 'F03_N_DE_MOTOR':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['motor']))
        return R

    if vid == 'F03_MARCA_CHASIS':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'COMPLETAR'))
        return R

    if vid == 'F03_N_DE_CHASIS':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, d['chasis']))
        return R

    if vid == 'F03_SOLICITUD_TIPO':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'X', 9))
        return R

    if vid == 'F03_GRADO':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, '1°'))
        return R

    if vid == 'F03_CLAUSULA_DE_ACTUALIZACION':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'X', 9))
        return R

    if vid == 'F03_CONCEPTO':
        for cc in coords:
            c, r = pt(cc)
            R.append(('text', c, r, 'X', 9))
        return R

    # Variable sin lógica definida -> avisar
    print(f"  ⚠ Variable_ID sin lógica de resolución: '{vid}'")
    return R


def _wrap_3lineas(coords, texto):
    """
    Distribuye 'texto' en hasta 3 líneas usando las coordenadas dadas.
    Cada entrada de 'coords' es un rango (col1,row1,col2,row2) que
    representa el ancho disponible de esa línea; (col1,row1) es el punto
    de inicio del texto y (col2-col1) columnas el ancho disponible
    (~1.5 caracteres por columna a tamaño 6.5). Wrap solo si el texto no
    entra en una línea — igual que H1_BIEN en el contrato de prenda.
    """
    R = []
    resto = texto
    for i, c in enumerate(coords):
        if not resto:
            break
        col1, row1 = c[0], c[1]
        ancho_cols = (c[2] - c[0]) if len(c) == 4 else 25
        max_chars = max(int(ancho_cols * 1.5), 5)
        if len(resto) <= max_chars or i == len(coords) - 1:
            R.append(('text', col1, row1, resto))
            resto = ''
        else:
            cut = resto[:max_chars].rfind(' ')
            if cut < 0: cut = max_chars
            R.append(('text', col1, row1, resto[:cut]))
            resto = resto[cut+1:]
    return R


# ── GENERADOR PRINCIPAL ─────────────────────────────────────
def generar_form03(solicitud_path, template_path, output_path, xlsx_path,
                    tipo_op='UVA_PI', carta_path=None, mutuo_path=None):

    if tipo_op not in TIPO_COL:
        print(f"⚠ Tipo '{tipo_op}' no válido. Usando UVA_PI.")
        tipo_op = 'UVA_PI'

    if tipo_op.startswith('FIJA'):
        print(f"⚠ El tipo '{tipo_op}' todavía no tiene valores definidos para el "
              f"Formulario 03 (columna pendiente '(por completar)'). "
              f"El PDF se generará sin overlay de datos.")

    print(f"Tipo de operación : {tipo_op}")
    print(f"Config Excel      : {xlsx_path} (hoja '{SHEET_NAME}')")
    print("Parseando solicitud...")
    d = parsear_solicitud(solicitud_path)
    d['tipo_op'] = tipo_op

    print("Parseando carta de aprobación...")
    carta_data = parsear_carta_aprobacion(carta_path)
    d.update(carta_data)
    if not d.get('marca') and carta_data.get('carta_marca'):
        d['marca'] = carta_data['carta_marca']
    if not d.get('modelo') and carta_data.get('carta_modelo'):
        d['modelo'] = carta_data['carta_modelo']
    if not d.get('chasis'):
        d['chasis'] = 'COMPLETAR CORRECTAMENTE'
    if not d.get('motor'):
        d['motor'] = 'COMPLETAR CORRECTAMENTE'

    if tipo_op.endswith('_PRE'):
        print("Parseando mutuo prendario (caso Pre-prenda)...")
        d.update(parsear_mutuo_prendario(mutuo_path))

    entradas = cargar_config_form03(xlsx_path, tipo_op)
    print(f"\nCampos activos para {tipo_op}: {len(entradas)}")

    reader = PdfReader(template_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        hoja_num = i + 1
        packet = io.BytesIO()
        cv = canvas.Canvas(packet, pagesize=(W, H))

        if hoja_num == HOJA_FORM03:
            for e in entradas:
                comandos = resolver_form03(e, d)
                for cmd in comandos:
                    if cmd[0] == 'text':
                        _, c, r, texto, *rest = cmd
                        size = rest[0] if rest else 6.5
                        draw(cv, c, r, texto, size)
                    elif cmd[0] == 'box':
                        _, c1, r1, c2, r2 = cmd
                        draw_box(cv, c1, r1, c2, r2, fill_color=red, alpha=0.12)

        cv.showPage(); cv.save(); packet.seek(0)
        page.merge_page(PdfReader(packet).pages[0])
        writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)
    print(f"\n✓ PDF generado: {output_path}")


# ── ENTRY POINT ─────────────────────────────────────────────
if __name__ == '__main__':
    sol      = sys.argv[1] if len(sys.argv) > 1 else None
    template = sys.argv[2] if len(sys.argv) > 2 else '/mnt/user-data/uploads/form03_grilla.pdf'
    out      = sys.argv[3] if len(sys.argv) > 3 else '/home/claude/form03_completado.pdf'
    xlsx     = sys.argv[4] if len(sys.argv) > 4 else '/mnt/user-data/outputs/parametros_contrato_prenda.xlsx'
    tipo_op  = sys.argv[5] if len(sys.argv) > 5 else 'UVA_PI'
    carta    = sys.argv[6] if len(sys.argv) > 6 else None
    mutuo    = sys.argv[7] if len(sys.argv) > 7 else None

    if not sol:
        print("Uso: generar_form03_autify.py <solicitud.pdf> <template.pdf> <output.pdf> "
              "<xlsx> [tipo_op] [carta.pdf] [mutuo.pdf]")
        sys.exit(1)

    generar_form03(sol, template, out, xlsx, tipo_op, carta, mutuo)
