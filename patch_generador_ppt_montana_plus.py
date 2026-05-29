# -*- coding: utf-8 -*-
"""
Parche SOW - Modo Montaña Plus para generar_ppt_meteo_final_v3.py

Uso en la carpeta del servidor PPT:
  python patch_generador_ppt_montana_plus.py

O indicando ruta:
  python patch_generador_ppt_montana_plus.py C:\\Users\\Geomatica1\\Documents\\ppt_meteo\\generar_ppt_meteo_final_v3.py

Qué hace:
- Crea backup del generador actual.
- Reemplaza SOLO la función slide_montana(prs, data).
- La PPT reflejará montana_resumen enviado por la app, y si no existe, lo calcula desde forecast72h.
"""

from pathlib import Path
import re
import sys
from datetime import datetime

DEFAULT_FILE = "generar_ppt_meteo_final_v3.py"

HELPERS = r'''

# ============================================================
# MONTAÑA PLUS - HELPERS SOW
# ============================================================
def _montana_estado(block):
    m = block.get("montana") or {}
    return str(m.get("estado") or block.get("estado_montana") or "SIN DATOS")


def _montana_motivo(block):
    m = block.get("montana") or {}
    return str(m.get("motivo") or "Sin datos")


def _montana_score(block):
    estado = normalize(_montana_estado(block))
    score = 0
    if "RESTRINGIDO" in estado:
        score += 100
    elif "MARGINAL" in estado:
        score += 50
    viento = valnum(block.get("viento"), 0) or 0
    rachas = valnum(block.get("rachas"), 0) or 0
    vis = valnum(block.get("visibilidad"), 99) or 99
    nieve = valnum(block.get("nieve_cm_h"), 0) or 0
    acum = valnum(block.get("nieve_acumulada_cm"), 0) or 0
    iso = valnum(block.get("isoterma_0_m"), None)
    techo = valnum(block.get("techo") or block.get("base_nubes_ft"), None)
    temp = valnum(block.get("temperatura"), 15) or 15
    if viento >= 35 or rachas >= 45:
        score += 35
    elif viento >= 20 or rachas >= 30:
        score += 15
    if vis < 2:
        score += 35
    elif vis < 5:
        score += 15
    if nieve >= 2:
        score += 30
    elif nieve >= 0.1:
        score += 10
    if acum >= 15:
        score += 25
    if iso is not None and iso > 0 and iso < 2200:
        score += 20
    if techo is not None and techo > 0 and techo < 1000:
        score += 20
    if temp <= -8:
        score += 20
    return score


def _montana_best_window(forecast):
    if not forecast:
        return "Sin datos"
    best = []
    cur = []
    for b in forecast:
        estado = normalize(_montana_estado(b))
        if "FAVORABLE" in estado:
            cur.append(b)
        else:
            if len(cur) > len(best):
                best = list(cur)
            cur = []
    if len(cur) > len(best):
        best = list(cur)
    if not best:
        return "Sin ventana favorable"
    return f"{format_dt(best[0].get('hora'), 'daylabel')} {format_dt(best[0].get('hora'))}-{format_dt(best[-1].get('hora'))}h"


def _montana_riesgos_from_forecast(forecast):
    if not forecast:
        return []
    def nums(key):
        return [valnum(b.get(key), None) for b in forecast if valnum(b.get(key), None) is not None]
    viento = max(nums("viento") or [0])
    rachas = max(nums("rachas") or [0])
    vis = min(nums("visibilidad") or [0])
    nieve = max(nums("nieve_cm_h") or [0])
    acum = max(nums("nieve_acumulada_cm") or [0])
    iso_vals = [v for v in nums("isoterma_0_m") if v > 0]
    techo_vals = [v for v in (nums("techo") + nums("base_nubes_ft")) if v > 0]
    temp = min(nums("temperatura") or [0])

    def ev_wind():
        if viento >= 35 or rachas >= 45: return "RESTRINGIDO"
        if viento >= 20 or rachas >= 30: return "MARGINAL"
        return "FAVORABLE"
    def ev_vis():
        if vis <= 0: return "SIN DATO"
        if vis < 2: return "RESTRINGIDO"
        if vis < 5: return "MARGINAL"
        return "FAVORABLE"
    def ev_snow():
        if nieve >= 2: return "RESTRINGIDO"
        if nieve >= 0.1: return "MARGINAL"
        return "FAVORABLE"
    def ev_acum():
        if acum >= 15: return "RESTRINGIDO"
        if acum >= 5: return "MARGINAL"
        return "FAVORABLE"
    def ev_iso():
        if not iso_vals: return "SIN DATO"
        m = min(iso_vals)
        if m < 2200: return "RESTRINGIDO"
        if m <= 3000: return "MARGINAL"
        return "FAVORABLE"
    def ev_techo():
        if not techo_vals: return "SIN DATO"
        m = min(techo_vals)
        if m < 1000: return "RESTRINGIDO"
        if m < 2000: return "MARGINAL"
        return "FAVORABLE"
    def ev_temp():
        if temp <= -8: return "RESTRINGIDO"
        if temp <= 0: return "MARGINAL"
        return "FAVORABLE"

    return [
        {"riesgo": "Viento/rachas", "estado": ev_wind(), "detalle": f"{viento:.0f}/{rachas:.0f} kt"},
        {"riesgo": "Visibilidad", "estado": ev_vis(), "detalle": f"{vis:.1f} km mín" if vis else "N/D"},
        {"riesgo": "Nevada", "estado": ev_snow(), "detalle": f"{nieve:.1f} cm/h máx"},
        {"riesgo": "Nieve acumulada", "estado": ev_acum(), "detalle": f"{acum:.1f} cm máx"},
        {"riesgo": "Isoterma 0", "estado": ev_iso(), "detalle": f"{min(iso_vals):.0f} m mín" if iso_vals else "N/D"},
        {"riesgo": "Techo/base nubes", "estado": ev_techo(), "detalle": f"{min(techo_vals):.0f} ft mín" if techo_vals else "N/D"},
        {"riesgo": "Frío/exposición", "estado": ev_temp(), "detalle": f"{temp:.1f}°C mín"},
    ]


def _montana_dias_from_forecast(forecast):
    dias = grouped_by_date(forecast)
    out = []
    for fecha, blocks in dias.items():
        worst = max(blocks, key=_montana_score) if blocks else None
        estados = [normalize(_montana_estado(b)) for b in blocks]
        estado = "RESTRINGIDO" if any("RESTRINGIDO" in e for e in estados) else "MARGINAL" if any("MARGINAL" in e for e in estados) else "FAVORABLE"
        out.append({
            "fecha": fecha,
            "estado": estado,
            "mejor_ventana": _montana_best_window(blocks),
            "peor_bloque": f"{format_dt(worst.get('hora'))}h" if worst else "-",
            "motivo": _montana_motivo(worst) if worst else "Sin datos",
        })
    return out


def _montana_resumen_auto(data):
    forecast = data.get("forecast72h", []) or []
    montana = data.get("montana") or {}
    if not forecast:
        return {
            "estado": montana.get("estado", "SIN DATOS"),
            "motivo": montana.get("motivo", "Sin datos"),
            "mejor_ventana": "Sin datos",
            "peor_bloque": "Sin datos",
            "peor_estado": "SIN DATOS",
            "riesgos": [],
            "dias": [],
            "conclusion": "Sin datos suficientes para evaluar montaña.",
            "fuente": "Open-Meteo / SOW",
        }
    worst = max(forecast, key=_montana_score)
    estados = [normalize(_montana_estado(b)) for b in forecast]
    rest = sum(1 for e in estados if "RESTRINGIDO" in e)
    marg = sum(1 for e in estados if "MARGINAL" in e)
    estado = "RESTRINGIDO" if rest else "MARGINAL" if marg else "FAVORABLE"
    motivo = _montana_motivo(worst)
    if estado == "RESTRINGIDO":
        conclusion = f"Operación de montaña NO recomendable en los bloques restringidos. Motivo principal: {motivo}. Reforzar navegación, abrigo, visibilidad, ruta de evacuación y actualización meteorológica."
    elif estado == "MARGINAL":
        conclusion = f"Operación de montaña posible solo con control permanente. Priorizar mejor ventana y reevaluar viento, visibilidad, nieve e isoterma 0. Motivo principal: {motivo}."
    else:
        conclusion = "Condiciones generales favorables para montaña. Mantener monitoreo por cambios rápidos en altura."
    return {
        "estado": estado,
        "motivo": motivo,
        "bloques_restringidos": rest,
        "bloques_marginales": marg,
        "mejor_ventana": _montana_best_window(forecast),
        "peor_bloque": f"{format_dt(worst.get('hora'), 'daylabel')} {format_dt(worst.get('hora'))}h",
        "peor_estado": _montana_estado(worst),
        "peor_motivo": motivo,
        "riesgos": _montana_riesgos_from_forecast(forecast),
        "dias": _montana_dias_from_forecast(forecast),
        "conclusion": conclusion,
        "fuente": "Open-Meteo / SOW",
    }
'''

SLIDE_FUNC = r'''
def slide_montana(prs, data):
    if not data.get("incluir_montana", True):
        return

    forecast = data.get("forecast72h", []) or []
    resumen = data.get("montana_resumen") or _montana_resumen_auto(data)
    montana = data.get("montana") or {}

    if not resumen and not montana and not forecast:
        return

    estado = str(resumen.get("estado") or montana.get("estado") or "SIN EVALUAR")
    motivo = str(resumen.get("motivo") or montana.get("motivo") or "Sin datos")
    riesgos = resumen.get("riesgos") or _montana_riesgos_from_forecast(forecast)
    dias = resumen.get("dias") or _montana_dias_from_forecast(forecast)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    add_title(slide, data, "MODO MONTAÑA Y NIEVE", "Evaluación operacional de isoterma 0, nieve, visibilidad, viento, techo y exposición.")

    # Panel principal
    add_rect(slide, 0.45, 1.55, 3.25, 1.45, PANEL2)
    add_text(slide, "ESTADO MONTAÑA", 0.62, 1.72, 2.7, 0.22, 9.5, True, MUTED)
    add_badge(slide, estado, 0.62, 2.03, 2.65, 0.48, status_color(estado))
    add_text(slide, text_wrap(motivo, 150), 0.62, 2.58, 2.75, 0.32, 7.8, False, WHITE)

    add_rect(slide, 3.95, 1.55, 3.05, 1.45, PANEL)
    add_text(slide, "MEJOR VENTANA", 4.12, 1.72, 2.6, 0.22, 9.5, True, GREEN)
    add_text(slide, str(resumen.get("mejor_ventana", "Sin ventana")), 4.12, 2.05, 2.5, 0.35, 15, True, WHITE)
    add_text(slide, "Priorizar ejecución en este bloque si el resto de factores se mantiene.", 4.12, 2.52, 2.55, 0.33, 7.3, False, MUTED)

    add_rect(slide, 7.22, 1.55, 2.75, 1.45, PANEL)
    add_text(slide, "PEOR BLOQUE", 7.38, 1.72, 2.35, 0.22, 9.5, True, RED)
    add_text(slide, str(resumen.get("peor_bloque", "Sin bloque crítico")), 7.38, 2.05, 2.25, 0.35, 13, True, WHITE)
    add_text(slide, text_wrap(str(resumen.get("peor_motivo", "Sin datos")), 90), 7.38, 2.52, 2.25, 0.33, 7.2, False, MUTED)

    add_rect(slide, 10.18, 1.55, 2.7, 1.45, PANEL)
    add_text(slide, "FUENTE", 10.35, 1.72, 2.2, 0.22, 9.5, True, CYAN)
    add_text(slide, str(resumen.get("fuente", "Open-Meteo / SOW")), 10.35, 2.05, 2.2, 0.3, 12, True, WHITE)
    add_text(slide, "Snow/Mountain-Forecast: referencia manual hasta contar con API autorizada.", 10.35, 2.43, 2.2, 0.42, 6.8, False, MUTED)

    # Datos críticos
    add_rect(slide, 0.45, 3.25, 5.9, 1.45, PANEL)
    add_text(slide, "DATOS CRÍTICOS", 0.62, 3.42, 5.3, 0.22, 10, True, CYAN)
    datos = [
        ("Temp", resumen.get("temp_rango", "-")),
        ("Viento/Rachas", resumen.get("viento_rachas", "-")),
        ("Vis mín", resumen.get("visibilidad_min", "-")),
        ("Nieve máx", resumen.get("nieve_max", "-")),
        ("Nieve acum", resumen.get("nieve_acum_max", "-")),
        ("Isoterma", resumen.get("isoterma_rango", "-")),
        ("Techo mín", resumen.get("techo_min", "-")),
    ]
    for i, (k, v) in enumerate(datos[:7]):
        x = 0.62 + (i % 4) * 1.42
        y = 3.78 + (i // 4) * 0.48
        add_text(slide, str(k).upper(), x, y, 1.25, 0.16, 6.6, True, MUTED)
        add_text(slide, str(v), x, y + 0.16, 1.30, 0.22, 8.2, True, WHITE)

    # Riesgos
    add_rect(slide, 6.6, 3.25, 6.28, 1.45, PANEL)
    add_text(slide, "RIESGOS OPERACIONALES", 6.78, 3.42, 5.6, 0.22, 10, True, CYAN)
    for i, r in enumerate((riesgos or [])[:4]):
        x = 6.78 + (i % 2) * 3.0
        y = 3.78 + (i // 2) * 0.48
        est = str(r.get("estado", "SIN DATO"))
        add_text(slide, str(r.get("riesgo", "-")), x, y, 1.6, 0.18, 7.2, True, WHITE)
        add_badge(slide, estado_abbr(est), x + 1.62, y - 0.01, 0.68, 0.22, status_color(est))
        add_text(slide, str(r.get("detalle", "-")), x + 2.35, y, 0.55, 0.18, 6.5, False, MUTED)

    # Resumen por día
    add_rect(slide, 0.45, 4.95, 12.43, 1.25, PANEL2)
    add_text(slide, "SÍNTESIS POR DÍA", 0.62, 5.10, 3.0, 0.22, 10, True, CYAN)
    x = 0.62
    for d in (dias or [])[:4]:
        est = str(d.get("estado", "SIN DATOS"))
        add_rect(slide, x, 5.42, 2.95, 0.55, PANEL)
        add_text(slide, str(d.get("fecha", "-")), x + 0.08, 5.48, 0.72, 0.16, 6.6, True, MUTED)
        add_badge(slide, estado_abbr(est), x + 0.78, 5.47, 0.70, 0.20, status_color(est))
        add_text(slide, f"Mejor: {d.get('mejor_ventana', '-')}", x + 1.58, 5.47, 1.25, 0.16, 6.4, False, WHITE)
        add_text(slide, text_wrap(f"Peor {d.get('peor_bloque', '-')}: {d.get('motivo', '-')}", 70), x + 0.08, 5.70, 2.72, 0.16, 6.2, False, MUTED)
        x += 3.05

    add_rect(slide, 0.45, 6.35, 12.43, 0.62, DARK)
    add_text(slide, "CONCLUSIÓN", 0.62, 6.43, 1.6, 0.18, 8.8, True, CYAN)
    add_text(slide, text_wrap(str(resumen.get("conclusion", "Mantener monitoreo meteorológico en montaña.")), 270), 1.75, 6.40, 10.75, 0.33, 8.1, False, WHITE)
'''


def main():
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(DEFAULT_FILE)
    if not target.exists():
        print(f"ERROR: no encontré {target}. Ejecuta este script en la carpeta donde está generar_ppt_meteo_final_v3.py o pasa la ruta completa.")
        sys.exit(1)

    text = target.read_text(encoding="utf-8")
    backup = target.with_suffix(target.suffix + f".bak_montana_plus_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    backup.write_text(text, encoding="utf-8")

    if "# MONTAÑA PLUS - HELPERS SOW" not in text:
        idx = text.find("def slide_montana")
        if idx < 0:
            print("ERROR: no encontré def slide_montana(prs, data) en el generador.")
            sys.exit(1)
        text = text[:idx] + HELPERS + "\n" + text[idx:]

    pattern = re.compile(r"def slide_montana\(prs, data\):.*?(?=\ndef slide_municion\(|\ndef slide_conclusiones\(|\n# ============================================================\n# MAIN)", re.S)
    if not pattern.search(text):
        print("ERROR: no pude aislar la función slide_montana. No modifiqué el archivo principal.")
        print(f"Backup creado en: {backup}")
        sys.exit(1)

    text = pattern.sub(SLIDE_FUNC.strip() + "\n\n", text, count=1)
    target.write_text(text, encoding="utf-8")
    print("OK: generador PPT actualizado con Modo Montaña Plus.")
    print(f"Backup: {backup}")
    print(f"Archivo modificado: {target}")


if __name__ == "__main__":
    main()
