# Función constructora para las tarjetas de KPI ejecutivas (Corregida sin sangrías de Markdown)
    def render_kpi_card(titulo, val_act, val_ant, val_meta, badge_txt, badge_class, fill_class, es_moneda=False):
        if es_moneda:
            txt_actual = f"${val_act:,.2f}"
            txt_meta = f"${val_meta:,.2f}"
        else:
            txt_actual = f"{val_act:,}"
            txt_meta = f"{val_meta:,}"

        # Variación porcentual vs periodo anterior
        delta = ((val_act - val_ant) / val_ant * 100) if val_ant > 0 else 0.0
        delta_class = "delta-pos" if delta >= 0 else "delta-neg"
        delta_sign = "+" if delta >= 0 else ""
        delta_icon = "▲" if delta >= 0 else "▼"

        # % de Cumplimiento de la meta
        cumplimiento = (val_act / val_meta * 100) if val_meta > 0 else 0.0
        ancho_barra = min(100.0, max(0.0, cumplimiento))

        # Se retorna como cadena HTML continua para que Markdown no lo tome como bloque de código
        return (
            f'<div class="kpi-card">'
            f'<div class="kpi-top">'
            f'<span class="kpi-label">{titulo}</span>'
            f'<span class="kpi-badge {badge_class}">{badge_txt}</span>'
            f'</div>'
            f'<div class="kpi-value">{txt_actual}</div>'
            f'<div class="kpi-progress-bg">'
            f'<div class="kpi-progress-fill {fill_class}" style="width: {ancho_barra:.1f}%;"></div>'
            f'</div>'
            f'<div class="kpi-bottom">'
            f'<span class="kpi-meta-text">🎯 Meta: <b>{txt_meta}</b> ({cumplimiento:.0f}%)</span>'
            f'<span class="kpi-delta-pill {delta_class}">{delta_icon} {delta_sign}{delta:.1f}% vs ant.</span>'
            f'</div>'
            f'</div>'
        )
