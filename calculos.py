"""
calculos.py Es un módulo de cálculo de liquidaciones de sueldo.
Aplica la lógica de remuneraciones chilena simplificada.
"""
#----------------TASAS Y CONSTANTES LEGALES
TASA_AFP          = 0.10      # 10 %---(tasa promedio; varía por AFP)
TASA_SALUD        = 0.07      # 7 %
TASA_CESANTIA_TRA = 0.006     # 0.6 %---(trabajador)
VALOR_HORA_EXTRA  = 5_000     # $5.000 por hora extra---(según empresa)

# Haberes no imponibles 
MONTO_COLACION     = 47_318
MONTO_MOVILIZACION = 30_500
MONTO_ASISTENCIA   = 25_000
MONTO_CARGA_FAMI   = 13_523    # Tramo A (aprox 2024)


def calcular_liquidacion(datos: dict) -> dict:
    """
    Recibe el diccionario de datos del formulario y retorna
    un diccionario con todos los montos calculados.
    """
    sueldo_base:        int  = datos["sueldo_base"]
    dias:               int  = datos["dias_trabajados"]
    horas_extra:        int  = datos["horas_extra"]
    bono_colacion:      bool = datos.get("bono_colacion", False)
    bono_movilizacion:  bool = datos.get("bono_movilizacion", False)
    bono_asistencia:    bool = datos.get("bono_asistencia", False)
    asig_familiar:      bool = datos.get("asig_familiar", False)
    cargas:             int  = datos.get("cargas_familiares", 0)
    desc_prestamo:      int  = datos.get("desc_prestamo", 0)
    desc_adelanto:      int  = datos.get("desc_adelanto", 0)
    desc_otros:         int  = datos.get("desc_otros", 0)

    # Sueldo base proporcional a días trabajados ──
    sueldo_proporcional = round(sueldo_base * dias / 30)

    # Haberes imponibles ──
    valor_horas_extra = horas_extra * VALOR_HORA_EXTRA
    # Las horas extra sí son imponibles para efectos de este cálculo
    total_imponible = sueldo_proporcional + valor_horas_extra

    # Haberes no imponibles ──
    val_colacion     = MONTO_COLACION     if bono_colacion     else 0
    val_movilizacion = MONTO_MOVILIZACION if bono_movilizacion else 0
    val_asistencia   = MONTO_ASISTENCIA   if bono_asistencia   else 0
    val_asig_fam     = MONTO_CARGA_FAMI * int(cargas) if asig_familiar else 0

    # Total haberes ──
    total_haberes = (
        total_imponible
        + val_colacion
        + val_movilizacion
        + val_asistencia
        + val_asig_fam
    )

    # Descuentos legales (se aplican sobre base imponible) ──
    desc_afp      = round(total_imponible * TASA_AFP)
    desc_salud    = round(total_imponible * TASA_SALUD)
    desc_cesantia = round(total_imponible * TASA_CESANTIA_TRA)
    total_desc_legales = desc_afp + desc_salud + desc_cesantia

    # Descuentos adicionales ──
    total_desc_adicionales = desc_prestamo + desc_adelanto + desc_otros

    # Total descuentos y líquido ──
    total_descuentos = total_desc_legales + total_desc_adicionales
    liquido_a_pagar  = total_haberes - total_descuentos
    if liquido_a_pagar < 0:
        liquido_a_pagar = 0

    return {
        # Haberes
        "sueldo_base_proporcional": sueldo_proporcional,
        "valor_horas_extra":        valor_horas_extra,
        "bono_colacion":            val_colacion,
        "bono_movilizacion":        val_movilizacion,
        "bono_asistencia":          val_asistencia,
        "asig_familiar":            val_asig_fam,
        "total_imponible":          total_imponible,
        "total_haberes":            total_haberes,
        # Descuentos
        "desc_afp":                 desc_afp,
        "desc_salud":               desc_salud,
        "desc_cesantia":            desc_cesantia,
        "desc_prestamo":            desc_prestamo,
        "desc_adelanto":            desc_adelanto,
        "desc_otros":               desc_otros,
        "total_desc_legales":       total_desc_legales,
        "total_desc_adicionales":   total_desc_adicionales,
        "total_descuentos":         total_descuentos,
        # Resultado
        "liquido_a_pagar":          liquido_a_pagar,
        # Meta
        "tasa_afp":     TASA_AFP,
        "tasa_salud":   TASA_SALUD,
        "tasa_cesantia": TASA_CESANTIA_TRA,
    }
