from app import *
print("AJUSTES_ASIGNACION:", st.session_state.get("ajustes_asignacion", {}))
asignacion_base = asignacion.copy()
asignacion_final = aplicar_ajustes_asignacion(asignacion_base)
print("Count of TEMPLO_ASIGNADO_FINAL:")
print(asignacion_final['TEMPLO_ASIGNADO_FINAL'].value_counts())
