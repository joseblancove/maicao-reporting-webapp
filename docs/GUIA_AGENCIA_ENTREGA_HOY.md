# Guia rapida para presentar hoy a la agencia

## Mensaje ejecutivo

El reporte fue transformado de una PPT estatica a un sistema automatizable:

```text
Google Sheet colaborativo
        ↓
Web app generadora
        ↓
PowerPoint editable final
```

Esto reduce trabajo manual, mejora consistencia visual, minimiza errores y permite escalar el mismo modelo a otros clientes.

---

## Que debe hacer el equipo de reporting cada mes

1. Duplicar el Google Sheet del mes anterior o actualizar el maestro.
2. Cambiar `00_Control > Mes actual`.
3. Completar KPIs en `13_Platform_KPIs`.
4. Actualizar audiencia en `02_Audience`.
5. Actualizar inversion en `03_Paid_Media`.
6. Actualizar insights en `07_Insights` y recomendaciones en `08_Recomendaciones`.
7. Entrar a la web app.
8. Generar el PPT.
9. Revisar alertas y descargar version final.

---

## Beneficios inmediatos

- No mas copy/paste manual de KPIs.
- Validacion antes de generar.
- Misma plantilla visual todos los meses.
- PowerPoint editable como entregable final.
- Google Sheet como fuente unica de verdad.
- Escalable a otros reportes mensuales.

---

## Decision pendiente

Para que quede 100% online se requiere una de estas opciones:

### Opcion recomendada
Web app Python alojada en Streamlit Cloud, Cloud Run, Render o servidor interno.

### Opcion de contingencia
Usar la web app local con `RUN_WEBAPP.command` en el computador del responsable de reporting.

---

## Roles sugeridos

- Data owner: actualiza Google Sheet.
- Reporting owner: revisa validaciones y genera PPT.
- Visual owner: revisa detalles finales de layout.
- Cliente/agencia: recibe PPT editable.
