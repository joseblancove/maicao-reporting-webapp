# Checklist para publicar online

## 1. Preparar repositorio

- [ ] Crear repo privado en GitHub.
- [ ] Subir archivos de `maicao_automation_webapp_v05`.
- [ ] Confirmar que `app.py` esta en la raiz.
- [ ] Confirmar que `requirements.txt` esta en la raiz.

## 2. Preparar Google Cloud

- [ ] Crear proyecto en Google Cloud.
- [ ] Activar Google Sheets API.
- [ ] Activar Google Drive API.
- [ ] Crear service account.
- [ ] Descargar JSON.
- [ ] Copiar `client_email`.

## 3. Preparar Google Sheet

- [ ] Subir Excel modelo a Google Drive.
- [ ] Abrir como Google Sheets.
- [ ] Confirmar hojas clave.
- [ ] Compartir Sheet con `client_email` del service account.

## 4. Publicar app

- [ ] Crear app en Streamlit Cloud o servidor interno.
- [ ] Apuntar a `app.py`.
- [ ] Agregar secrets del service account.
- [ ] Probar con Google Sheet.
- [ ] Descargar PPT generado.

## 5. QA final

- [ ] Verificar que aparece el mes correcto.
- [ ] Verificar overview.
- [ ] Verificar Instagram, Facebook y TikTok.
- [ ] Verificar Squad y MMPP.
- [ ] Revisar `validation_report`.
- [ ] Descargar PPT y abrirlo en PowerPoint.
