# Finanzas Personales (Google Apps Script)

App web para registrar ingresos y gastos (fijos y variables) con fechas y descripción.
Los ingresos/gastos fijos se programan (día del mes, monto, vigencia) y se registran solos.
Datos en una hoja de Google Sheets que el script crea la primera vez.

## Subir a Apps Script con clasp (usa tu cuenta personal de Google)
1. Activa la API: https://script.google.com/home/usersettings
2. `npm install -g @google/clasp`
3. `clasp logout` (si tenías otra cuenta) y luego `clasp login` con tu cuenta personal
4. En esta carpeta: `clasp push -f`
5. En el editor: ejecuta `instalarDisparadorDiario` una vez y crea la implementación (Implementar > Nueva implementación > Aplicación web).
