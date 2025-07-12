# 🔍 Comparador de Reservas: Snowflake vs Drive

Este proyecto en Python genera una comparación entre:

- Una **consulta de datos desde Snowflake**
- Un archivo cargado desde **Google Drive** (o local)

El objetivo es verificar que por cada **número de reserva (Booking ID)**, ambas fuentes tengan los mismos valores, y si no es así, mostrar las diferencias en valor para los servicios asociados.

---

## ⚙️ Funcionalidades principales

- Conexión a Snowflake y extracción de datos mediante SQL.
- Lectura de archivo externo en formato `.csv` o `.xlsx` desde Drive o local.
- Comparación fila por fila usando el campo `Booking ID`.
- Identificación de diferencias numéricas por servicio.
- Generación de un reporte con todas las discrepancias detectadas.
