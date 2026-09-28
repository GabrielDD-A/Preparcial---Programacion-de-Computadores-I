# Preparcial---Programacion-de-Computadores-I
# Caja de Recaudo con Tope de Seguridad

**Integrantes del grupo:**
- Gabriel David Diaz Argote
- Saray Valentina Silva Galinda
- Paola Valentina Sanchez Rozo

**Asignatura:** Programación 1
**Docente:** Ing. Carlos Carrascal

---

## 1. Descripción del proyecto

PagaYa S.A.S. recauda pagos de facturas de servicios para terceros. Por política de seguridad, cada cajero tiene un límite (tope) de efectivo que puede acumular en su caja durante el turno. Al alcanzarlo o superarlo, la caja debe suspender el servicio para que el efectivo sea retirado.

Este proyecto implementa, en Python sobre Google Colab, un módulo de control de caja que:

- Inicia un turno con un cajero y un tope de recaudo asignado.
- Atiende clientes en orden de llegada, validando cada monto recibido.
- Rechaza montos inválidos (menores o iguales a cero, o no numéricos) sin contarlos ni acumularlos.
- Suspende la caja automáticamente cuando el recaudo acumulado alcanza o supera el tope.
- Permite cerrar el turno manualmente cuando se agota la cola (digitando `0` o `FIN`).
- Genera un reporte final del turno (cajero, tope, transacciones, recaudo total, promedio por transacción y motivo de cierre).
- Guarda automáticamente todo el histórico de turnos y transacciones en Excel y CSV cada vez que ocurre un evento relevante, sin necesidad de exportar manualmente.

## 2. Requerimientos funcionales

| Código | Requerimiento |
|--------|---------------|
| RF1 | **Inicio de turno.** Solicita el cajero y el tope máximo de recaudo (debe ser mayor que cero). |
| RF2 | **Atención de clientes.** Se atiende en orden de llegada; por cada uno se solicita el monto recibido, se acumula el total recaudado y se cuenta la transacción. |
| RF3 | **Validación de montos.** Se rechazan montos ≤ 0 y valores no numéricos. Un monto rechazado no se cuenta ni acumula, y se vuelve a solicitar. |
| RF4 | **Suspensión por seguridad.** Cuando el recaudo acumulado ≥ tope, se muestra `CAJA SUSPENDIDA: se alcanzo el tope de recaudo. Dirijase a tesoreria.` y se dejan de aceptar transacciones. |
| RF5 | **Cierre por fin de cola.** Si la cola se agota antes de alcanzar el tope (el cajero digita `0` o `FIN`), el turno termina normalmente y se reporta. |
| RF6 | **Reporte final.** Cajero y tope asignado; número de transacciones; recaudo total; promedio por transacción (0 si no hubo transacciones); motivo de cierre (*tope alcanzado* o *fin de cola*). |

## 3. Reglas de negocio

| Tema | Regla |
|------|-------|
| Transacción que supera el tope | Se registra completa y luego se suspende la caja. Por eso el recaudo total puede ser mayor que el tope. |
| Condición de suspensión | `recaudo_total >= tope` |
| Moneda | Pesos colombianos (COP), valores enteros |
| Formato del reporte | Miles con separador y promedio con dos decimales |

## 4. Herramientas utilizadas

- **Python 3** en **Google Colab**
- [`pandas`](https://pandas.pydata.org/) — manejo de tablas y exportación a CSV/Excel
- [`ipywidgets`](https://ipywidgets.readthedocs.io/) — interfaz interactiva (widgets)
- [`openpyxl`](https://openpyxl.readthedocs.io/) — formato visual del Excel (encabezado, bordes, tabla, filtro)
- `google.colab.files` — descarga de los archivos generados al computador

## 5. Cómo ejecutarlo

1. Abrir [Google Colab](https://colab.research.google.com/) y crear un nuevo notebook.
2. Copiar el contenido de `caja_recaudo_colab.py` en una celda.
3. Ejecutar la celda: aparecerán los widgets de **Inicio de turno**.
4. Registrar el **cajero** y el **tope** y presionar **Iniciar turno**.
5. Ingresar montos en **Monto** y presionar **Registrar transaccion** por cada cliente.
   - Para cerrar el turno manualmente, escribir `0` o `FIN` y presionar **Registrar transaccion**, o usar el botón **Finalizar turno (FIN)**.
6. Usar **Exportar CSV** o **Exportar Excel** para descargar los archivos al computador en cualquier momento (aunque ya se van guardando automáticamente con cada acción).

## 6. Archivos generados

| Archivo | Contenido | Formato |
|---------|-----------|---------|
| `caja_recaudo.xlsx` | Hojas **Reportes** (un turno cerrado por fila) y **Detalle** (cada transacción individual) | Excel con encabezado azul, bordes, tabla con filtro/bandas, ancho automático y primera fila congelada |
| `reportes_turnos.csv` | Histórico de reportes de todos los turnos cerrados | CSV separado por `;`, codificación `utf-8-sig` |
| `detalle_transacciones.csv` | Histórico de todas las transacciones registradas | CSV separado por `;`, codificación `utf-8-sig` |

Estos archivos se acumulan turno tras turno: nunca se sobrescriben ni se vacían, sino que se regeneran con todo el histórico cada vez que se registra una transacción o se cierra un turno.

## 7. Ejemplo de ejecución

**Entrada:** cajero `C-102`, tope `$1.000.000`

| Cliente | Monto | Acumulado |
|---------|-------|-----------|
| 1 | 350.000 | 350.000 |
| 2 | -20.000 | *rechazado* |
| 3 | 400.000 | 750.000 |
| 4 | 300.000 | 1.050.000 (se supera el tope) |

**Salida esperada:**
