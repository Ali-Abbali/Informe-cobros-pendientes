"""Genera un CSV de facturas 100% ficticias para probar el informe.

Todos los nombres, importes y fechas son inventados. No hay ningún dato real
de ninguna empresa. Con la misma semilla siempre sale el mismo archivo.

Uso:
    python src/generar_datos.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

# Nombres inventados. Cualquier parecido con empresas reales es casualidad.
CLIENTES = [
    "Ferretería Demo Norte S.L.",
    "Taller Ficticio Sur S.L.",
    "Panadería Ejemplo S.A.",
    "Clínica Modelo S.L.",
    "Distribuciones Prueba S.L.",
    "Asesoría Muestra S.L.",
    "Transportes Ficticios S.A.",
    "Librería Ejemplo S.L.",
    "Construcciones Demo S.L.",
    "Informática Modelo S.L.",
    "Hostelería Prueba S.L.",
    "Agrícola Muestra S.A.",
]

TIPO_IVA = 0.21
FECHA_HOY = date(2026, 10, 22)  # fecha fija para que el ejemplo sea reproducible


def generar(num_facturas: int = 60, semilla: int = 42) -> list[dict]:
    """Devuelve una lista de facturas inventadas (pagadas, parciales y pendientes)."""
    random.seed(semilla)
    facturas = []

    for numero in range(1, num_facturas + 1):
        fecha_emision = FECHA_HOY - timedelta(days=random.randint(5, 200))
        plazo_dias = random.choice([30, 30, 60, 90])
        fecha_vencimiento = fecha_emision + timedelta(days=plazo_dias)

        base = round(random.uniform(150, 6500), 2)
        iva = round(base * TIPO_IVA, 2)
        total = round(base + iva, 2)

        importe_cobrado = 0.0
        fecha_cobro = ""
        suerte = random.random()
        dias_hasta_cobro = random.randint(5, plazo_dias + 20)
        cobro = fecha_emision + timedelta(days=dias_hasta_cobro)

        if cobro <= FECHA_HOY:
            if suerte < 0.55:  # factura pagada entera
                importe_cobrado = total
                fecha_cobro = cobro.isoformat()
            elif suerte < 0.65:  # cobro parcial
                importe_cobrado = round(total * random.choice([0.3, 0.5, 0.7]), 2)
                fecha_cobro = cobro.isoformat()

        facturas.append(
            {
                "num_factura": f"F-2026-{numero:04d}",
                "cliente": random.choice(CLIENTES),
                "fecha_emision": fecha_emision.isoformat(),
                "fecha_vencimiento": fecha_vencimiento.isoformat(),
                "base_imponible": base,
                "iva": iva,
                "total": total,
                "importe_cobrado": importe_cobrado,
                "fecha_cobro": fecha_cobro,
            }
        )
    return facturas


def guardar_csv(facturas: list[dict], ruta: Path) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("w", newline="", encoding="utf-8") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=list(facturas[0].keys()))
        escritor.writeheader()
        escritor.writerows(facturas)


if __name__ == "__main__":
    carpeta_proyecto = Path(__file__).resolve().parent.parent
    destino = carpeta_proyecto / "datos" / "facturas_ejemplo.csv"
    guardar_csv(generar(), destino)
    print(f"Datos ficticios guardados en {destino}")
