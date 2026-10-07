from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

TAMANOS = (16, 24, 32, 48, 64, 128, 256)
BASE = 1024
SUPERMUESTREO = 4

AZUL = (28, 96, 200, 255)
BLANCO = (255, 255, 255, 255)
PLIEGUE = (211, 219, 232, 255)
AMARILLO = (255, 193, 7, 255)
BORDE_RAYO = (176, 122, 0, 255)


def dibujar_lienzo() -> Image.Image:
    escala = SUPERMUESTREO
    lienzo = Image.new("RGBA", (BASE * escala, BASE * escala), (0, 0, 0, 0))
    dibujo = ImageDraw.Draw(lienzo)

    def punto(x: int, y: int) -> tuple[int, int]:
        return (x * escala, y * escala)

    def rect(x1: int, y1: int, x2: int, y2: int) -> list[int]:
        return [c * escala for c in (x1, y1, x2, y2)]

    dibujo.rounded_rectangle(rect(56, 56, 968, 968), radius=196 * escala, fill=AZUL)

    x1, y1, x2, y2 = 297, 214, 727, 790
    doblez = 150
    dibujo.rounded_rectangle(rect(x1, y1, x2, y2), radius=40 * escala, fill=BLANCO)
    dibujo.polygon(
        [punto(x2 - doblez, y1), punto(x2, y1), punto(x2, y1 + doblez)],
        fill=AZUL,
    )
    dibujo.polygon(
        [punto(x2 - doblez, y1), punto(x2, y1 + doblez), punto(x2 - doblez, y1 + doblez)],
        fill=PLIEGUE,
    )

    rayo = [
        (578, 306),
        (414, 552),
        (506, 552),
        (444, 756),
        (634, 496),
        (540, 496),
    ]
    puntos = [punto(x, y) for x, y in rayo]
    dibujo.line(puntos + [puntos[0]], fill=BORDE_RAYO, width=14 * escala, joint="curve")
    dibujo.polygon(puntos, fill=AMARILLO)

    return lienzo.resize((BASE, BASE), Image.Resampling.LANCZOS)


def reducir(imagen: Image.Image, tamano: int) -> Image.Image:
    actual = imagen
    while actual.size[0] > tamano * 2:
        mitad = actual.size[0] // 2
        actual = actual.resize((mitad, mitad), Image.Resampling.LANCZOS)
    return actual.resize((tamano, tamano), Image.Resampling.LANCZOS)


def main() -> None:
    maestra = dibujar_lienzo()
    icono = Path(__file__).resolve().parent / "icono.ico"
    marcos = [reducir(maestra, tamano) for tamano in TAMANOS]
    marcos[-1].save(
        icono,
        format="ICO",
        sizes=[(tamano, tamano) for tamano in TAMANOS],
        append_images=marcos[:-1],
    )
    vista = Path(__file__).resolve().parent / "vista_previa_icono.png"
    reducir(maestra, 256).save(vista)
    print(f"Icono creado: {icono}")
    print(f"Vista previa: {vista}")


if __name__ == "__main__":
    main()
