"""Génère les icônes de l'application à partir du logo fourni (resources/logo_source.png) :
icon.png (512 px, aussi utilisé dans l'en-tête), icon.ico (Windows, 16 à 256 px) et
icon.icns (macOS, jusqu'à 1024 px).

Le logo est posé sur une tuile blanche aux coins arrondis, identique sur toutes les
plateformes. Aux petites tailles (16 à 32 px), le cercle extérieur et l'étoile, trop fins,
deviennent illisibles : on utilise un recadrage serré sur l'anticorps et la loupe.
Utilisation : python packaging/make_icon.py   (Pillow requis)"""
import os
from PIL import Image, ImageDraw, ImageOps

HERE = os.path.dirname(os.path.abspath(__file__))
RES = os.path.join(HERE, "..", "resources")


def load_logo():
    im = Image.open(os.path.join(RES, "logo_source.png")).convert("RGB")
    # fond presque blanc (253,253,253) -> blanc pur, pour se fondre dans la tuile
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            if r > 246 and g > 246 and b > 246:
                px[x, y] = (255, 255, 255)
    return im


def square_crop(im, box):
    x0, y0, x1, y1 = box
    side = max(x1 - x0, y1 - y0)
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    return im.crop((int(cx - side / 2), int(cy - side / 2), int(cx + side / 2), int(cy + side / 2)))


def tile(art, size, inset):
    """Tuile blanche arrondie (rayon 22 %), liseré gris très léger pour les fonds sombres."""
    S = size * 4                                   # sur-échantillonnage pour des bords nets
    base = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    mask = Image.new("L", (S, S), 0)
    r = int(S * 0.2237)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), r, fill=255)
    white = Image.new("RGBA", (S, S), (255, 255, 255, 255))
    base.paste(white, (0, 0), mask)
    ImageDraw.Draw(base).rounded_rectangle((1, 1, S - 2, S - 2), r, outline=(0, 0, 0, 36), width=max(2, S // 256))
    inner = int(S * (1 - 2 * inset))
    a = art.resize((inner, inner), Image.LANCZOS)
    base.paste(a, ((S - inner) // 2, (S - inner) // 2))
    return base.resize((size, size), Image.LANCZOS)


def main():
    logo = load_logo()
    full = square_crop(logo, (74, 30, 1191, 1139))        # logo entier
    tight = square_crop(logo, (250, 215, 1195, 985))      # anticorps + loupe
    sizes = {}
    for s in (16, 24, 32, 48, 64, 128, 256, 512, 1024):
        sizes[s] = tile(tight if s <= 32 else full, s, 0.06 if s <= 32 else 0.08)
    sizes[512].save(os.path.join(RES, "icon.png"))
    # .ico : chaque taille fournie explicitement (petites tailles = recadrage serré)
    ico_sizes = [16, 24, 32, 48, 64, 128, 256]
    sizes[256].save(os.path.join(RES, "icon.ico"), sizes=[(s, s) for s in ico_sizes],
                    append_images=[sizes[s] for s in ico_sizes if s != 256])
    sizes[1024].save(os.path.join(RES, "icon.icns"))
    sizes[1024].save(os.path.join(RES, "icon_1024.png"))
    print("icônes générées : icon.png (512), icon.ico (16-256), icon.icns (1024)")


if __name__ == "__main__":
    main()
