"""Génère les icônes de l'application (1.3.1) : une loupe posée sur une double hélice d'ADN,
aux couleurs de la coloration hématoxyline-éosine de l'interface.
Sorties : resources/icon.png (512 px), icon.ico (Windows), icon.icns (macOS).
Utilisation : python packaging/make_icon.py   (PySide6 et Pillow requis)"""
import math
import os
import sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (QColor, QGuiApplication, QImage, QPainter, QPainterPath, QPen,
                           QRadialGradient, QLinearGradient)

HEMA, EOSIN, INK, SLIDE = QColor("#3A3D8F"), QColor("#C8467A"), QColor("#1C1D33"), QColor("#F5F4F8")


def helix(p, S, cx, y0, y1, amp, turns, width, rung_w, alpha=255):
    """Double hélice verticale : brins sinusoïdaux déphasés de π, barreaux (paires de bases)
    entre les deux ; les segments « derrière » sont plus pâles pour donner la profondeur."""
    n = 240
    pts = []
    for i in range(n + 1):
        t = i / n
        y = y0 + (y1 - y0) * t
        a = 2 * math.pi * turns * t
        pts.append((y, a))
    # barreaux
    rungs = int(turns * 10)
    for k in range(1, rungs):
        t = k / rungs
        y = y0 + (y1 - y0) * t
        a = 2 * math.pi * turns * t
        x1, x2 = cx + amp * math.sin(a), cx + amp * math.sin(a + math.pi)
        if abs(x1 - x2) < amp * 0.25:
            continue
        c = QColor("#9A9BC4"); c.setAlpha(alpha)
        p.setPen(QPen(c, rung_w, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(QPointF(x1, y), QPointF(x2, y))
    # brins : d'abord les parties arrière (cos < 0), puis les parties avant
    for front in (False, True):
        for phase, col in ((0.0, HEMA), (math.pi, EOSIN)):
            for i in range(n):
                (ya, aa), (yb, ab) = pts[i], pts[i + 1]
                depth = math.cos(aa + phase)
                if (depth >= 0) != front:
                    continue
                c = QColor(col)
                if not front:
                    c = QColor(int(col.red() * .55 + 255 * .45), int(col.green() * .55 + 255 * .45),
                               int(col.blue() * .55 + 255 * .45))
                c.setAlpha(alpha)
                p.setPen(QPen(c, width, Qt.SolidLine, Qt.RoundCap))
                p.drawLine(QPointF(cx + amp * math.sin(aa + phase), ya),
                           QPointF(cx + amp * math.sin(ab + phase), yb))


def draw(S):
    img = QImage(S, S, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    m = S * .03
    bg = QPainterPath(); bg.addRoundedRect(QRectF(m, m, S - 2 * m, S - 2 * m), S * .21, S * .21)
    p.fillPath(bg, SLIDE)
    p.setClipPath(bg)
    # hélice entière (taille normale)
    hx = S * .40
    helix(p, S, hx, S * .08, S * .92, S * .12, 1.6, S * .045, S * .018)
    # loupe : lentille qui grossit l'hélice
    lc, lr = QPointF(S * .43, S * .43), S * .265
    lens = QPainterPath(); lens.addEllipse(lc, lr, lr)
    p.save()
    p.setClipPath(lens)
    g = QRadialGradient(lc + QPointF(-lr * .3, -lr * .3), lr * 1.3)
    g.setColorAt(0, QColor("#FFFFFF")); g.setColorAt(1, QColor("#E3E4F3"))
    p.fillPath(lens, g)
    zoom = 1.75
    p.translate(lc); p.scale(zoom, zoom); p.translate(-lc)
    helix(p, S, hx, S * .08, S * .92, S * .12, 1.6, S * .045, S * .018)
    p.restore()
    # manche
    ang = math.radians(45)
    a = lc + QPointF(math.cos(ang) * (lr + S * .02), math.sin(ang) * (lr + S * .02))
    b = QPointF(S * .86, S * .86)
    grad = QLinearGradient(a, b); grad.setColorAt(0, INK); grad.setColorAt(1, QColor("#33355A"))
    p.setPen(QPen(grad, S * .095, Qt.SolidLine, Qt.RoundCap))
    p.drawLine(a, b)
    # monture
    p.setPen(QPen(INK, S * .05)); p.setBrush(Qt.NoBrush)
    p.drawEllipse(lc, lr, lr)
    # reflet
    p.setPen(QPen(QColor(255, 255, 255, 200), S * .022, Qt.SolidLine, Qt.RoundCap))
    p.drawArc(QRectF(lc.x() - lr * .72, lc.y() - lr * .72, lr * 1.44, lr * 1.44), 100 * 16, 55 * 16)
    p.end()
    return img


if __name__ == "__main__":
    app = QGuiApplication(sys.argv)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "resources")
    big = draw(1024)
    big.scaled(512, 512, Qt.KeepAspectRatio, Qt.SmoothTransformation).save(os.path.join(out, "icon.png"))
    tmp = os.path.join(out, "_icon1024.png")
    big.save(tmp)
    from PIL import Image
    im = Image.open(tmp).convert("RGBA")
    im.save(os.path.join(out, "icon.ico"),
            sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    im.save(os.path.join(out, "icon.icns"))
    os.remove(tmp)
    print("icônes générées : icon.png, icon.ico, icon.icns")
