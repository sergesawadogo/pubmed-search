"""Génère resources/icon.png et resources/icon.ico : un polynucléaire neutrophile
(cytoplasme éosinophile, noyau trilobé coloré à l'hématoxyline)."""
import os, sys
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPainterPath, QPen, QRadialGradient

app = QGuiApplication(sys.argv)
S = 512
img = QImage(S, S, QImage.Format_ARGB32)
img.fill(Qt.transparent)
p = QPainter(img)
p.setRenderHint(QPainter.Antialiasing)
# fond : lame
bg = QPainterPath(); bg.addRoundedRect(QRectF(16, 16, S - 32, S - 32), 110, 110)
p.fillPath(bg, QColor("#F5F4F8"))
# cytoplasme
g = QRadialGradient(QPointF(S * .45, S * .42), S * .42)
g.setColorAt(0, QColor("#F2B6CC")); g.setColorAt(1, QColor("#C8467A"))
p.setBrush(g); p.setPen(Qt.NoPen)
p.drawEllipse(QPointF(S / 2, S / 2), S * .36, S * .36)
# noyau trilobé (lobes reliés par des ponts de chromatine)
lobes = [(.36, .40, .105), (.55, .33, .095), (.62, .56, .11)]
pen = QPen(QColor("#3A3D8F"), S * .045, Qt.SolidLine, Qt.RoundCap)
p.setPen(pen)
for (x1, y1, _), (x2, y2, _) in zip(lobes, lobes[1:]):
    p.drawLine(QPointF(S * x1, S * y1), QPointF(S * x2, S * y2))
p.setPen(Qt.NoPen); p.setBrush(QColor("#3A3D8F"))
for x, y, r in lobes:
    p.drawEllipse(QPointF(S * x, S * y), S * r, S * r)
p.end()
out = os.path.join(os.path.dirname(__file__), "..", "resources")
img.save(os.path.join(out, "icon.png"))
from PIL import Image
Image.open(os.path.join(out, "icon.png")).save(os.path.join(out, "icon.ico"),
    sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("icônes générées")
