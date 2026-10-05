"""Affiche la version de glibc la plus récente exigée par les binaires d'un dossier."""
import os, re, subprocess, sys
best = (2, 17)
for d, _, fs in os.walk(sys.argv[1]):
    for f in fs:
        p = os.path.join(d, f)
        if os.path.islink(p):
            continue
        with open(p, "rb") as fh:
            if fh.read(4) != b"\x7fELF":
                continue
        out = subprocess.run(["objdump", "-T", p], capture_output=True, text=True).stdout
        for v in re.findall(r"GLIBC_(\d+)\.(\d+)", out):
            best = max(best, (int(v[0]), int(v[1])))
print(f"{best[0]}.{best[1]}")
