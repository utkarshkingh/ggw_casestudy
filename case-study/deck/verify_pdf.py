"""Ground-truth layout check: read the RENDERED pdf, not the source geometry.

audit.py works from the pptx and therefore trusts the row heights and text-box
sizes we asked for. PowerPoint and LibreOffice both grow a row or a text box to
fit its content, so a table whose header wraps silently pushes everything below
it. This reads the glyph boxes back out of the pdf, where that has already
happened, and reports words that physically overlap.
"""
import subprocess
import sys
import re
from collections import defaultdict

PT = 72.0


def words(pdf):
    xml = subprocess.run(["pdftotext", "-bbox", pdf, "-"],
                         capture_output=True, text=True, check=True).stdout
    pages, cur = [], None
    for line in xml.splitlines():
        if "<page " in line:
            cur = []
            pages.append(cur)
        m = re.search(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', line)
        if m and cur is not None:
            x0, y0, x1, y1 = (float(m.group(i)) / PT for i in range(1, 5))
            cur.append((x0, y0, x1, y1, m.group(5)))
    return pages


def overlaps(a, b):
    ox = min(a[2], b[2]) - max(a[0], b[0])
    oy = min(a[3], b[3]) - max(a[1], b[1])
    if ox <= 0 or oy <= 0:
        return 0.0
    area = ox * oy
    smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
    return area / smaller if smaller else 0.0


def main(pdf):
    issues = 0
    for pi, ws in enumerate(words(pdf), 1):
        # bucket by row band so we only compare words that could collide
        bands = defaultdict(list)
        for w in ws:
            for band in range(int(w[1] * 4), int(w[3] * 4) + 1):
                bands[band].append(w)
        seen = set()
        for band, group in bands.items():
            for i in range(len(group)):
                for j in range(i + 1, len(group)):
                    a, b = group[i], group[j]
                    key = tuple(sorted([id(a), id(b)]))
                    if key in seen:
                        continue
                    seen.add(key)
                    frac = overlaps(a, b)
                    if frac > 0.20:
                        print(f"P{pi} TEXT-COLLISION {frac*100:.0f}%  "
                              f"\"{a[4]}\" @({a[0]:.2f},{a[1]:.2f})  x  \"{b[4]}\" @({b[0]:.2f},{b[1]:.2f})")
                        issues += 1
        for w in ws:
            if w[0] < -0.01 or w[2] > 13.333 + 0.01 or w[1] < -0.01 or w[3] > 7.5 + 0.01:
                print(f"P{pi} OFF-PAGE \"{w[4]}\" [{w[0]:.2f},{w[1]:.2f},{w[2]:.2f},{w[3]:.2f}]")
                issues += 1
    print(f"\nRENDERED ISSUES: {issues}")
    return issues


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1]) else 0)
